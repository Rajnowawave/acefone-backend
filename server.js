require("dotenv").config();

const express = require("express");
const cors = require("cors");
const admin = require("firebase-admin");
const fs = require("fs");
const ExcelJS = require("exceljs");

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// ─── FIX #1 (v2): FIREBASE_SERVICE_ACCOUNT loading ───────────────────────────
// Manually pasting the whole service-account JSON into a .env value is very
// error-prone — copy/paste from Word/Notion/some editors silently converts
// straight quotes (") into "smart quotes" (" "), which breaks JSON.parse with
// exactly the error you saw ("Expected property name or '}'"). So this now
// supports TWO ways to provide credentials, and prefers the safer one:
//
//   Option A (recommended): FIREBASE_SERVICE_ACCOUNT_PATH=./serviceAccountKey.json
//     Download the key file from Firebase Console → Project Settings →
//     Service Accounts → Generate New Private Key, save it next to server.js
//     as serviceAccountKey.json (add it to .gitignore!), and just point to it.
//     No copy/paste of JSON into .env at all — zero chance of quote corruption.
//     On Render: use the "Secret Files" feature to upload this same file at
//     the same path, then set FIREBASE_SERVICE_ACCOUNT_PATH to that path.
//
//   Option B: FIREBASE_SERVICE_ACCOUNT=<single-line JSON> in .env (old way,
//     kept for backwards compatibility). If this keeps failing, switch to
//     Option A instead — it's simpler and avoids this whole class of bug.
let serviceAccount;

if (process.env.FIREBASE_SERVICE_ACCOUNT_PATH) {
  const keyPath = process.env.FIREBASE_SERVICE_ACCOUNT_PATH;
  try {
    const raw = fs.readFileSync(keyPath, "utf8");
    serviceAccount = JSON.parse(raw);
    console.log("✅ Firebase credentials loaded from file:", keyPath);
  } catch (e) {
    console.error(`❌ Could not read/parse FIREBASE_SERVICE_ACCOUNT_PATH ("${keyPath}"):`, e.message);
    process.exit(1);
  }
} else if (process.env.FIREBASE_SERVICE_ACCOUNT) {
  const raw = process.env.FIREBASE_SERVICE_ACCOUNT;
  try {
    serviceAccount = JSON.parse(raw);
  } catch (e) {
    // Print a diagnostic hexdump of the first ~20 chars so you can SEE the
    // bad character (e.g. a curly quote 0x201C instead of a straight " 0x22).
    const preview = raw.slice(0, 20);
    const codes = [...preview].map((c) => `${c === '"' ? '"' : c}(U+${c.charCodeAt(0).toString(16).padStart(4, "0")})`).join(" ");
    console.error(
      "❌ FIREBASE_SERVICE_ACCOUNT is set but is not valid JSON.\n" +
        "   Parse error:", e.message, "\n" +
        "   First 20 characters + Unicode code points (look for anything that\n" +
        "   is NOT U+0022, the straight double quote, where a quote should be):\n" +
        "   " + codes + "\n\n" +
        "   Strong recommendation: switch to FIREBASE_SERVICE_ACCOUNT_PATH instead —\n" +
        "   save the downloaded key as serviceAccountKey.json next to server.js and set:\n" +
        "   FIREBASE_SERVICE_ACCOUNT_PATH=./serviceAccountKey.json\n" +
        "   This avoids manual JSON-in-.env copy/paste entirely."
    );
    process.exit(1);
  }
} else {
  console.error(
    "❌ No Firebase credentials found.\n" +
      "   Set ONE of these in your .env:\n" +
      "   FIREBASE_SERVICE_ACCOUNT_PATH=./serviceAccountKey.json   (recommended)\n" +
      '   FIREBASE_SERVICE_ACCOUNT={"type":"service_account", ... }   (single line)\n' +
      "   Restart the server after editing .env — it is only read on startup."
  );
  process.exit(1);
}

admin.initializeApp({ credential: admin.credential.cert(serviceAccount) });
const db = admin.firestore();

const PORT = Number(process.env.PORT || 5000);
const ACEFONE_CALLER_ID = process.env.ACEFONE_CALLER_ID || "+918062491504";
const ACEFONE_USER_ID = process.env.ACEFONE_USER_ID || "219085";

// ─── FIX #2 (v2): switched from Puppeteer login-scraping to Acefone's
// OFFICIAL Click-to-Call REST API. This removes the entire fragile
// browser-automation/session.json/CSRF-cookie system — that was the actual
// root cause of the intermittent 401s (Puppeteer failing to log in / stale
// cookies / CSRF token expiring). The official API just needs a Bearer token.
//
// Get/renew this token from: Acefone Console → API Connect → API Tokens.
// Per Acefone's docs, portal-generated tokens don't expire, but if you
// generated this one via their "Generate a Token" API, it may have an
// expiry (check the `exp` field by decoding the JWT) — Acefone provides a
// "Refresh a Token" API for that case.
const ACEFONE_TOKEN = process.env.ACEFONE_TOKEN;
if (!ACEFONE_TOKEN) {
  console.error(
    "❌ ACEFONE_TOKEN is not set in .env — Click-to-Call will fail.\n" +
      "   Get it from: Acefone Console → API Connect → API Tokens → Generate Token."
  );
}

const AGENTS = [
  { id: "0502190850001", name: "Neelam", number: "919251651958" },
  { id: "0502190850002", name: "Bhavika", number: "919251651956" },
  { id: "0502190850003", name: "Tushar Bhandari", number: "917976630010" },
  { id: "0502190850004", name: "Vikash Singhvi", number: "919509805201" },
  { id: "0502190850005", name: "Amit Sharma", number: "918094121221" },
];

// ─── Helpers ──────────────────────────────────────────────────────────────────
const num = (v) => {
  const n = Number(v);
  return isNaN(n) ? 0 : n;
};

const clean = (v) => {
  if (v === undefined || v === null) return "";
  const s = String(v).trim();
  if (/^\$[a-z_]+$/i.test(s)) return "";
  return s;
};

const pick = (d, ...keys) => {
  for (const k of keys) {
    const v = clean(d[k]);
    if (v) return v;
    const v2 = clean(d["$" + k]);
    if (v2) return v2;
  }
  return "";
};

function detectDirection(d) {
  const raw = pick(d, "direction").toLowerCase();
  if (raw.includes("inbound")) return "inbound";
  if (raw.includes("outbound")) return "outbound";
  if (raw.includes("clicktocall")) return "outbound"; // ✅ Click-to-call is outbound
  return "inbound";
}

// ─── Acefone Click-to-Call (official API — no browser automation needed) ─────
async function makeAcefoneCall(customerNumber, agentId) {
  const digits = String(customerNumber).replace(/\D/g, "").slice(-10);
  const phone = "91" + digits;
  const agent = AGENTS.find((a) => a.id === agentId);

  if (!agent) {
    throw new Error("Invalid agent selected");
  }

  console.log(`📞 Calling ${phone} via agent: ${agent.name} (${agent.id})`);

  const response = await fetch("https://api.acefone.in/v1/click_to_call", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${ACEFONE_TOKEN}`,
    },
    body: JSON.stringify({
      agent_number: agent.id, // Acefone Agent ID (050XXXXXXX)
      destination_number: phone,
      caller_id: ACEFONE_CALLER_ID.replace(/^\+/, ""),
      async: 1,
    }),
  });

  return { response, agent };
}

// ✅ WEBHOOK HANDLER FUNCTION
async function handleWebhook(d, res) {
  try {
    console.log("🔔 Webhook Data:\n", JSON.stringify(d, null, 2));

    const direction = detectDirection(d);
    const uuid = pick(d, "uuid", "call_id");
    const callTo = pick(d, "call_to_number");
    const callerNum = pick(d, "caller_id_number");
    const callStatus = pick(d, "call_status") || "completed";

    let recordingUrl = pick(d, "recording_url");
    console.log(`🎙 Recording URL: ${recordingUrl ? "✅ Found" : "❌ Not found"}`);

    const billsec = num(pick(d, "billsec"));
    const duration = num(pick(d, "duration")) || billsec;

    let agentName = "";
    let agentNumber = "";
    let agentId = "";

    const answeredAgent = d.answered_agent;
    if (answeredAgent && typeof answeredAgent === "object") {
      agentName = answeredAgent.name || "";
      agentNumber = answeredAgent.number || answeredAgent.agent_number || "";
      agentId = answeredAgent.id || "";
    } else {
      agentName = pick(d, "answered_agent_name");
      agentNumber = pick(d, "answered_agent_number");
      agentId = pick(d, "answered_agent");
    }

    const customerNoWithPrefix = pick(
      d,
      "customer_no_with_prefix",
      "customer_number_with_prefix",
      "customer_no_with_prefix "
    );

    console.log(
      `📌 ${direction.toUpperCase()} | UUID: ${uuid} | Status: ${callStatus} | Duration: ${billsec}s | Agent: ${agentName} | Recording: ${recordingUrl ? "✅ Available" : "❌ None"}`
    );

    let clientNumber = "";
    let didNumber = "";

    if (direction === "inbound") {
      clientNumber =
        pick(d, "client_number") ||
        pick(d, "caller_id_num") ||
        pick(d, "customer_no_with_prefix") ||
        "";

      didNumber = pick(d, "did_number") || pick(d, "call_to_number") || "";

      clientNumber = String(clientNumber).replace(/^\+91/, "");
      didNumber = String(didNumber).replace(/^\+91/, "");
    }

    const doc = {
      direction,
      uuid,
      call_id: pick(d, "call_id") || uuid,
      call_to_number: callTo,
      caller_id_number: callerNum,
      customer_no_with_prefix: customerNoWithPrefix,
      start_stamp: pick(d, "start_stamp"),
      answer_stamp: pick(d, "answer_stamp"),
      end_stamp: pick(d, "end_stamp"),
      billsec,
      duration,
      outbound_sec: num(pick(d, "outbound_sec")),
      agent_ring_time: num(pick(d, "agent_ring_time")),
      agent_transfer_ring_time: num(pick(d, "agent_transfer_ring_time")),
      customer_ring_time: num(pick(d, "customer_ring_time")),
      answered_agent: agentId,
      answered_agent_name: agentName,
      answered_agent_number: agentNumber,
      missed_agent: pick(d, "missed_agent"),
      call_status: callStatus,
      call_connected: pick(d, "call_connected"),
      call_flow: d.call_flow || [],
      digits_dialed: pick(d, "digits_dialed"),
      billing_circle: d.billing_circle || {},
      campaign_name: pick(d, "campaign_name"),
      campaign_id: pick(d, "campaign_id"),
      broadcast_lead_fields: pick(d, "broadcast_lead_fields"),
      recording_url: recordingUrl || "",
      aws_call_recording_identifier: pick(d, "aws_call_recording_identifier"),
      reason_key: pick(d, "reason_key"),
      hangup_cause_description: pick(d, "hangup_cause_description"),
      hangup_cause_code: pick(d, "hangup_cause_code"),
      hangup_cause_key: pick(d, "hangup_cause_key"),
      client_number: clientNumber,
      did_number: didNumber,
      ref_id: pick(d, "ref_id"),
      raw: d,
      updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    };

    if (direction === "outbound") {
      const existing = await db
        .collection("calls")
        .where("direction", "==", "outbound")
        .where("call_status", "==", "initiated")
        .orderBy("createdAt", "desc")
        .limit(1)
        .get();

      if (!existing.empty) {
        const prev = existing.docs[0].data();

        await existing.docs[0].ref.update({
          ...doc,
          createdAt: prev.createdAt,
          leadId: prev.leadId || "",
          name: prev.name || "",
          answered_agent_name: agentName || prev.answered_agent_name || "",
          answered_agent_number: agentNumber || prev.answered_agent_number || "",
          answered_agent: agentId || prev.answered_agent || "",
          updatedAt: admin.firestore.FieldValue.serverTimestamp(),
        });

        console.log(
          `✅ OUTBOUND updated: ${existing.docs[0].id} | Recording: ${recordingUrl ? "✅" : "❌"}`
        );

        return res.status(200).send("OK");
      }
    }

    const newCall = await db.collection("calls").add({
      ...doc,
      createdAt: admin.firestore.FieldValue.serverTimestamp(),
      name: "",
      leadId: "",
    });

    console.log(`✅ INBOUND created: ${newCall.id} | Recording: ${recordingUrl ? "✅" : "❌"}`);
    return res.status(200).send("OK");
  } catch (e) {
    console.error("❌ Webhook error:", e);
    return res.status(500).send("ERROR");
  }
}

// ✅ RECORDING PROXY (For authenticated playback)
app.get("/recording/:id", async (req, res) => {
  try {
    const doc = await db.collection("calls").doc(req.params.id).get();
    if (!doc.exists) return res.status(404).send("Call not found");

    const data = doc.data();
    const url = data.recording_url;
    if (!url) return res.status(404).send("No recording available");

    const response = await fetch(url);
    if (!response.ok) return res.status(404).send("Recording not accessible");

    res.setHeader("Content-Type", "audio/mpeg");
    res.setHeader("Accept-Ranges", "bytes");
    res.setHeader("Cache-Control", "public, max-age=3600");

    response.body.pipe(res);
  } catch (e) {
    console.error("Recording proxy error:", e);
    return res.status(500).send("Error fetching recording");
  }
});

// ─── CONTACTS ─────────────────────────────────────────────────────────────────
app.get("/contacts", async (req, res) => {
  try {
    const snap = await db.collection("contacts").get();
    const contacts = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
    res.json(contacts);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ─── FIX #3: Dialer.jsx calls POST /contacts, PUT /contacts/:id and
// DELETE /contacts/:id, but those routes never existed on the backend —
// so "Add Contact" / "Edit Contact" / "Delete Contact" silently failed
// (the frontend swallows the fetch error in a try/catch). Added below.
app.post("/contacts", async (req, res) => {
  try {
    const { name, number, email, company, color } = req.body;
    if (!name || !number) {
      return res.status(400).json({ error: "Name and number are required" });
    }
    const ref = await db.collection("contacts").add({
      name: name.trim(),
      number: String(number).replace(/\D/g, ""),
      email: email || "",
      company: company || "",
      color: color || "",
      status: "offline",
      lastCall: null,
      createdAt: admin.firestore.FieldValue.serverTimestamp(),
    });
    return res.json({ id: ref.id, success: true });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

app.put("/contacts/:id", async (req, res) => {
  try {
    const { name, number, email, company, color } = req.body;
    const update = { updatedAt: admin.firestore.FieldValue.serverTimestamp() };
    if (name !== undefined) update.name = name.trim();
    if (number !== undefined) update.number = String(number).replace(/\D/g, "");
    if (email !== undefined) update.email = email;
    if (company !== undefined) update.company = company;
    if (color !== undefined) update.color = color;

    await db.collection("contacts").doc(req.params.id).update(update);
    return res.json({ success: true });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

app.delete("/contacts/:id", async (req, res) => {
  try {
    await db.collection("contacts").doc(req.params.id).delete();
    return res.json({ success: true });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

// ✅ Get recording info (URL only)
app.get("/recording-info/:id", async (req, res) => {
  try {
    const doc = await db.collection("calls").doc(req.params.id).get();
    if (!doc.exists) return res.status(404).json({ error: "Call not found" });

    const data = doc.data();
    const url = data.recording_url;
    if (!url) return res.status(404).json({ error: "No recording available" });

    return res.json({
      url,
      callId: data.call_id,
      uuid: data.uuid,
      duration: data.billsec,
      status: data.call_status,
    });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

// ─── Routes ───────────────────────────────────────────────────────────────────
app.get("/", (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html>
    <head>
      <title>Acefone Call Portal</title>
      <style>
        body { font-family: Arial; margin: 40px; background: #f5f5f5; }
        .container { background: white; padding: 30px; border-radius: 8px; max-width: 600px; margin: 0 auto; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #2563eb; margin: 0 0 20px 0; }
        .status { color: #16a34a; font-size: 18px; font-weight: bold; }
        .endpoint { background: #f1f5f9; padding: 12px; border-radius: 4px; margin: 10px 0; font-family: monospace; word-break: break-all; }
        .label { color: #64748b; font-size: 12px; margin-bottom: 4px; }
        .url { color: #0ea5e9; }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>📡 Acefone Call Portal</h1>
        <div class="status">✅ Backend Running</div>
        <p>Server is active and ready to receive webhooks.</p>
        <div class="endpoint">
          <div class="label">Webhook Endpoint:</div>
          <div class="url">${req.protocol}://${req.get("host")}/webhook</div>
        </div>
        <div class="endpoint">
          <div class="label">Time:</div>
          <div class="url">${new Date().toLocaleString("en-IN")}</div>
        </div>
      </div>
    </body>
    </html>
  `);
});

app.get("/webhook", async (req, res) => {
  return await handleWebhook(req.query || {}, res);
});

app.post("/webhook", async (req, res) => {
  return await handleWebhook(req.body || {}, res);
});

app.get("/agents", (req, res) => res.json(AGENTS));

// ─── LEADS ────────────────────────────────────────────────────────────────────
app.get("/leads", async (req, res) => {
  try {
    const snap = await db.collection("leads").orderBy("createdAt", "desc").get();
    const leads = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
    return res.json(leads);
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

app.post("/leads", async (req, res) => {
  try {
    const { date, name, contact, city, propertyType, remarks, status, calledBy } = req.body;

    if (!name || !contact) {
      return res.status(400).json({ error: "Name and contact are required" });
    }

    const leadData = {
      date: date || new Date().toISOString().split("T")[0],
      name: name.trim(),
      contact: contact.trim(),
      city: city?.trim() || "",
      propertyType: propertyType || "",
      remarks: remarks || "",
      status: status || "Not Called",
      calledBy: calledBy || "",
      followUps: [],
      callCount: 0,
      createdAt: admin.firestore.FieldValue.serverTimestamp(),
      updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    };

    const ref = await db.collection("leads").add(leadData);
    return res.json({ id: ref.id, success: true });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

app.post("/leads/bulk", async (req, res) => {
  try {
    const leads = req.body;
    if (!Array.isArray(leads) || leads.length === 0) {
      return res.status(400).json({ error: "Invalid data format" });
    }

    const batch = db.batch();
    let count = 0;

    leads.forEach((lead) => {
      if (lead.name && lead.contact) {
        const ref = db.collection("leads").doc();
        batch.set(ref, {
          date: lead.date || new Date().toISOString().split("T")[0],
          name: lead.name.trim(),
          contact: lead.contact.trim(),
          city: lead.city?.trim() || "",
          propertyType: lead.propertyType || "",
          remarks: lead.remarks || "",
          status: lead.status || "Not Called",
          calledBy: lead.calledBy || "",
          followUps: [],
          callCount: 0,
          createdAt: admin.firestore.FieldValue.serverTimestamp(),
          updatedAt: admin.firestore.FieldValue.serverTimestamp(),
        });
        count++;
      }
    });

    await batch.commit();
    return res.json({ success: true, count });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

app.put("/leads/:id", async (req, res) => {
  try {
    const { name, contact, city, propertyType, remarks, status, calledBy, followUps } = req.body;

    const updateData = { updatedAt: admin.firestore.FieldValue.serverTimestamp() };

    if (name !== undefined) updateData.name = name.trim();
    if (contact !== undefined) updateData.contact = contact.trim();
    if (city !== undefined) updateData.city = city.trim();
    if (propertyType !== undefined) updateData.propertyType = propertyType;
    if (remarks !== undefined) updateData.remarks = remarks;
    if (status !== undefined) updateData.status = status;
    if (calledBy !== undefined) updateData.calledBy = calledBy;
    if (followUps !== undefined) updateData.followUps = followUps;

    if (status && status !== "Not Called") {
      const doc = await db.collection("leads").doc(req.params.id).get();
      if (doc.exists && (!doc.data().status || doc.data().status === "Not Called")) {
        updateData.callCount = admin.firestore.FieldValue.increment(1);
        updateData.lastCalledAt = admin.firestore.FieldValue.serverTimestamp();
      }
    }

    await db.collection("leads").doc(req.params.id).update(updateData);
    return res.json({ success: true });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

app.delete("/leads/:id", async (req, res) => {
  try {
    await db.collection("leads").doc(req.params.id).delete();
    return res.json({ success: true });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

app.get("/leads/:id/history", async (req, res) => {
  try {
    const leadDoc = await db.collection("leads").doc(req.params.id).get();
    if (!leadDoc.exists) return res.status(404).json({ error: "Lead not found" });

    const leadData = leadDoc.data();
    const contact = leadData.contact;

    const callsSnap = await db.collection("calls").orderBy("createdAt", "desc").limit(100).get();

    const history = callsSnap.docs
      .filter((doc) => {
        const data = doc.data();
        const callNumber = (data.call_to_number || data.caller_id_number || "").replace(/\D/g, "").slice(-10);
        const leadNumber = contact.replace(/\D/g, "").slice(-10);
        return callNumber === leadNumber;
      })
      .map((doc) => {
        const data = doc.data();
        return {
          timestamp: data.createdAt?._seconds
            ? new Date(data.createdAt._seconds * 1000).toISOString()
            : new Date(data.createdAt).toISOString(),
          status: data.call_status || "Unknown",
          calledBy: data.answered_agent_name || "Unknown",
          duration: data.billsec || 0,
          remarks: data._remark || "",
        };
      });

    return res.json(history);
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

app.get("/leads/export", async (req, res) => {
  try {
    const { date, status } = req.query;

    let query = db.collection("leads").orderBy("createdAt", "desc");
    const snap = await query.get();
    let leads = snap.docs.map((d) => ({ id: d.id, ...d.data() }));

    if (date) leads = leads.filter((l) => l.date === date);
    if (status) leads = leads.filter((l) => l.status === status);

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Leads");

    sheet.columns = [
      { header: "Date", key: "date", width: 12 },
      { header: "Name", key: "name", width: 25 },
      { header: "Contact", key: "contact", width: 15 },
      { header: "City", key: "city", width: 20 },
      { header: "Property Type", key: "propertyType", width: 18 },
      { header: "Remarks", key: "remarks", width: 40 },
      { header: "Status", key: "status", width: 15 },
      { header: "Called By", key: "calledBy", width: 20 },
      { header: "Call Count", key: "callCount", width: 12 },
      { header: "Last Called", key: "lastCalledAt", width: 20 },
      { header: "Follow Up 1", key: "followUp1", width: 30 },
      { header: "Follow Up 2", key: "followUp2", width: 30 },
      { header: "Follow Up 3", key: "followUp3", width: 30 },
    ];

    sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" }, size: 11 };
    sheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2C3E6B" } };
    sheet.getRow(1).alignment = { horizontal: "center", vertical: "middle" };
    sheet.getRow(1).height = 25;

    leads.forEach((lead, index) => {
      const row = sheet.addRow({
        date: lead.date || "—",
        name: lead.name || "—",
        contact: lead.contact || "—",
        city: lead.city || "—",
        propertyType: lead.propertyType || "—",
        remarks: lead.remarks || "—",
        status: lead.status || "Not Called",
        calledBy: lead.calledBy || "—",
        callCount: lead.callCount || 0,
        lastCalledAt: lead.lastCalledAt
          ? new Date(
              lead.lastCalledAt._seconds ? lead.lastCalledAt._seconds * 1000 : lead.lastCalledAt
            ).toLocaleString("en-IN")
          : "—",
        followUp1: lead.followUps?.[0]
          ? `${new Date(lead.followUps[0].date).toLocaleDateString("en-IN")} - ${lead.followUps[0].note}`
          : "—",
        followUp2: lead.followUps?.[1]
          ? `${new Date(lead.followUps[1].date).toLocaleDateString("en-IN")} - ${lead.followUps[1].note}`
          : "—",
        followUp3: lead.followUps?.[2]
          ? `${new Date(lead.followUps[2].date).toLocaleDateString("en-IN")} - ${lead.followUps[2].note}`
          : "—",
      });

      if (index % 2 === 0) {
        row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
      }
      if (!lead.status || lead.status === "Not Called") {
        row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFBEB" } };
      }
    });

    sheet.autoFilter = { from: "A1", to: "M1" };
    sheet.views = [{ state: "frozen", ySplit: 1 }];

    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        cell.border = {
          top: { style: "thin", color: { argb: "FFE2E8F0" } },
          left: { style: "thin", color: { argb: "FFE2E8F0" } },
          bottom: { style: "thin", color: { argb: "FFE2E8F0" } },
          right: { style: "thin", color: { argb: "FFE2E8F0" } },
        };
      });
    });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename=leads-${new Date().toISOString().split("T")[0]}.xlsx`);

    await workbook.xlsx.write(res);
    res.end();
  } catch (e) {
    console.error("❌ Excel export error:", e);
    return res.status(500).json({ error: e.message });
  }
});

// ─── CALL ─────────────────────────────────────────────────────────────────────
app.post("/call", async (req, res) => {
  try {
    const { customer, agentId, leadId, name } = req.body;

    if (!customer) return res.status(400).json({ error: "Customer number required" });
    if (!agentId) return res.status(400).json({ error: "Please select an agent" });

    const { response, agent } = await makeAcefoneCall(customer, agentId);
    const json = await response.json().catch(() => ({}));
    console.log("Acefone →", response.status, json);

    const digits = String(customer).replace(/\D/g, "").slice(-10);
    const refId = json.ref_id || "";

    const callDoc = await db.collection("calls").add({
      direction: "outbound",
      call_to_number: "91" + digits,
      caller_id_number: ACEFONE_CALLER_ID,
      answered_agent_name: agent.name,
      answered_agent_number: agent.number,
      answered_agent: agent.id,
      name: name || "",
      leadId: leadId || "",
      call_status: json.success ? "initiated" : "failed",
      ref_id: refId, // used to match the async webhook back to this record
      recording_url: "",
      billsec: 0,
      createdAt: admin.firestore.FieldValue.serverTimestamp(),
      updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    });

    console.log(`✅ Call logged: ${callDoc.id} | Agent: ${agent.name} | ref_id: ${refId}`);

    if (leadId) {
      const lRef = db.collection("leads").doc(leadId);
      const lDoc = await lRef.get();
      if (lDoc.exists) {
        await lRef.update({
          callCount: (lDoc.data().callCount || 0) + 1,
          status: json.success ? "called" : "failed",
          lastCalledAt: admin.firestore.FieldValue.serverTimestamp(),
          lastAgent: agent.name,
        });
      }
    }

    if (json.success) {
      // Acefone's API is asynchronous: this only means the request was
      // accepted, NOT that the call connected. Actual progress arrives via
      // /webhook. The frontend polls /call-status/:id, which reads whatever
      // the webhook has written to this Firestore doc.
      return res.json({ success: true, agent: agent.name, callId: callDoc.id, refId });
    }

    return res.status(response.status || 400).json({ error: json.message || "Call failed", details: json });
  } catch (e) {
    console.error("Call error:", e);
    return res.status(500).json({ error: e.message });
  }
});

// ─── FIX #5: these two routes did not exist before, so the dialer's
// call-status polling and the "End Call" button were silently failing
// (Dialer.jsx already calls them, it just had nothing to talk to). ────────────
app.get("/call-status/:id", async (req, res) => {
  try {
    const doc = await db.collection("calls").doc(req.params.id).get();
    if (!doc.exists) return res.status(404).json({ error: "Call not found" });
    const data = doc.data();
    return res.json({ status: data.call_status || "unknown", ...data });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

app.post("/call/:id/end", async (req, res) => {
  try {
    // Acefone's click-to-call API doesn't expose a public "hang up" endpoint,
    // so this marks the call as completed in our own records. The webhook
    // will still overwrite this with the real final status once Acefone
    // sends it.
    await db.collection("calls").doc(req.params.id).update({
      call_status: "completed",
      updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    });
    return res.json({ success: true });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

// ✅ RECORDING PROXY
app.get("/recording-proxy/:id", async (req, res) => {
  try {
    const doc = await db.collection("calls").doc(req.params.id).get();
    if (!doc.exists) return res.status(404).json({ error: "Call not found" });

    const data = doc.data();
    const url = data.recording_url || "";
    if (!url) return res.status(404).json({ error: "No recording available" });

    return res.json({ url });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

// ─── CALL LOGS ────────────────────────────────────────────────────────────────
app.get("/call-logs", async (req, res) => {
  try {
    const { direction, status, search, customer } = req.query;
    const lim = Math.min(Number(req.query.limit) || 200, 500);

    let ref = db.collection("calls").orderBy("createdAt", "desc").limit(lim);
    if (direction && direction !== "all") ref = ref.where("direction", "==", direction);

    const snap = await ref.get();
    let logs = snap.docs.map((d) => ({ id: d.id, ...d.data() }));

    if (status && status !== "all") {
      logs = logs.filter((c) => (c.call_status || "").toLowerCase() === status.toLowerCase());
    }

    if (customer) {
      const cDigits = customer.replace(/\D/g, "").slice(-10);
      logs = logs.filter((c) =>
        [c.call_to_number, c.caller_id_number, c.customer_no_with_prefix].some(
          (v) => v && String(v).replace(/\D/g, "").includes(cDigits)
        )
      );
    }

    if (search) {
      const q = search.toLowerCase();
      logs = logs.filter((c) =>
        [
          c.call_to_number, c.caller_id_number, c.answered_agent_name,
          c.customer_no_with_prefix, c.campaign_name, c.name, c.uuid, c.call_id,
        ].some((v) => v && String(v).toLowerCase().includes(q))
      );
    }

    return res.json(logs);
  } catch (e) {
    console.error("call-logs error:", e.message);
    return res.status(500).json({ error: e.message });
  }
});

// ✅ EXCEL EXPORT
app.get("/export-excel", async (req, res) => {
  try {
    const { direction, status, agent, dateFrom, dateTo } = req.query;

    let ref = db.collection("calls").orderBy("createdAt", "desc").limit(1000);
    if (direction && direction !== "all") ref = ref.where("direction", "==", direction);

    const snap = await ref.get();
    let calls = snap.docs.map((d) => ({ id: d.id, ...d.data() }));

    if (status && status !== "all") {
      calls = calls.filter((c) => (c.call_status || "").toLowerCase() === status.toLowerCase());
    }
    if (agent && agent !== "all") {
      calls = calls.filter((c) => c.answered_agent_name === agent);
    }
    if (dateFrom) {
      const fromTs = new Date(dateFrom).getTime();
      calls = calls.filter((c) => {
        const ts = c.createdAt?._seconds ? c.createdAt._seconds * 1000 : new Date(c.createdAt).getTime();
        return ts >= fromTs;
      });
    }
    if (dateTo) {
      const toTs = new Date(dateTo).getTime() + 86400000;
      calls = calls.filter((c) => {
        const ts = c.createdAt?._seconds ? c.createdAt._seconds * 1000 : new Date(c.createdAt).getTime();
        return ts <= toTs;
      });
    }

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Call Logs");

    sheet.columns = [
      { header: "Call ID", key: "id", width: 25 },
      { header: "UUID", key: "uuid", width: 35 },
      { header: "Direction", key: "direction", width: 12 },
      { header: "Customer Number", key: "customer", width: 18 },
      { header: "Customer Name", key: "name", width: 20 },
      { header: "Caller ID", key: "callerId", width: 18 },
      { header: "Agent Name", key: "agent", width: 20 },
      { header: "Agent Number", key: "agentNum", width: 18 },
      { header: "Status", key: "status", width: 15 },
      { header: "Duration (sec)", key: "duration", width: 15 },
      { header: "Start Time", key: "startTime", width: 22 },
      { header: "Answer Time", key: "answerTime", width: 22 },
      { header: "End Time", key: "endTime", width: 22 },
      { header: "Recording URL", key: "recording", width: 70 },
      { header: "Campaign", key: "campaign", width: 20 },
      { header: "Hangup Cause", key: "hangup", width: 25 },
    ];

    sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" }, size: 11 };
    sheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2563EB" } };
    sheet.getRow(1).alignment = { horizontal: "center", vertical: "middle" };
    sheet.getRow(1).height = 25;

    calls.forEach((c, index) => {
      const ts = c.createdAt?._seconds ? c.createdAt._seconds * 1000 : new Date(c.createdAt).getTime();
      const startDate = ts ? new Date(ts).toLocaleString("en-IN") : "—";

      const row = sheet.addRow({
        id: c.id || "—",
        uuid: c.uuid || c.call_id || "—",
        direction: (c.direction || "—").toUpperCase(),
        customer: c.call_to_number || c.customer_no_with_prefix || "—",
        name: c.name || "—",
        callerId: c.caller_id_number || "—",
        agent: c.answered_agent_name || "—",
        agentNum: c.answered_agent_number || "—",
        status: c.call_status || "—",
        duration: c.billsec || 0,
        startTime: c.start_stamp || startDate,
        answerTime: c.answer_stamp || "—",
        endTime: c.end_stamp || "—",
        recording: c.recording_url || "—",
        campaign: c.campaign_name || "—",
        hangup: c.hangup_cause_description || "—",
      });

      if (index % 2 === 0) {
        row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF8FAFC" } };
      }

      if (c.direction === "inbound") {
        row.getCell("direction").font = { color: { argb: "FF7C3AED" }, bold: true };
      } else if (c.direction === "outbound") {
        row.getCell("direction").font = { color: { argb: "FF059669" }, bold: true };
      }

      const status = (c.call_status || "").toLowerCase();
      if (["answered", "completed", "connected"].includes(status)) {
        row.getCell("status").font = { color: { argb: "FF16A34A" }, bold: true };
      } else if (["missed", "no-answer", "failed"].includes(status)) {
        row.getCell("status").font = { color: { argb: "FFDC2626" }, bold: true };
      }
    });

    sheet.autoFilter = { from: "A1", to: "P1" };
    sheet.views = [{ state: "frozen", ySplit: 1 }];

    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        cell.border = {
          top: { style: "thin", color: { argb: "FFE2E8F0" } },
          left: { style: "thin", color: { argb: "FFE2E8F0" } },
          bottom: { style: "thin", color: { argb: "FFE2E8F0" } },
          right: { style: "thin", color: { argb: "FFE2E8F0" } },
        };
      });
    });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename=call-logs-${new Date().toISOString().slice(0, 10)}.xlsx`);

    await workbook.xlsx.write(res);
    res.end();
  } catch (e) {
    console.error("❌ Excel export error:", e);
    return res.status(500).json({ error: e.message });
  }
});

// ─── STATS ────────────────────────────────────────────────────────────────────
app.get("/stats", async (req, res) => {
  try {
    const [callsSnap, leadsSnap] = await Promise.all([
      db.collection("calls").orderBy("createdAt", "desc").limit(500).get(),
      db.collection("leads").get(),
    ]);

    const calls = callsSnap.docs.map((d) => d.data());
    const inbound = calls.filter((c) => c.direction === "inbound");
    const outbound = calls.filter((c) => c.direction === "outbound");
    const answered = calls.filter((c) =>
      ["answered", "completed", "connected", "called"].includes((c.call_status || "").toLowerCase())
    );
    const missed = calls.filter((c) =>
      ["missed", "no-answer", "no_answer", "failed"].includes((c.call_status || "").toLowerCase())
    );
    const withRec = calls.filter((c) => c.recording_url && c.recording_url.startsWith("http"));
    const totalDur = calls.reduce((s, c) => s + (c.billsec || 0), 0);

    const agentStats = {};
    calls.forEach((c) => {
      const name = c.answered_agent_name || "Unknown";
      if (!agentStats[name]) agentStats[name] = { name, calls: 0, duration: 0, missed: 0 };
      agentStats[name].calls++;
      agentStats[name].duration += Number(c.billsec || 0);
      if (["missed", "no-answer", "failed"].includes((c.call_status || "").toLowerCase()))
        agentStats[name].missed++;
    });

    return res.json({
      totalCalls: calls.length,
      inboundCalls: inbound.length,
      outboundCalls: outbound.length,
      answeredCalls: answered.length,
      missedCalls: missed.length,
      totalLeads: leadsSnap.size,
      withRecording: withRec.length,
      totalDuration: totalDur,
      avgDuration: calls.length ? Math.round(totalDur / calls.length) : 0,
      agentStats,
    });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

// Quick sanity check that ACEFONE_TOKEN is set and accepted by Acefone.
// (Doesn't place a real call — just validates auth by calling a lightweight
// read endpoint.)
app.get("/check-acefone-auth", async (req, res) => {
  if (!ACEFONE_TOKEN) {
    return res.status(500).json({ ok: false, error: "ACEFONE_TOKEN not set in .env" });
  }
  try {
    const r = await fetch("https://api.acefone.in/v1/active_calls", {
      headers: { Authorization: `Bearer ${ACEFONE_TOKEN}` },
    });
    const body = await r.json().catch(() => ({}));
    return res.json({ ok: r.ok, status: r.status, body });
  } catch (e) {
    return res.status(500).json({ ok: false, error: e.message });
  }
});

// ─── REMARKS ─────────────────────────────────────
app.post("/remarks", async (req, res) => {
  try {
    const { callId, remark, followUpDate, outcome } = req.body;
    if (!callId || !remark) {
      return res.status(400).json({ error: "callId & remark required" });
    }
    await db.collection("remarks").add({
      callId, remark, outcome: outcome || "", followUpDate: followUpDate || null,
      createdAt: admin.firestore.FieldValue.serverTimestamp(),
    });
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get("/remarks/:callId", async (req, res) => {
  try {
    const snap = await db.collection("remarks")
      .where("callId", "==", req.params.callId)
      .orderBy("createdAt", "desc")
      .limit(5)
      .get();

    res.json(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ─── CUSTOMER HISTORY ─────────────────────────
app.get("/customer-history/:number", async (req, res) => {
  try {
    const n = req.params.number.replace(/\D/g, "").slice(-10);
    const snap = await db.collection("calls").orderBy("createdAt", "desc").limit(20).get();
    const calls = snap.docs
      .map((d) => ({ id: d.id, ...d.data() }))
      .filter((c) =>
        [c.call_to_number, c.caller_id_number, c.customer_no_with_prefix].some(
          (v) => v && String(v).includes(n)
        )
      );
    res.json(calls);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.patch("/call-logs/:id", async (req, res) => {
  try {
    const { name } = req.body;
    await db.collection("calls").doc(req.params.id).update({
      name: name || "", updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    });
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.delete("/call-logs/:id", async (req, res) => {
  try {
    await db.collection("calls").doc(req.params.id).delete();
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ─── Start ────────────────────────────────────────────────────────────────────
app.listen(PORT, async () => {
  console.log(`
╔══════════════════════════════════════════════════════════════╗
║                                                              ║
║   🚀 Acefone Call Portal Backend                             ║
║                                                              ║
║   ✅ Server Running: http://localhost:${PORT}                ║
║                                                              ║
║   📡 Webhook: /webhook (GET & POST)                          ║
║                                                              ║
╚══════════════════════════════════════════════════════════════╝
  `);

  if (ACEFONE_TOKEN) {
    console.log("✅ ACEFONE_TOKEN found — using official Click-to-Call API (no browser login needed)");
  } else {
    console.error("❌ ACEFONE_TOKEN missing — calls will fail until you add it to .env");
  }
});