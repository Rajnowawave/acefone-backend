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

const serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT);
admin.initializeApp({ credential: admin.credential.cert(serviceAccount) });
const db = admin.firestore();

const PORT = Number(process.env.PORT || 5000);
const ACEFONE_CALLER_ID = process.env.ACEFONE_CALLER_ID || "+918062491504";
const ACEFONE_USER_ID = process.env.ACEFONE_USER_ID || "219085";
const ACEFONE_EMAIL = process.env.ACEFONE_EMAIL || "customercare@adinath.net.in";
const ACEFONE_PASSWORD = process.env.ACEFONE_PASSWORD || "Office@2005";

const AGENTS = [
  { id: "0502190850001", name: "Neelam", number: "919251651958" },
  { id: "0502190850002", name: "Bhavika", number: "919251651956" },
  { id: "0502190850003", name: "Tushar Bhandari", number: "917976630010" },
  { id: "0502190850004", name: "Vikash Singhvi", number: "919509805201" },
  { id: "0502190850005", name: "Amit Sharma", number: "918094121221" },
];

let sessionData = { cookieString: "", csrfToken: "" };
let isRefreshing = false;

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

// ─── Session Management ───────────────────────────────────────────────────────
function loadSession() {
  try {
    const raw = fs.readFileSync("./session.json", "utf8");
    sessionData = JSON.parse(raw);
    console.log("✅ Session loaded:", sessionData.savedAt);
  } catch {
    console.error("❌ session.json not found");
  }
}

async function refreshSession() {
  if (isRefreshing) {
    await new Promise((r) => setTimeout(r, 8000));
    return sessionData.csrfToken ? true : false;
  }
  isRefreshing = true;
  console.log("🔄 Refreshing session...");

  try {
    const puppeteer = require("puppeteer");
    const browser = await puppeteer.launch({
      headless: true,
      args: ["--no-sandbox", "--disable-setuid-sandbox"],
    });

    const page = await browser.newPage();
    await page.setUserAgent(
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    );
    await page.goto("https://console.acefone.in/login", {
      waitUntil: "networkidle2",
      timeout: 30000,
    });
    await page.waitForSelector("#loginId", { timeout: 15000 });

    await page.$eval("#loginId", (el) => (el.value = ""));
    await page.$eval("#password", (el) => (el.value = ""));
    await page.type("#loginId", ACEFONE_EMAIL, { delay: 60 });
    await page.type("#password", ACEFONE_PASSWORD, { delay: 60 });
    await page.click("#login_button");

    await page.waitForFunction(
      () => !window.location.href.includes("/login"),
      { timeout: 25000 }
    );

    await page.goto("https://console.acefone.in/click_to_call_api_keys", {
  waitUntil: "networkidle2",
  timeout: 20000,
    });

    const csrfToken = await page.evaluate(
      () =>
        document.querySelector("meta[name='csrf-token']")?.content ||
        document.querySelector("input[name='_token']")?.value ||
        ""
    );

    const cookies = await page.cookies();
    const cookieString = cookies.map((c) => `${c.name}=${c.value}`).join("; ");

    sessionData = {
      cookieString,
      csrfToken,
      savedAt: new Date().toISOString(),
    };
    fs.writeFileSync("./session.json", JSON.stringify(sessionData, null, 2));

    console.log("✅ Session refreshed:", sessionData.savedAt);
    await browser.close();
    isRefreshing = false;
    return true;
  } catch (err) {
    console.error("❌ Session refresh failed:", err.message);
    isRefreshing = false;
    return false;
  }
}

setInterval(() => {
  console.log("⏰ Auto session refresh...");
  refreshSession();
}, 6 * 60 * 60 * 1000);

// ─── Acefone Click-to-Call ────────────────────────────────────────────────────
async function makeAcefoneCall(customerNumber, agentId) {
  const digits = String(customerNumber).replace(/\D/g, "").slice(-10);
  const phone = "91" + digits;
  const agent = AGENTS.find((a) => a.id === agentId);

  if (!agent) {
    throw new Error("Invalid agent selected");
  }

  console.log(`📞 Calling ${phone} via agent: ${agent.name} (${agent.id})`);

  const doCall = async () => {
    const params = new URLSearchParams({
      phone_ctc: phone,
      ctc_caller_id: ACEFONE_CALLER_ID,
      ctc_agent_id: agent.id,
      user_id: ACEFONE_USER_ID,
      is_extension_call: "false",
      _token: sessionData.csrfToken,
    });

    return fetch("https://console.acefone.in/click-to-call", {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "X-CSRF-TOKEN": sessionData.csrfToken,
        "X-Requested-With": "XMLHttpRequest",
        Cookie: sessionData.cookieString,
        Referer: "https://console.acefone.in/click_to_call_api_keys",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
      },
      body: params.toString(),
    });
  };

  let response = await doCall();

  if (response.status === 419) {
    console.log("⚠️ 419 — refreshing session...");
    const ok = await refreshSession();
    if (ok) {
      console.log("🔁 Retrying call...");
      response = await doCall();
    }
  }

  return { response, agent };
}

// ✅ WEBHOOK HANDLER FUNCTION
// ✅ WEBHOOK HANDLER - Recording URL properly saved
async function handleWebhook(d, res) {
  try {
    console.log("🔔 Webhook Data:\n", JSON.stringify(d, null, 2));

    const direction = detectDirection(d);
    const uuid = pick(d, "uuid", "call_id");
    const callTo = pick(d, "call_to_number");
    const callerNum = pick(d, "caller_id_number");
    const callStatus = pick(d, "call_status") || "completed";
    
    // ✅ RECORDING URL - Direct from webhook (with token)
    let recordingUrl = pick(d, "recording_url");
    
    console.log(`🎙 Recording URL: ${recordingUrl ? "✅ Found" : "❌ Not found"}`);

    const billsec = num(pick(d, "billsec"));
    const duration = num(pick(d, "duration")) || billsec;
    
    // ✅ Handle answered_agent object or string
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
    // ✅ FINAL INBOUND CUSTOMER NUMBER FIX
let clientNumber = "";
let didNumber = "";

if (direction === "inbound") {
  clientNumber =
    pick(d, "client_number") ||
    pick(d, "caller_id_num") ||
    pick(d, "customer_no_with_prefix") ||
    "";

  didNumber =
    pick(d, "did_number") ||
    pick(d, "call_to_number") ||
    "";

  // remove +91
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
      recording_url: recordingUrl || "", // ✅ Save as-is
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

    // OUTBOUND: Update latest initiated call
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
          answered_agent_name:
            agentName || prev.answered_agent_name || "",

          answered_agent_number:
            agentNumber || prev.answered_agent_number || "",

          answered_agent:
            agentId || prev.answered_agent || "",

          updatedAt: admin.firestore.FieldValue.serverTimestamp(),
        });

        console.log(
          `✅ OUTBOUND updated: ${existing.docs[0].id} | Recording: ${recordingUrl ? "✅" : "❌"}`
        );

        return res.status(200).send("OK");
      }
    }

    // INBOUND: Create new
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
    if (!doc.exists) {
      return res.status(404).send("Call not found");
    }

    const data = doc.data();
    const url = data.recording_url;

    if (!url) {
      return res.status(404).send("No recording available");
    }

    // ✅ Proxy the recording through our server
    // This allows playback even if Acefone requires authentication
    const response = await fetch(url);
    
    if (!response.ok) {
      return res.status(404).send("Recording not accessible");
    }

    // Set proper headers for audio streaming
    res.setHeader('Content-Type', 'audio/mpeg');
    res.setHeader('Accept-Ranges', 'bytes');
    res.setHeader('Cache-Control', 'public, max-age=3600');
    
    // Stream the audio
    response.body.pipe(res);
  } catch (e) {
    console.error("Recording proxy error:", e);
    return res.status(500).send("Error fetching recording");
  }
});

app.get("/contacts", async (req, res) => {
  try {
    const snap = await db.collection("contacts").get();
    const contacts = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    res.json(contacts);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ✅ Get recording info (URL only)
app.get("/recording-info/:id", async (req, res) => {
  try {
    const doc = await db.collection("calls").doc(req.params.id).get();
    if (!doc.exists) {
      return res.status(404).json({ error: "Call not found" });
    }

    const data = doc.data();
    const url = data.recording_url;

    if (!url) {
      return res.status(404).json({ error: "No recording available" });
    }

    return res.json({ 
      url,
      callId: data.call_id,
      uuid: data.uuid,
      duration: data.billsec,
      status: data.call_status
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
          <div class="url">${req.protocol}://${req.get('host')}/webhook</div>
        </div>
        
        <div class="endpoint">
          <div class="label">Time:</div>
          <div class="url">${new Date().toLocaleString('en-IN')}</div>
        </div>
      </div>
    </body>
    </html>
  `);
});

// ✅ WEBHOOK - Both GET and POST
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
    return res.json(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

app.post("/leads", async (req, res) => {
  try {
    const { number, name, notes } = req.body;
    if (!number) return res.status(400).json({ error: "Number required" });
    const ref = await db.collection("leads").add({
      number, name: name || "", notes: notes || "", status: "pending", callCount: 0, createdAt: admin.firestore.FieldValue.serverTimestamp(),
    });
    return res.json({ id: ref.id });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

app.put("/leads/:id", async (req, res) => {
  try {
    const { number, name, notes, status } = req.body;
    await db.collection("leads").doc(req.params.id).update({
      number, name, notes, status, updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    });
    return res.json({ success: true });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

app.delete("/leads/:id", async (req, res) => {
  try {
    await db.collection("leads").doc(req.params.id).delete();
    return res.json({ message: "Deleted" });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

// ═══════════════════════════════════════════════
// LEAD MANAGEMENT ROUTES
// ═══════════════════════════════════════════════

// Get all leads
app.get("/leads", async (req, res) => {
  try {
    const snap = await db.collection("leads").orderBy("createdAt", "desc").get();
    const leads = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
    return res.json(leads);
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

// Create single lead
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

// Bulk upload leads
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

// Update lead
app.put("/leads/:id", async (req, res) => {
  try {
    const {
      name,
      contact,
      city,
      propertyType,
      remarks,
      status,
      calledBy,
      followUps,
    } = req.body;

    const updateData = {
      updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    };

    if (name !== undefined) updateData.name = name.trim();
    if (contact !== undefined) updateData.contact = contact.trim();
    if (city !== undefined) updateData.city = city.trim();
    if (propertyType !== undefined) updateData.propertyType = propertyType;
    if (remarks !== undefined) updateData.remarks = remarks;
    if (status !== undefined) updateData.status = status;
    if (calledBy !== undefined) updateData.calledBy = calledBy;
    if (followUps !== undefined) updateData.followUps = followUps;

    // If status changed from "Not Called", increment call count
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

// Delete lead
app.delete("/leads/:id", async (req, res) => {
  try {
    await db.collection("leads").doc(req.params.id).delete();
    return res.json({ success: true });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

// Get lead history
app.get("/leads/:id/history", async (req, res) => {
  try {
    const leadDoc = await db.collection("leads").doc(req.params.id).get();
    if (!leadDoc.exists) {
      return res.status(404).json({ error: "Lead not found" });
    }

    const leadData = leadDoc.data();
    const contact = leadData.contact;

    // Find all call logs for this contact
    const callsSnap = await db
      .collection("calls")
      .orderBy("createdAt", "desc")
      .limit(100)
      .get();

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

// Export leads to Excel
app.get("/leads/export", async (req, res) => {
  try {
    const { date, status } = req.query;

    let query = db.collection("leads").orderBy("createdAt", "desc");

    const snap = await query.get();
    let leads = snap.docs.map((d) => ({ id: d.id, ...d.data() }));

    // Filter by date
    if (date) {
      leads = leads.filter((l) => l.date === date);
    }

    // Filter by status
    if (status) {
      leads = leads.filter((l) => l.status === status);
    }

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Leads");

    // Define columns
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

    // Style header
    sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" }, size: 11 };
    sheet.getRow(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF2C3E6B" },
    };
    sheet.getRow(1).alignment = { horizontal: "center", vertical: "middle" };
    sheet.getRow(1).height = 25;

    // Add data
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
              lead.lastCalledAt._seconds
                ? lead.lastCalledAt._seconds * 1000
                : lead.lastCalledAt
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

      // Alternate row colors
      if (index % 2 === 0) {
        row.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFF8FAFC" },
        };
      }

      // Highlight "Not Called" rows
      if (!lead.status || lead.status === "Not Called") {
        row.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFFFBEB" },
        };
      }
    });

    // Auto-filter
    sheet.autoFilter = { from: "A1", to: "M1" };

    // Freeze header row
    sheet.views = [{ state: "frozen", ySplit: 1 }];

    // Add borders
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

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename=leads-${new Date().toISOString().split("T")[0]}.xlsx`
    );

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
    const text = await response.text();
    console.log("Acefone →", response.status, text);

    const digits = String(customer).replace(/\D/g, "").slice(-10);

    const callDoc = await db.collection("calls").add({
      direction: "outbound",
      call_to_number: "91" + digits,
      caller_id_number: ACEFONE_CALLER_ID,
      answered_agent_name: agent.name,
      answered_agent_number: agent.number,
      answered_agent: agent.id,
      name: name || "",
      leadId: leadId || "",
      call_status: response.ok ? "initiated" : "failed",
      recording_url: "",
      billsec: 0,
      createdAt: admin.firestore.FieldValue.serverTimestamp(),
      updatedAt: admin.firestore.FieldValue.serverTimestamp()
    });

    console.log(`✅ Call logged: ${callDoc.id} | Agent: ${agent.name}`);

    if (leadId) {
      const lRef = db.collection("leads").doc(leadId);
      const lDoc = await lRef.get();
      if (lDoc.exists) {
        await lRef.update({
          callCount: (lDoc.data().callCount || 0) + 1,
          status: response.ok ? "called" : "failed",
          lastCalledAt: admin.firestore.FieldValue.serverTimestamp(),
          lastAgent: agent.name,
        });
      }
    }

    if (response.status === 419) {
      return res.status(401).json({ error: "Session expired — please refresh" });
    }

    if (response.ok) {
      return res.json({ success: true, agent: agent.name });
    }

    return res.status(response.status).json({ error: "Call failed", details: text });
  } catch (e) {
    console.error("Call error:", e);
    return res.status(500).json({ error: e.message });
  }
});

// ✅ RECORDING PROXY
app.get("/recording-proxy/:id", async (req, res) => {
  try {
    const doc = await db.collection("calls").doc(req.params.id).get();
    if (!doc.exists) {
      return res.status(404).json({ error: "Call not found" });
    }

    const data = doc.data();
    const url = data.recording_url || "";

    if (!url) {
      return res.status(404).json({ error: "No recording available" });
    }

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
    if (direction && direction !== "all")
      ref = ref.where("direction", "==", direction);

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
    if (direction && direction !== "all")
      ref = ref.where("direction", "==", direction);

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
      if (!agentStats[name])
        agentStats[name] = { name, calls: 0, duration: 0, missed: 0 };
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

app.post("/refresh-session", async (req, res) => {
  const ok = await refreshSession();
  if (ok) return res.json({ success: true, savedAt: sessionData.savedAt });
  return res.status(500).json({ error: "Session refresh failed" });
});

// ─── REMARKS ─────────────────────────────────────
app.post("/remarks", async (req, res) => {
  try {
    const { callId, remark, followUpDate, outcome } = req.body;
    if (!callId || !remark) {
      return res.status(400).json({ error: "callId & remark required" });
    }
    await db.collection("remarks").add({
      callId, remark, outcome: outcome || "", followUpDate: followUpDate || null, createdAt: admin.firestore.FieldValue.serverTimestamp(),
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

    res.json(snap.docs.map(d => ({
      id: d.id,
      ...d.data()
    })));
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ─── CUSTOMER HISTORY ─────────────────────────
app.get("/customer-history/:number", async (req, res) => {
  try {
    const num = req.params.number.replace(/\D/g, "").slice(-10);
    const snap = await db.collection("calls").orderBy("createdAt", "desc").limit(20).get();
    const calls = snap.docs
      .map((d) => ({ id: d.id, ...d.data() }))
      .filter((c) =>
        [c.call_to_number, c.caller_id_number, c.customer_no_with_prefix].some(
          (v) => v && String(v).includes(num)
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
║     
║   📡 Webhook: /webhook (GET & POST)                          ║
║                                                              ║
╚══════════════════════════════════════════════════════════════╝
  `);

  loadSession();

  if (sessionData.savedAt) {
    const savedTime = new Date(sessionData.savedAt).getTime();
    const hoursSince = (Date.now() - savedTime) / (1000 * 60 * 60);
    if (hoursSince > 5) {
      console.log(`⚠️ Session ${Math.round(hoursSince)}h old — refreshing...`);
      await refreshSession();
    }
  }
});