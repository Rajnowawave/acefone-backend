// test.mjs
const res = await fetch("const API = import.meta.env.VITE_API_URL || "http://localhost:5000";/call", {
  method: "POST",
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify({ customer: "7792097977" })
});

console.log("Status:", res.status);
console.log("Response:", await res.text());