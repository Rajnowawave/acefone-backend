const admin = require("firebase-admin");
const bcrypt = require("bcryptjs");
const serviceAccount = require("./firebase.json");

admin.initializeApp({ credential: admin.credential.cert(serviceAccount) });
const db = admin.firestore();

const agents = [
  { 
    name: "Neelam", 
    agentId: "0502190850001", 
    phone: "+919251651958", 
    email: "customercare@adinath.net.in", 
    password: "Office@2005", // Plain password (will be hashed)
    role: "agent",
    status: "active" 
  },
  { 
    name: "Bhavika", 
    agentId: "0502190850002", 
    phone: "+919251651956", 
    email: "Bhavikakewalramani70@gmail.com", 
    password: "Office@2004",
    role: "agent",
    status: "active" 
  },
  { 
    name: "Amit Kumar", 
    agentId: "0502190850003", 
    phone: "+919876543212", 
    email: "amit@company.com", 
    password: "amit123",
    role: "agent",
    status: "active" 
  },
  { 
    name: "Sneha Patel", 
    agentId: "0502190850004", 
    phone: "+919876543213", 
    email: "sneha@company.com", 
    password: "sneha123",
    role: "agent",
    status: "active" 
  },
  { 
    name: "Vikram Mehta", 
    agentId: "0502190850005", 
    phone: "+919876543214", 
    email: "vikram@company.com", 
    password: "vikram123",
    role: "agent",
    status: "inactive" 
  },
  {
    name: "Admin User",
    agentId: "ADMIN001",
    phone: "+919999999999",
    email: "admin@company.com",
    password: "admin123",
    role: "admin",
    status: "active"
  }
];

(async () => {
  console.log("🔒 Adding agents with encrypted passwords...\n");
  
  for (const agent of agents) {
    // Hash password
    const hashedPassword = await bcrypt.hash(agent.password, 10);
    
    await db.collection("agents").add({
      name: agent.name,
      agentId: agent.agentId,
      phone: agent.phone,
      email: agent.email,
      password: hashedPassword, // Store hashed password
      role: agent.role,
      status: agent.status,
      totalCalls: 0,
      answeredCalls: 0,
      missedCalls: 0,
      totalDuration: 0,
      createdAt: new Date(),
    });
    
    console.log(`✅ Added: ${agent.name}`);
    console.log(`   Email: ${agent.email}`);
    console.log(`   Password: ${agent.password}`);
    console.log(`   Role: ${agent.role}\n`);
  }
  
  console.log("🎉 All agents added!");
  console.log("\n📋 Login Credentials:");
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  console.log("Admin Login:");
  console.log("  Email: admin@company.com");
  console.log("  Password: admin123");
  console.log("\nAgent Logins:");
  agents.filter(a => a.role === "agent").forEach(a => {
    console.log(`  ${a.name}: ${a.email} / ${a.password}`);
  });
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");
  
  process.exit(0);
})();