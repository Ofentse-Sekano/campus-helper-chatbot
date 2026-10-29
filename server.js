const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');

const mailTransporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: process.env.SMTP_USER,
    pass: process.env.SMTP_PASS,
  },
});

async function pushFaultToSharePointViaEmail(faultObj) {
  if (!process.env.SMTP_USER || !process.env.CAMPUS_OUTLOOK_EMAIL) {
    console.warn("⚠️ Email bridge skipped: Set SMTP_USER and CAMPUS_OUTLOOK_EMAIL in .env");
    return;
  }
  try {
    await mailTransporter.sendMail({
      from: `"Helper AI Bot" <${process.env.SMTP_USER}>`,
      to: process.env.CAMPUS_OUTLOOK_EMAIL,
      subject: `[HELPER-NEW-FAULT] ${faultObj.FaultID}`,
      text: JSON.stringify(faultObj), // Pure JSON in the email body for Power Automate
    });
    console.log(`📧 Pushed ${faultObj.FaultID} to Power Automate SharePoint Bridge!`);
  } catch (err) {
    console.error("❌ Email Bridge Error:", err.message);
  }
}
const { GoogleGenerativeAI, HarmCategory, HarmBlockThreshold } = require('@google/generative-ai');
require('dotenv').config();

const app = express();
const port = process.env.PORT || 3000;
const DATA_FILE = path.join(__dirname, 'sharepoint_data.json');

app.use(express.json());
app.use(cors());
app.use(express.static(path.join(__dirname, 'public')));

const API_KEY = process.env.GEMINI_API_KEY;
const genAI = new GoogleGenerativeAI(API_KEY);

// Helper functions to read/write SharePoint synced JSON data
function getSharePointFaults() {
  try {
    if (!fs.existsSync(DATA_FILE)) return [];
    return JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
  } catch (e) {
    console.error("Error reading SharePoint data file:", e);
    return [];
  }
}

function saveSharePointFault(newFault) {
  const faults = getSharePointFaults();
  faults.push(newFault);
  fs.writeFileSync(DATA_FILE, JSON.stringify(faults, null, 2), 'utf8');
  return newFault;
}

// Build dynamic system instruction with LIVE SharePoint FaultReports data
function getDynamicSystemInstruction() {
  const liveFaults = getSharePointFaults();
  const nextIdNumber = liveFaults.length + 1;
  const nextFaultID = `F${String(nextIdNumber).padStart(3, '0')}`;

  return `
You are Helper, the intelligent First-Line IT & Facilities Support Assistant for the Belgium Campus Pretoria Equipment Fault Reporting & Tracking System.
You are directly connected to the team's SharePoint List: "FaultReports" (plus "Equipment" and "Classrooms").

### 1. LIVE SHAREPOINT "FaultReports" DATABASE (CURRENT ITEMS):
${JSON.stringify(liveFaults, null, 2)}

### 2. WHAT YOU CAN DO WITH SHAREPOINT:
- **Status Lookup**: If a user asks about an existing fault (e.g., "Check F001", "What faults are logged?", "Is the projector fault fixed?"), look at the LIVE SHAREPOINT DATABASE above and give them the exact Title, FaultID, Equipment Type, Description, Priority, and Status!
- **Log a New Fault directly to SharePoint**:
  1. First, verify the user's **Full Name**, **Student/Staff Number**, and **Role** (if not already verified). When verified, include this tag once at the top:
     [VERIFIED: FullName | ID: Number | Role: RoleName]
  2. Give a quick 30-second troubleshooting check for their hardware issue.
  3. Ask them: *"Would you like me to log this directly into the SharePoint FaultReports list for you?"* (Or if they explicitly ask you to log/submit it right now, do it immediately).
  4. When logging a new fault into SharePoint, include this EXACT single-line JSON tag at the very bottom of your message (our backend will parse it and add it to SharePoint):
     [CREATE_FAULT: {"Title": "Short Title", "FaultID": "${nextFaultID}", "EquipmentType": "Projector | Classroom PC | Wi-Fi/Network", "Description": "Full description", "Priority": "High | Medium | Low", "Status": "Assigned", "RoomNumber": "Room/Lab"}]

### 3. SHAREPOINT SCHEMA RULES:
- Match the exact SharePoint column choices shown on our site:
  * **Title**: Short summary (e.g., "Projector not working", "PC not powering On")
  * **FaultID**: Format F001, F002, F003... (Next available ID is ${nextFaultID})
  * **Equipment Type**: "Projector", "Classroom PC", "Wi-Fi/Network", "Speakers", "Screen"
  * **Priority**: "High", "Medium", "Low"
  * **Status**: "Assigned", "In Progress", "Resolved"

Keep responses friendly, structured with bullet points, and under 160 words.
`;
}

const sessions = new Map();
const MODEL_PRIORITY = ["gemini-3.8-flash", "gemini-3.8-flash-lite", "gemini-3.8-pro"];
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function sendMessageWithRetry(sessionId, userMessage) {
  let lastError = null;
  const currentInstruction = getDynamicSystemInstruction();

  for (const modelName of MODEL_PRIORITY) {
    for (let attempt = 1; attempt <= 2; attempt++) {
      try {
        const model = genAI.getGenerativeModel({
          model: modelName,
          systemInstruction: currentInstruction,
        });

        const existingHistory = sessions.get(sessionId) || [];
        const chat = model.startChat({
          history: existingHistory,
          generationConfig: { temperature: 0.25, maxOutputTokens: 450 },
          safetySettings: [
            {
              category: HarmCategory.HARM_CATEGORY_HARASSMENT,
              threshold: HarmBlockThreshold.BLOCK_MEDIUM_AND_ABOVE,
            },
          ],
        });

        const result = await chat.sendMessage(userMessage);
        const updatedHistory = await chat.getHistory();
        sessions.set(sessionId, updatedHistory);

        return result.response.text();
      } catch (err) {
        lastError = err;
        const isOverloaded = err.status === 503 || (err.message && err.message.includes("503"));
        if (isOverloaded) {
          console.warn(`⚠️ [${modelName}] Busy (503). Retrying ${attempt}/2...`);
          await delay(1500);
          continue;
        }
        break;
      }
    }
  }
  throw lastError;
}

// REST Endpoints for SharePoint / Power Apps / Frontend
app.get('/api/faults', (req, res) => {
  res.json(getSharePointFaults());
});

app.post('/api/faults', (req, res) => {
  const current = getSharePointFaults();
  const nextID = `F${String(current.length + 1).padStart(3, '0')}`;
  const fault = {
    Title: req.body.Title || "Equipment Fault",
    FaultID: req.body.FaultID || nextID,
    EquipmentType: req.body.EquipmentType || "Classroom PC",
    Description: req.body.Description || "",
    Priority: req.body.Priority || "Medium",
    Status: req.body.Status || "Assigned",
    RoomNumber: req.body.RoomNumber || "Unspecified"
  };
  saveSharePointFault(fault);
  res.json({ success: true, fault });
});

// Bulk sync endpoint from SharePoint / Power Automate
app.post('/api/sync-sharepoint', (req, res) => {
  if (Array.isArray(req.body.items)) {
    fs.writeFileSync(DATA_FILE, JSON.stringify(req.body.items, null, 2), 'utf8');
    return res.json({ success: true, count: req.body.items.length });
  }
  res.status(400).json({ error: "Expected items array" });
});

app.post('/chat', async (req, res) => {
  try {
    const userMessage = req.body.userInput || req.body.message;
    const sessionId = req.body.sessionId || "default-session";

    if (!userMessage) return res.status(400).json({ error: "Please enter a message." });
    if (req.body.reset) sessions.delete(sessionId);

    const rawText = await sendMessageWithRetry(sessionId, userMessage);

    // 1. Check for [VERIFIED: ...] tag
    let verifiedUser = null;
    const verifyRegex = /\[VERIFIED:\s*([^|\]]+)\s*\|\s*ID:\s*([^|\]]+)\s*\|\s*Role:\s*([^\]]+)\]/i;
    const verifyMatch = rawText.match(verifyRegex);
    if (verifyMatch) {
      verifiedUser = {
        fullName: verifyMatch[1].trim(),
        idNumber: verifyMatch[2].trim(),
        role: verifyMatch[3].trim()
      };
    }

  // 2. Check for [CREATE_FAULT: {...}] tag to log new item in SharePoint store
    let createdFault = null;
    const createRegex = /\[CREATE_FAULT:\s*(\{.*?\})\]/s;
    const createMatch = rawText.match(createRegex);
    if (createMatch) {
      try {
        const faultObj = JSON.parse(createMatch[1]);
        createdFault = saveSharePointFault(faultObj);
        console.log("✅ New Fault Saved Locally:", createdFault);
        
        // Trigger the Power Automate SharePoint Bridge!
        pushFaultToSharePointViaEmail(createdFault);
      } catch (e) {
        console.error("Failed to parse CREATE_FAULT JSON:", e);
      }
    }
    
    const cleanText = rawText
      .replace(verifyRegex, '')
      .replace(createRegex, '')
      .trim();

    res.json({
      response: cleanText,
      verifiedUser,
      createdFault,
      allFaults: getSharePointFaults()
    });
  } catch (error) {
    console.error("Gemini API Error:", error);
    res.status(500).json({
      error: "Google's AI server is momentarily busy. Please click Send again in a few seconds."
    });
  }
});

app.listen(port, () => {
  console.log(`✅ Helper + SharePoint Bridge running at http://localhost:${port}`);
});