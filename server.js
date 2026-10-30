const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');
const { GoogleGenerativeAI, HarmCategory, HarmBlockThreshold } = require('@google/generative-ai');
require('dotenv').config();

const app = express();
const port = process.env.PORT || 3000;

const DATA_FILE = path.join(__dirname, 'sharepoint_data.json');
const CSV_FILE = path.join(__dirname, 'FaultReports.csv');

app.use(express.json({ limit: '10mb' }));
app.use(cors());
app.use(express.static(path.join(__dirname, 'public')));

const API_KEY = process.env.GEMINI_API_KEY;
const genAI = new GoogleGenerativeAI(API_KEY || '');

// Email Bridge for Power Automate -> Real SharePoint List & Microsoft Teams
const mailTransporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: process.env.SMTP_USER,
    pass: process.env.SMTP_PASS,
  },
});

async function pushBotFaultToPowerAutomate(faultObj) {
  if (!process.env.SMTP_USER || !process.env.SMTP_PASS || !process.env.CAMPUS_OUTLOOK_EMAIL) {
    console.warn("⚠️ [RENDER POWER AUTOMATE SKIPPED]: Add SMTP_USER, SMTP_PASS, and CAMPUS_OUTLOOK_EMAIL to Render Environment Variables!");
    return false;
  }
  try {
    await mailTransporter.sendMail({
      from: `"Helper AI Bot" <${process.env.SMTP_USER}>`,
      to: process.env.CAMPUS_OUTLOOK_EMAIL,
      subject: `[HELPER-NEW-FAULT] ${faultObj.FaultID}`,
      text: JSON.stringify(faultObj),
    });
    console.log(`🚀 [POWER AUTOMATE SENT]: Pushed ${faultObj.FaultID} to ${process.env.CAMPUS_OUTLOOK_EMAIL}`);
    return true;
  } catch (err) {
    console.error("❌ [POWER AUTOMATE EMAIL ERROR]:", err.message);
    return false;
  }
}

const DEFAULT_FAULTS = [
  {
    Title: "Projector not working",
    FaultID: "F001",
    EquipmentType: "Projector",
    Description: "The projector does not display an image when connected to the classroom computer.",
    Priority: "High",
    Status: "Assigned"
  },
  {
    Title: "PC not powering On",
    FaultID: "F002",
    EquipmentType: "Classroom PC",
    Description: "PC does not switch on at the power button or wall socket.",
    Priority: "High",
    Status: "In Progress"
  }
];

let cachedFaults = null;

function parseSharePointCSV(csvText) {
  const cleanText = csvText.replace(/^\uFEFF/, '');
  const lines = cleanText.split(/\r?\n/).filter(line => line.trim() !== '');
  if (lines.length < 2) return [];

  const firstLine = lines[0];
  const delimiter = (firstLine.split(';').length > firstLine.split(',').length) ? ';' : ',';

  const splitCSVLine = (line) => {
    const result = [];
    let cur = '', inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') {
        inQuotes = !inQuotes;
      } else if (ch === delimiter && !inQuotes) {
        result.push(cur.trim());
        cur = '';
      } else {
        cur += ch;
      }
    }
    result.push(cur.trim());
    return result.map(val => val.replace(/^"|"$/g, '').trim());
  };

  const rawHeaders = splitCSVLine(lines[0]);
  const normHeaders = rawHeaders.map(h => h.toLowerCase().replace(/[\s_]+/g, ''));

  const items = [];
  for (let i = 1; i < lines.length; i++) {
    const cols = splitCSVLine(lines[i]);
    if (cols.every(c => c === '')) continue;

    const row = {};
    normHeaders.forEach((h, idx) => {
      row[h] = cols[idx] || '';
    });

    items.push({
      Title: row['title'] || 'Equipment Fault',
      FaultID: row['faultid'] || `F${String(i).padStart(3, '0')}`,
      EquipmentType: row['equipmenttype'] || 'Projector',
      Description: row['description'] || '',
      Priority: row['priority'] || 'High',
      Status: row['status'] || 'Assigned'
    });
  }
  return items;
}

function getNextFaultID(faults) {
  let maxNum = faults.length;
  faults.forEach(f => {
    const m = String(f.FaultID || '').match(/F(\d+)/i);
    if (m) {
      const n = parseInt(m[1], 10);
      if (n > maxNum) maxNum = n;
    }
  });
  return `F${String(maxNum + 1).padStart(3, '0')}`;
}

function getSharePointFaults(forceReload = false) {
  if (cachedFaults && !forceReload) {
    return cachedFaults;
  }

  try {
    if (fs.existsSync(DATA_FILE)) {
      const data = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
      if (Array.isArray(data) && data.length > 0) {
        cachedFaults = data;
        return cachedFaults;
      }
    }

    if (fs.existsSync(CSV_FILE)) {
      const csvContent = fs.readFileSync(CSV_FILE, 'utf8');
      const parsed = parseSharePointCSV(csvContent);
      if (parsed.length > 0) {
        cachedFaults = parsed;
        fs.writeFileSync(DATA_FILE, JSON.stringify(parsed, null, 2), 'utf8');
        return cachedFaults;
      }
    }
  } catch (e) {
    console.error("Error reading fault data:", e);
  }

  cachedFaults = [...DEFAULT_FAULTS];
  return cachedFaults;
}

function saveNewFault(newFault) {
  const faults = getSharePointFaults();
  if (!newFault.FaultID || faults.some(f => f.FaultID === newFault.FaultID)) {
    newFault.FaultID = getNextFaultID(faults);
  }
  faults.push(newFault);
  cachedFaults = faults;
  fs.writeFileSync(DATA_FILE, JSON.stringify(faults, null, 2), 'utf8');
  return newFault;
}

function isFaultReportMessage(userMessage) {
  const msg = userMessage.toLowerCase();
  if (msg.startsWith('🔍') || msg.includes('what faults') || msg.includes('check all') || msg.includes('check status')) {
    return false;
  }
  const hasSymptom =
    msg.includes('log') ||
    msg.includes('broken') ||
    msg.includes('not working') ||
    msg.includes('not powering') ||
    msg.includes('crash') ||
    msg.includes('fault') ||
    msg.includes('issue') ||
    msg.includes('offline') ||
    msg.includes('flicker');

  const hasEquipment =
    msg.includes('projector') ||
    msg.includes('pc') ||
    msg.includes('computer') ||
    msg.includes('lab') ||
    msg.includes('room') ||
    msg.includes('wifi') ||
    msg.includes('wi-fi') ||
    msg.includes('network') ||
    msg.includes('screen') ||
    msg.includes('speaker');

  return hasSymptom && hasEquipment;
}

function buildFaultFromUserMessage(userMessage) {
  const faults = getSharePointFaults();
  const nextId = getNextFaultID(faults);
  const msg = userMessage.toLowerCase();

  const eqType = (msg.includes('pc') || msg.includes('computer') || msg.includes('lab'))
    ? 'Classroom PC'
    : (msg.includes('wifi') || msg.includes('wi-fi') || msg.includes('network') || msg.includes('internet'))
    ? 'Wi-Fi/Network'
    : 'Projector';

  const priority = (msg.includes('high') || msg.includes('exam') || msg.includes('urgent') || msg.includes('crash'))
    ? 'High'
    : (msg.includes('low')) ? 'Low' : 'Medium';

  const cleanDesc = userMessage.replace(/^➕\s*/i, '').replace(/^🚨\s*/i, '').trim();
  const title = eqType === 'Projector' ? 'Projector not working'
              : eqType === 'Classroom PC' ? 'PC not powering On'
              : 'Network connectivity fault';

  return {
    Title: title,
    FaultID: nextId,
    EquipmentType: eqType,
    Description: cleanDesc,
    Priority: priority,
    Status: 'Assigned'
  };
}

function getDynamicSystemInstruction() {
  const liveFaults = getSharePointFaults();
  const nextFaultID = getNextFaultID(liveFaults);
  const compactFaults = JSON.stringify(liveFaults);

  return `You are Helper, the IT Support Assistant for Belgium Campus Equipment Fault Reporting.
Connected SharePoint "FaultReports" list: ${compactFaults}

RULES:
1. Fault Lookup: Use the JSON list above to answer questions about any FaultID, Title, EquipmentType, Description, Priority, or Status.
2. Verification: When a user provides their Name and Student/Staff ID, output once at top:
[VERIFIED: FullName | ID: Number | Role: RoleName]
3. Next Fault ID: The next available FaultID is ${nextFaultID}.
Keep replies under 120 words with bullet points.`;
}

const sessions = new Map();
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function withTimeout(promise, ms) {
  return Promise.race([
    promise,
    new Promise((_, reject) =>
      setTimeout(() => reject(new Error(`TIMEOUT_AFTER_${ms}MS`)), ms)
    )
  ]);
}

function generateInstantLocalResponse(userMessage, createdFault = null) {
  const faults = getSharePointFaults();
  const msg = userMessage.toLowerCase();

  let verifyTag = '';
  const idMatch = userMessage.match(/\b(\d{5,8})\b/);
  if (idMatch) {
    const idNum = idMatch[1];
    const role = msg.includes('lecturer') ? 'Lecturer' : msg.includes('staff') ? 'Staff' : 'Student';
    const nameClean = userMessage
      .replace(idNum, '')
      .replace(/student|lecturer|staff|i am|i'm|my name is|,|-|\|/gi, '')
      .trim() || 'Ofentse Sekano';
    verifyTag = `[VERIFIED: ${nameClean} | ID: ${idNum} | Role: ${role}]\n`;
  }

  if (createdFault) {
    return `${verifyTag}I have logged this **${createdFault.Priority} Priority** fault for the **${createdFault.EquipmentType}** into the **SharePoint FaultReports** list:\n* **Assigned FaultID:** **${createdFault.FaultID}**\n* **Issue:** ${createdFault.Title} (${createdFault.Description})\n* **Quick 30s Check:** Please verify the power cable and display connection are firmly seated while a technician is dispatched.`;
  }

  const faultIdMatch = userMessage.match(/\b(F\d{3})\b/i);
  if (faultIdMatch) {
    const targetId = faultIdMatch[1].toUpperCase();
    const found = faults.find(f => f.FaultID.toUpperCase() === targetId);
    if (found) {
      return `${verifyTag}Here is the live SharePoint status for **${found.FaultID}**:\n* **Title:** ${found.Title}\n* **Equipment Type:** ${found.EquipmentType}\n* **Priority:** ${found.Priority}\n* **Status:** ${found.Status || 'Assigned'}\n* **Description:** ${found.Description}`;
    }
  }

  if (msg.includes('check') || msg.includes('what faults') || msg.includes('status') || msg.includes('list') || msg.includes('all faults')) {
    const listSummary = faults
      .map(f => `* **${f.FaultID} — ${f.Title}** (${f.EquipmentType}) | Priority: **${f.Priority}** | Status: **${f.Status || 'Assigned'}** — ${f.Description}`)
      .join('\n');
    return `${verifyTag}Here are the current fault reports synced from **SharePoint (FaultReports)** (${faults.length} total):\n\n${listSummary}`;
  }

  if (verifyTag) {
    return `${verifyTag}Thank you for verifying your details! Tell me about any broken **Projector**, **Classroom PC**, or **Wi-Fi** issue and I will log it immediately.`;
  }

  return `Hi! I'm **Helper**, connected to your **SharePoint FaultReports** list (${faults.length} active faults loaded).\n* Ask me to check **F001**, **F002**, or all current faults.\n* Describe an equipment fault (e.g., *"The PC in Lab 14 is not working, priority high"*) to log a new ticket!`;
}

async function sendMessageFast(sessionId, userMessage, createdFault = null) {
  if (createdFault) {
    return generateInstantLocalResponse(userMessage, createdFault);
  }

  const currentInstruction = getDynamicSystemInstruction();
  const fullHistory = sessions.get(sessionId) || [];
  const recentHistory = fullHistory.slice(-4);
  const modelName = "gemini-3.8-flash";

  for (let attempt = 1; attempt <= 2; attempt++) {
    try {
      const startTime = Date.now();
      const model = genAI.getGenerativeModel({
        model: modelName,
        systemInstruction: currentInstruction,
      });

      const chat = model.startChat({
        history: recentHistory,
        generationConfig: {
          temperature: 0.2,
          maxOutputTokens: 300,
        },
        safetySettings: [
          {
            category: HarmCategory.HARM_CATEGORY_HARASSMENT,
            threshold: HarmBlockThreshold.BLOCK_MEDIUM_AND_ABOVE,
          },
        ],
      });

      const result = await withTimeout(chat.sendMessage(userMessage), 8000);
      const responseText = result.response.text();

      const updatedHistory = await chat.getHistory();
      sessions.set(sessionId, updatedHistory);

      console.log(`⚡ [${modelName}] Responded in ${Date.now() - startTime}ms`);
      return responseText;
    } catch (err) {
      console.warn(`⚠️ [${modelName}] Attempt ${attempt}/2 failed (${err.status || err.message}).`);
      if (attempt === 1) {
        await delay(1000);
        continue;
      }
      break;
    }
  }

  return generateInstantLocalResponse(userMessage, createdFault);
}

app.get('/ping', (req, res) => res.json({ ok: true }));

app.get('/api/faults', (req, res) => {
  const force = req.query.refresh === 'true';
  res.json(getSharePointFaults(force));
});

app.post('/api/import-sharepoint-csv', (req, res) => {
  try {
    const csvText = req.body.csv;
    if (!csvText) return res.status(400).json({ error: "No CSV text provided." });

    const items = parseSharePointCSV(csvText);
    if (items.length === 0) return res.status(400).json({ error: "No valid fault rows found in CSV." });

    if (fs.existsSync(CSV_FILE)) {
      fs.writeFileSync(CSV_FILE, csvText, 'utf8');
    }
    fs.writeFileSync(DATA_FILE, JSON.stringify(items, null, 2), 'utf8');
    cachedFaults = items;

    console.log(`✅ Imported ${items.length} faults into RAM & disk!`);
    res.json({ success: true, count: items.length, items });
  } catch (err) {
    console.error("❌ CSV Import Error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.post('/chat', async (req, res) => {
  try {
    const userMessage = req.body.userInput || req.body.message;
    const sessionId = req.body.sessionId || "default-session";

    if (!userMessage) return res.status(400).json({ error: "Please enter a message." });
    if (req.body.reset) sessions.delete(sessionId);

    // 1. Deterministic Fault Creation: Immediately saves to right-hand table & pushes to Power Automate
    let createdFault = null;
    if (isFaultReportMessage(userMessage)) {
      const faultObj = buildFaultFromUserMessage(userMessage);
      createdFault = saveNewFault(faultObj);
      console.log(`✅ Saved New Fault ${createdFault.FaultID} to Right-Side Table!`);
      pushBotFaultToPowerAutomate(createdFault);
    }

    const rawText = await sendMessageFast(sessionId, userMessage, createdFault);

    // 2. Check for [VERIFIED: ...] tag
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

    // 3. Check for AI-generated [CREATE_FAULT: {...}] tag (cleanly initialized)
    const createRegex = /\[CREATE_FAULT:\s*(\{.*?\})\]/s;
    const aiCreateMatch = rawText.match(createRegex);
    if (!createdFault && aiCreateMatch) {
      try {
        const faultObj = JSON.parse(aiCreateMatch[1]);
        createdFault = saveNewFault(faultObj);
        console.log(`✅ Saved AI Fault ${createdFault.FaultID} to Right-Side Table!`);
        pushBotFaultToPowerAutomate(createdFault);
      } catch (e) {
        console.error("Failed to parse AI CREATE_FAULT JSON:", e.message);
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
    console.error("❌ Unexpected Chat Error:", error);
    res.json({
      response: generateInstantLocalResponse(req.body.userInput || "help"),
      allFaults: getSharePointFaults()
    });
  }
});

getSharePointFaults(true);

app.listen(port, () => {
  console.log(`✅ Fast Helper Server running on port ${port}`);
});