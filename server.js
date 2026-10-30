const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
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
if (!API_KEY) {
  console.error("❌ ERROR: GEMINI_API_KEY is missing in your .env file.");
}
const genAI = new GoogleGenerativeAI(API_KEY);

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

// IN-MEMORY CACHE: Zero disk lag during chat!
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

function getSharePointFaults(forceReload = false) {
  if (cachedFaults && !forceReload) {
    return cachedFaults;
  }

  try {
    if (fs.existsSync(CSV_FILE)) {
      const csvContent = fs.readFileSync(CSV_FILE, 'utf8');
      const parsed = parseSharePointCSV(csvContent);
      if (parsed.length > 0) {
        cachedFaults = parsed;
        fs.writeFileSync(DATA_FILE, JSON.stringify(parsed, null, 2), 'utf8');
        return cachedFaults;
      }
    }

    if (fs.existsSync(DATA_FILE)) {
      cachedFaults = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
      return cachedFaults;
    }
  } catch (e) {
    console.error("Error reading fault data:", e);
  }

  cachedFaults = DEFAULT_FAULTS;
  return cachedFaults;
}

function saveNewFault(newFault) {
  const faults = getSharePointFaults();
  faults.push(newFault);
  cachedFaults = faults;
  fs.writeFileSync(DATA_FILE, JSON.stringify(faults, null, 2), 'utf8');
  return newFault;
}

function getDynamicSystemInstruction() {
  const liveFaults = getSharePointFaults();
  const nextIdNumber = liveFaults.length + 1;
  const nextFaultID = `F${String(nextIdNumber).padStart(3, '0')}`;
  const compactFaults = JSON.stringify(liveFaults);

  return `You are Helper, the IT Support Assistant for Belgium Campus Equipment Fault Reporting.
Connected SharePoint "FaultReports" list: ${compactFaults}

RULES:
1. Fault Lookup: Use the JSON list above to answer questions about any FaultID, Title, EquipmentType, Description, Priority, or Status.
2. Verification: Before logging a new fault, ask for Full Name, Student/Staff Number, and Role. When provided, output once at top:
[VERIFIED: FullName | ID: Number | Role: RoleName]
3. Log Fault: When asked to log a fault, output this single-line JSON tag at the very bottom:
[CREATE_FAULT: {"Title": "Short Title", "FaultID": "${nextFaultID}", "EquipmentType": "Projector | Classroom PC | Wi-Fi/Network", "Description": "Issue details", "Priority": "High | Medium | Low", "Status": "Assigned"}]
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

// Instant Local Smart Fallback: Answers in <5ms using live SharePoint RAM data if Google API hits 429/503!
function generateInstantLocalResponse(userMessage) {
  const faults = getSharePointFaults();
  const nextId = `F${String(faults.length + 1).padStart(3, '0')}`;
  const msg = userMessage.toLowerCase();

  // 1. Check if user is providing verification details (Name + 5-8 digit Student/Staff ID)
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

  // 2. Check if user is asking about a specific FaultID (e.g., F001, F002, F003)
  const faultIdMatch = userMessage.match(/\b(F\d{3})\b/i);
  if (faultIdMatch && !msg.includes('log')) {
    const targetId = faultIdMatch[1].toUpperCase();
    const found = faults.find(f => f.FaultID.toUpperCase() === targetId);
    if (found) {
      return `${verifyTag}Here is the live SharePoint status for **${found.FaultID}**:\n* **Title:** ${found.Title}\n* **Equipment Type:** ${found.EquipmentType}\n* **Priority:** ${found.Priority}\n* **Status:** ${found.Status || 'Assigned'}\n* **Description:** ${found.Description}`;
    }
  }

  // 3. Check if user is asking to see all faults / check SharePoint list
  if (msg.includes('check') || msg.includes('what faults') || msg.includes('status') || msg.includes('list') || msg.includes('all faults')) {
    const listSummary = faults
      .map(f => `* **${f.FaultID} — ${f.Title}** (${f.EquipmentType}) | Priority: **${f.Priority}** | Status: **${f.Status || 'Assigned'}** — ${f.Description}`)
      .join('\n');
    return `${verifyTag}Here are the current fault reports synced from **SharePoint (FaultReports)**:\n\n${listSummary}\n\nWould you like details on a specific FaultID or need me to log a new fault?`;
  }

  // 4. Check if user wants to log/report a new fault
  if (msg.includes('log') || msg.includes('broken') || msg.includes('not working') || msg.includes('crash') || msg.includes('fault') || msg.includes('projector') || msg.includes('pc')) {
    const eqType = (msg.includes('pc') || msg.includes('computer') || msg.includes('lab')) ? 'Classroom PC'
                 : (msg.includes('wifi') || msg.includes('network') || msg.includes('internet')) ? 'Wi-Fi/Network'
                 : 'Projector';
    const priority = (msg.includes('high') || msg.includes('exam') || msg.includes('urgent') || msg.includes('crash')) ? 'High' : 'Medium';
    const title = eqType === 'Projector' ? 'Projector not working'
                : eqType === 'Classroom PC' ? 'PC not powering On / Crash'
                : 'Network connectivity fault';

    const newFaultObj = {
      Title: title,
      FaultID: nextId,
      EquipmentType: eqType,
      Description: userMessage.replace(/^➕\s*/i, '').replace(/^🚨\s*/i, '').trim(),
      Priority: priority,
      Status: 'Assigned'
    };

    return `${verifyTag}I have logged this **${priority} Priority** fault for the **${eqType}** into the FaultReports list:\n* **Quick 30s Check:** Please verify the power cable and display connection are firmly seated.\n* **Assigned FaultID:** **${nextId}**\n[CREATE_FAULT: ${JSON.stringify(newFaultObj)}]`;
  }

  // 5. General greeting / verified response
  if (verifyTag) {
    return `${verifyTag}Thank you for verifying your details! How can I assist you today?\n* Ask me to **check any FaultID** (like **F001** or **F002**).\n* Describe a broken **Projector**, **Classroom PC**, or **Wi-Fi** issue and I will log it into **FaultReports** for you.`;
  }

  return `Hi! I'm **Helper**, connected to your **SharePoint FaultReports** list (${faults.length} active faults loaded).\n* Ask me to check **F001**, **F002**, or all current faults.\n* Tell me your **Name & Student/Staff Number** or describe an equipment fault to log a new ticket!`;
}

async function sendMessageFast(sessionId, userMessage) {
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
        await delay(1200);
        continue;
      }
      break;
    }
  }

  console.log("🛡️ Google API rate-limited (429/503) — serving instant response from live SharePoint RAM cache!");
  return generateInstantLocalResponse(userMessage);
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

    const rawText = await sendMessageFast(sessionId, userMessage);

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

    let createdFault = null;
    const createRegex = /\[CREATE_FAULT:\s*(\{.*?\})\]/s;
    const createMatch = rawText.match(createRegex);
    if (createMatch) {
      try {
        const faultObj = JSON.parse(createMatch[1]);
        createdFault = saveNewFault(faultObj);
        console.log("✅ New Fault Added:", createdFault);
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
    console.error("❌ Unexpected Chat Error:", error);
    const fallbackText = generateInstantLocalResponse(req.body.userInput || "help");
    res.json({
      response: fallbackText,
      allFaults: getSharePointFaults()
    });
  }
});

// Preload faults into RAM on startup
getSharePointFaults(true);

app.listen(port, () => {
  console.log(`✅ Fast Helper Server running at http://localhost:${port}`);
});