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

// Allow up to 10MB so CSV uploads never get blocked
app.use(express.json({ limit: '10mb' }));
app.use(cors());
app.use(express.static(path.join(__dirname, 'public')));

const API_KEY = process.env.GEMINI_API_KEY;
if (!API_KEY) {
  console.error("❌ ERROR: GEMINI_API_KEY is missing in your .env file.");
}
const genAI = new GoogleGenerativeAI(API_KEY);

// Default initial data matching your SharePoint list
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

// Parses raw SharePoint CSV text (handles UTF-8 BOM, commas, semicolons, and quotes)
function parseSharePointCSV(csvText) {
  const cleanText = csvText.replace(/^\uFEFF/, '');
  const lines = cleanText.split(/\r?\n/).filter(line => line.trim() !== '');
  if (lines.length < 2) return [];

  // Auto-detect comma vs semicolon delimiter
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

  // Normalize header names (lowercase, remove spaces/underscores) for flexible matching
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

// Reads faults from FaultReports.csv (if dropped in folder) or sharepoint_data.json
function getSharePointFaults() {
  try {
    // 1. If user saved FaultReports.csv in the root folder, parse it automatically
    if (fs.existsSync(CSV_FILE)) {
      const csvContent = fs.readFileSync(CSV_FILE, 'utf8');
      const parsed = parseSharePointCSV(csvContent);
      if (parsed.length > 0) {
        fs.writeFileSync(DATA_FILE, JSON.stringify(parsed, null, 2), 'utf8');
        return parsed;
      }
    }

    // 2. Otherwise read sharepoint_data.json (updated whenever you upload a CSV in the UI)
    if (fs.existsSync(DATA_FILE)) {
      return JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
    }
  } catch (e) {
    console.error("Error reading fault data:", e);
  }
  return DEFAULT_FAULTS;
}

function saveNewFault(newFault) {
  const faults = getSharePointFaults();
  faults.push(newFault);
  fs.writeFileSync(DATA_FILE, JSON.stringify(faults, null, 2), 'utf8');
  return newFault;
}

function getDynamicSystemInstruction() {
  const liveFaults = getSharePointFaults();
  const nextIdNumber = liveFaults.length + 1;
  const nextFaultID = `F${String(nextIdNumber).padStart(3, '0')}`;

  return `
You are Helper, the intelligent First-Line IT & Facilities Support Assistant for the Belgium Campus Pretoria Equipment Fault Reporting & Tracking System.
You are connected to the SharePoint List: "FaultReports".

### 1. CURRENT SHAREPOINT "FaultReports" LIST:
${JSON.stringify(liveFaults, null, 2)}

### 2. YOUR CAPABILITIES:
- **Fault Status Lookup**: When a user asks about existing faults (e.g., "What faults are logged?", "Check F001", "Check F003"), read the CURRENT SHAREPOINT LIST above and report the exact Title, FaultID, EquipmentType, Description, Priority, and Status.
- **User Verification**: Before logging a new fault, politely ask for the user's **Full Name**, **Student/Staff Number**, and **Role** (Student, Lecturer, or Staff). If they are in a live exam/test crisis, give a 30-second fix first. When they provide their Name and ID, include this exact tag once at the top of your reply:
  [VERIFIED: FullName | ID: Number | Role: RoleName]
- **Logging a New Fault**: If an issue needs a technician and the user asks you to log it, include this exact single-line JSON tag at the very bottom of your reply so the table updates immediately:
  [CREATE_FAULT: {"Title": "Short Title", "FaultID": "${nextFaultID}", "EquipmentType": "Projector | Classroom PC | Wi-Fi/Network", "Description": "Clear description of fault", "Priority": "High | Medium | Low", "Status": "Assigned"}]

Keep responses concise, friendly, and formatted with bullet points and bold text (under 150 words).
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

// Get current faults
app.get('/api/faults', (req, res) => {
  res.json(getSharePointFaults());
});

// Upload & parse SharePoint CSV from the browser button
app.post('/api/import-sharepoint-csv', (req, res) => {
  try {
    const csvText = req.body.csv;
    if (!csvText) {
      return res.status(400).json({ error: "No CSV text provided." });
    }

    const items = parseSharePointCSV(csvText);
    if (items.length === 0) {
      return res.status(400).json({ error: "No valid fault rows found in CSV." });
    }

    // Save to sharepoint_data.json (and remove old FaultReports.csv if present so JSON takes priority)
    if (fs.existsSync(CSV_FILE)) {
      fs.writeFileSync(CSV_FILE, csvText, 'utf8');
    }
    fs.writeFileSync(DATA_FILE, JSON.stringify(items, null, 2), 'utf8');

    console.log(`✅ Imported ${items.length} faults from SharePoint CSV!`);
    res.json({ success: true, count: items.length, items });
  } catch (err) {
    console.error("❌ CSV Import Error:", err);
    res.status(500).json({ error: err.message });
  }
});

// Chat endpoint
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

    // 2. Check for [CREATE_FAULT: {...}] tag
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
    console.error("Gemini API Error:", error);
    res.status(500).json({
      error: "Google's AI server is momentarily busy. Please wait 3 seconds and click Send again."
    });
  }
});

app.listen(port, () => {
  console.log(`✅ Helper Server running at http://localhost:${port}`);
});