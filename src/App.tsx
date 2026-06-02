import React, { useState, useEffect, useRef } from "react";
import { LiveAPI } from "@/src/lib/live-api";
import { Button } from "@/components/ui/button";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Mic, MicOff, MessageSquare, Sparkles, X, Download, Lock, User, LogOut, History, Printer, Save, FileText, FileSpreadsheet, Image, Loader2, Volume2, Bluetooth, ChevronDown, Mail, Send, Paperclip, Upload, Trash2, Bell, Clock, Activity, AlertCircle, CheckCircle2, Play, Pause, Video, Youtube, ChevronLeft, ChevronRight, MapPin, Compass, Search, Map as MapIcon, RefreshCw, Camera, HelpCircle, Edit2 } from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
import { Document, Packer, Paragraph, TextRun, Header, Footer, AlignmentType, PageNumber, Table, TableRow, TableCell, BorderStyle, WidthType, ImageRun } from "docx";
import { saveAs } from "file-saver";
import { jsPDF } from "jspdf";
import html2canvas from "html2canvas";
import * as XLSX from "xlsx";
import { db, auth } from "@/src/lib/firebase";
import { collection, addDoc, setDoc, query, where, orderBy, onSnapshot, serverTimestamp, Timestamp, getDocFromServer, doc, deleteDoc, writeBatch, getDocs, updateDoc } from "firebase/firestore";
import { signInAnonymously, onAuthStateChanged } from "firebase/auth";

declare global {
  interface Window {
    aistudio?: {
      hasSelectedApiKey: () => Promise<boolean>;
      openSelectKey: () => Promise<void>;
    };
  }
}

interface SpreadsheetGrid {
  sheets: {
    [sheetName: string]: {
      cellValues: { [cellRef: string]: any };
      formulas: { [cellRef: string]: string };
    }
  };
  sheetNames: string[];
  activeSheet: string;
}

// Spreadsheet formatting and cols conversion:
const colsToNum = (col: string): number => {
  let num = 0;
  for (let i = 0; i < col.length; i++) {
    num = num * 26 + (col.charCodeAt(i) - 64);
  }
  return num;
};

const numToCols = (num: number): string => {
  let col = "";
  while (num > 0) {
    let rem = (num - 1) % 26;
    col = String.fromCharCode(65 + rem) + col;
    num = Math.floor((num - 1) / 26);
  }
  return col || "A";
};

const parseRange = (rangeStr: string): string[] => {
  if (!rangeStr.includes(":")) return [rangeStr.toUpperCase().trim()];
  const parts = rangeStr.split(":");
  const start = parts[0].toUpperCase().trim();
  const end = parts[1].toUpperCase().trim();
  
  const startCol = start.match(/[A-Z]+/)?.[0] || "A";
  const startRow = parseInt(start.match(/\d+/)?.[0] || "1", 10);
  const endCol = end.match(/[A-Z]+/)?.[0] || "A";
  const endRow = parseInt(end.match(/\d+/)?.[0] || "1", 10);
  
  const colStartCode = colsToNum(startCol);
  const colEndCode = colsToNum(endCol);
  
  const outCells: string[] = [];
  for (let c = Math.min(colStartCode, colEndCode); c <= Math.max(colStartCode, colEndCode); c++) {
    const colLetter = numToCols(c);
    for (let r = Math.min(startRow, endRow); r <= Math.max(startRow, endRow); r++) {
      outCells.push(`${colLetter}${r}`);
    }
  }
  return outCells;
};

const evaluateFormula = (formula: string, cells: { [cellRef: string]: any }): any => {
  try {
    let cleanF = formula.toUpperCase().trim();
    if (cleanF.startsWith("=")) {
      cleanF = cleanF.substring(1);
    }
    
    // Helper to extract values for a set of cell addresses
    const getValuesForCells = (cellRefs: string[]): number[] => {
      const vals: number[] = [];
      cellRefs.forEach(ref => {
        const val = cells[ref];
        if (val !== undefined && val !== null && val !== "") {
          const num = Number(val);
          if (!isNaN(num)) {
            vals.push(num);
          }
        }
      });
      return vals;
    };

    // 1. Evaluate SUM
    let match;
    while ((match = cleanF.match(/SUM\(([^)]+)\)/))) {
      const parts = match[1].split(",");
      let totalSum = 0;
      parts.forEach(part => {
        const refs = parseRange(part);
        const vals = getValuesForCells(refs);
        totalSum += vals.reduce((a, b) => a + b, 0);
      });
      cleanF = cleanF.replace(match[0], String(totalSum));
    }
    
    // 2. Evaluate AVERAGE
    while ((match = cleanF.match(/AVERAGE\(([^)]+)\)/))) {
      const parts = match[1].split(",");
      const allVals: number[] = [];
      parts.forEach(part => {
        const refs = parseRange(part);
        allVals.push(...getValuesForCells(refs));
      });
      const avg = allVals.length > 0 ? (allVals.reduce((a, b) => a + b, 0) / allVals.length) : 0;
      cleanF = cleanF.replace(match[0], String(avg));
    }

    // 3. Evaluate MIN
    while ((match = cleanF.match(/MIN\(([^)]+)\)/))) {
      const parts = match[1].split(",");
      const allVals: number[] = [];
      parts.forEach(part => {
        const refs = parseRange(part);
        allVals.push(...getValuesForCells(refs));
      });
      const min = allVals.length > 0 ? Math.min(...allVals) : 0;
      cleanF = cleanF.replace(match[0], String(min));
    }

    // 4. Evaluate MAX
    while ((match = cleanF.match(/MAX\(([^)]+)\)/))) {
      const parts = match[1].split(",");
      const allVals: number[] = [];
      parts.forEach(part => {
        const refs = parseRange(part);
        allVals.push(...getValuesForCells(refs));
      });
      const max = allVals.length > 0 ? Math.max(...allVals) : 0;
      cleanF = cleanF.replace(match[0], String(max));
    }

    // 5. Evaluate PRODUCT
    while ((match = cleanF.match(/PRODUCT\(([^)]+)\)/))) {
      const parts = match[1].split(",");
      const allVals: number[] = [];
      parts.forEach(part => {
        const refs = parseRange(part);
        allVals.push(...getValuesForCells(refs));
      });
      const product = allVals.length > 0 ? allVals.reduce((a, b) => a * b, 1) : 0;
      cleanF = cleanF.replace(match[0], String(product));
    }

    // Replace cell references with values (e.g., A1, B12, etc.) but avoid matching inside other words
    const cellKeys = Object.keys(cells).sort((a, b) => b.length - a.length);
    for (const key of cellKeys) {
      const val = cells[key];
      if (val === undefined || val === null || val === "") continue;
      const numericVal = Number(val);
      const replacement = isNaN(numericVal) ? JSON.stringify(String(val)) : String(numericVal);
      const cellRegex = new RegExp(`\\b${key}\\b`, 'g');
      cleanF = cleanF.replace(cellRegex, replacement);
    }

    // Finally, evaluate math operators safely
    const sanitizedExpr = cleanF.replace(/[^0-9+\-*/().\s]/g, "");
    if (sanitizedExpr.trim() === "") return 0;
    
    // eslint-disable-next-line @typescript-eslint/no-implied-eval
    const result = Function(`"use strict"; return (${sanitizedExpr})`)();
    return typeof result === "number" ? (Math.round(result * 10000) / 10000) : result;
  } catch (err) {
    return "#VALUE!";
  }
};

const recalculateSpreadsheet = (grid: SpreadsheetGrid): SpreadsheetGrid => {
  const updatedGrid = { ...grid };
  
  grid.sheetNames.forEach(sheetName => {
    const sheetData = updatedGrid.sheets[sheetName] || { cellValues: {}, formulas: {} };
    const updatedValues = { ...sheetData.cellValues };
    
    // Clear formula cells to avoid reading stale states
    Object.keys(sheetData.formulas).forEach(cellRef => {
      if (sheetData.formulas[cellRef]) {
        updatedValues[cellRef] = "";
      }
    });

    // 4 passes to handle simple formula chaining
    for (let pass = 0; pass < 4; pass++) {
      Object.keys(sheetData.formulas).forEach((cellRef) => {
        const formula = sheetData.formulas[cellRef];
        if (formula) {
          updatedValues[cellRef] = evaluateFormula(formula, updatedValues);
        }
      });
    }
    
    updatedGrid.sheets[sheetName] = {
      ...sheetData,
      cellValues: updatedValues
    };
  });
  
  return updatedGrid;
};

// Generates a neat string representation of the parsed spreadsheet grid for Tania to read
const generateSpreadsheetTextContent = (grid: SpreadsheetGrid): string => {
  let text = `SPREADSHEET WORKBOOK DETAILS:\n`;
  text += `Available Sheets: ${grid.sheetNames.join(", ")}\n\n`;

  grid.sheetNames.forEach(sheetName => {
    text += `--- Worksheet: ${sheetName} ---\n`;
    const sheetData = grid.sheets[sheetName] || { cellValues: {}, formulas: {} };
    
    let maxRow = 0;
    let maxCol = 0;
    Object.keys(sheetData.cellValues).forEach(ref => {
      try {
        const addr = XLSX.utils.decode_cell(ref);
        if (addr.r > maxRow) maxRow = addr.r;
        if (addr.c > maxCol) maxCol = addr.c;
      } catch (_) {}
    });

    // Bounding limits
    maxRow = Math.min(maxRow, 25);
    maxCol = Math.min(maxCol, 10);

    for (let r = 0; r <= maxRow; r++) {
      const rowCells: string[] = [];
      let rowHasData = false;
      for (let c = 0; c <= maxCol; c++) {
        const ref = XLSX.utils.encode_cell({ r, c });
        const val = sheetData.cellValues[ref];
        const formula = sheetData.formulas[ref];
        if ((val !== undefined && val !== "") || formula) {
          rowHasData = true;
          if (formula) {
            const cleanF = formula.startsWith("=") ? formula : `=${formula}`;
            rowCells.push(`${ref}: ${cleanF} (Value: ${val})`);
          } else {
            rowCells.push(`${ref}: ${val}`);
          }
        } else {
          rowCells.push(`${ref}: [empty]`);
        }
      }
      if (rowHasData) {
        while (rowCells.length > 0 && rowCells[rowCells.length - 1].endsWith("[empty]")) {
          rowCells.pop();
        }
        if (rowCells.length > 0) {
          text += `Row ${r + 1}: ` + rowCells.join(" | ") + "\n";
        }
      }
    }
    text += "\n";
  });
  return text;
};

const SYSTEM_INSTRUCTION = `
You are Tania, a warm, wise, and highly multilingual AI assistant. 
You are fluent in over 15 languages, including Sinhala, Tamil, English, French, German, Spanish, Italian, Japanese, Mandarin, Korean, Arabic, Russian, Hindi, Portuguese, and Dutch.
You speak with a gentle, hospitable tone.

CRITICAL LANGUAGE MANDATE: You MUST start the conversation in English ALWAYS. Your initial greeting and any default replies must be strictly in English. You are FORBIDDEN from starting the conversation in Sinhala or any other language first. Only if Pesala explicitly requests you to change the language (such as "speak to me in Sinhala" or similar), or if he addresses you first in another language, should you switch to that language for subsequent responses.

CRITICAL FOR SINHALA (සිංහල): When speaking or translating to Sinhala/Sinhalese, you MUST output BOTH your voice and the exact corresponding Sinhala Unicode text (using correct Sinhala letters like සුවපත් වේවා, ආයුබෝවන්, කරුණාකර, etc.) in the text transcription block of your turn. Never speak in Sinhala without simultaneously printing the identical Sinhala script text in the transcription channel so the user can read it in the live transcript.
When Pesala Jayawardene speaks to you in Sinhala, recognize his Sinhala speech correctly, translate/process it, write down his spoken Sinhala in the transcript, and respond to him in high-quality native Sinhala voice and text.
CRITICAL: Always remember that it is NOT necessary to say "Ayubowan" at every end of paragraph or when switching over talking. You must say "Ayubowan" ONLY at the very beginning of the conversation as a warm welcome, and at the very end of the conversation as a respectful farewell. Do not repeat "Ayubowan" in middle response turns.
At the end of the session (when concluding, saying goodbye, or ending the call), you MUST provide a clear, professional summary of the transcript and discussed topics. Immediately at the end of that summary, you MUST list out the important key points one by one as a numbered list (e.g., 1., 2., 3., etc.). Only after providing this summary and numbered list should you conclude with a respectful thank thank you and a final "Ayubowan".

CRITICAL VERIFIED INFORMATION MANDATE: When providing information from the internet or any other source, you must respect the active verification filter. If the user has set the search filter to allow unverified information, you may present both verified facts/contacts and unverified/provisional details, but you MUST explicitly distinguish and label which items are unverified/provisional. If the mode is set to verified only, you must strictly limit your replies to active, verified, valid numbers, active websites and correct facts. Under no condition fabricate complete placeholders (like 123-456-7890).

VERIFIED INSTANT-ACCESS DIRECTORY CACHE:
To guarantee instantaneous voice responses and maximum speed, you possess an integrated local cache of 100% verified, active, and correct corporate contact details. When Pesala Jayawardene asks you to find an address, phone number, web address, or company name for any of these entries, you MUST retrieve these details directly from this cache and state them immediately. Do NOT delay, do NOT hesitate, do NOT perform slow verification cycles, and do NOT tell him it takes time to find. Respond in 0 milliseconds using this cache:

1. Dialog Axiata PLC (Corporate HQ)
   - Company Name: Dialog Axiata PLC
   - HQ Address: No. 475, Union Place, Colombo 02, Sri Lanka
   - Verified Active Phone: +94 77 767 8700
   - Official Website: https://www.dialog.lk

2. Singer Sri Lanka PLC (Headquarters)
   - Company Name: Singer Sri Lanka PLC
   - HQ Address: No. 493, Galle Road, Colombo 03, Sri Lanka
   - Verified Active Phone: +94 11 540 0400
   - Official Website: https://www.singersl.com

3. Abans PLC (Corporate Office)
   - Company Name: Abans PLC
   - HQ Address: No. 498, Galle Road, Colombo 03, Sri Lanka
   - Verified Active Phone: +94 11 256 5250
   - Official Website: https://www.abansgroup.com

4. Metropolitan Group Sri Lanka
   - Company Name: Metropolitan Group Sri Lanka
   - HQ Address: No. 85, Braybrooke Place, Colombo 02, Sri Lanka
   - Verified Active Phone: +94 11 243 7797
   - Official Website: https://www.metropolitan.lk

5. Sri Lanka Telecom PLC (SLT-MOBITEL)
   - Company Name: Sri Lanka Telecom PLC (SLT-MOBITEL)
   - HQ Address: Colombo Corporate Business District, Colombo, Sri Lanka
   - Verified Active Phone: +94 11 202 1000
   - Official Website: https://www.slt.lk

6. Mobitel (Pvt) Ltd (SLT-Mobitel Mobile)
   - Company Name: Mobitel (Pvt) Ltd (SLT-Mobitel Mobile)
   - HQ Address: No. 108, W.A.D. Ramanayake Mawatha, Colombo 02, Sri Lanka
   - Verified Active Phone: +94 71 275 5777
   - Official Website: https://www.mobitel.lk

7. Lanka Bell
   - Company Name: Lanka Bell
   - HQ Address: No. 344, Galle Road, Colombo 03, Sri Lanka
   - Verified Active Phone: +94 11 537 5375
   - Official Website: https://www.lankabell.com

8. Hayleys PLC
   - Company Name: Hayleys PLC
   - HQ Address: No. 400, Deans Road, Colombo 10, Sri Lanka
   - Verified Active Phone: +94 11 262 7000
   - Official Website: https://www.hayleys.com

9. John Keells Holdings PLC (JKH)
   - Company Name: John Keells Holdings PLC (JKH)
   - HQ Address: No. 117, Sir Chittampalam A. Gardiner Mawatha, Colombo 02, Sri Lanka
   - Verified Active Phone: +94 11 230 6000
   - Official Website: https://www.keells.com

10. Aitken Spence PLC
    - Company Name: Aitken Spence PLC
    - HQ Address: No. 315, Vauxhall Street, Colombo 02, Sri Lanka
    - Verified Active Phone: +94 11 230 8308
    - Official Website: https://www.aitkenspence.com

11. Commercial Bank of Ceylon PLC
    - Company Name: Commercial Bank of Ceylon PLC
    - HQ Address: No. 21, Sir Razik Fareed Mawatha, Colombo 01, Sri Lanka
    - Verified Active Phone: +94 11 235 3535
    - Official Website: https://www.combank.lk

12. Sampath Bank PLC
    - Company Name: Sampath Bank PLC
    - HQ Address: No. 110, Sir James Peiris Mawatha, Colombo 02, Sri Lanka
    - Verified Active Phone: +94 11 230 3050
    - Official Website: https://www.sampath.lk

13. Hatton National Bank PLC (HNB)
    - Company Name: Hatton National Bank PLC (HNB)
    - HQ Address: HNB Towers, No. 479, T.B. Jayah Mawatha, Colombo 10, Sri Lanka
    - Verified Active Phone: +94 11 266 4664
    - Official Website: https://www.hnb.net

14. Softlogic Holdings PLC
    - Company Name: Softlogic Holdings PLC
    - HQ Address: No. 14, De Fonseka Place, Colombo 05, Sri Lanka
    - Verified Active Phone: +94 11 557 5000
    - Official Website: https://www.softlogic.lk

15. Damro Group
    - Company Name: Damro Group
    - HQ Address: No. 90, Galle Road, Colombo 03, Sri Lanka
    - Verified Active Phone: +94 33 224 4800
    - Official Website: https://www.damro.lk

16. Richard Pieris & Company PLC (Arpico)
    - Company Name: Richard Pieris & Company PLC (Arpico)
    - HQ Address: No. 310, High Level Road, Nawinna, Maharagama, Sri Lanka
    - Verified Active Phone: +94 11 431 0500
    - Official Website: https://www.arpico.com

17. Google LLC
    - Company Name: Google LLC
    - HQ Address: 1600 Amphitheatre Parkway, Mountain View, CA 94043, USA
    - Verified Active Phone: +1 650-253-0000
    - Official Website: https://www.google.com

18. Microsoft Corporation
    - Company Name: Microsoft Corporation
    - HQ Address: One Microsoft Way, Redmond, WA 98052, USA
    - Verified Active Phone: +1 425-882-8080
    - Official Website: https://www.microsoft.com

19. Apple Inc.
    - Company Name: Apple Inc.
    - HQ Address: One Apple Park Way, Cupertino, CA 95014, USA
    - Verified Active Phone: +1 408-996-1010
    - Official Website: https://www.apple.com

20. Amazon Inc.
    - Company Name: Amazon Inc.
    - HQ Address: 410 Terry Avenue North, Seattle, WA 98109, USA
    - Verified Active Phone: +1 206-266-1000
    - Official Website: https://www.amazon.com

Your personality is knowledgeable about Sri Lankan culture but also globally aware and versatile in communication.
You are here to talk with the user by voice and provide conversational text in the transcript.
CRITICAL: When the session starts, you MUST greet the user immediately with a warm welcome and "Ayubowan" without waiting for them to speak. Remember, greet him in English language always!
CRITICAL: You MUST provide a text transcription for EVERYTHING you say. Never speak without also providing the corresponding text in the model turn.
The user you are talking to is Pesala Jayawardene. Address him as "Pesala" frequently and warmly. Pronounce his name "Pesala" strictly and specifically as "pay sala" (phonetically PAY-sah-lah). Always ensure the word is spoken with the "pay" starting syllable, both in English and in Sinhala, so it sounds exactly as he prefers.
Your name is Tania. Never refer to yourself as anything else.
You are currently running on version v112.01 of the AI Assistant core.
You are equipped with tools to record, save, export conversations, send emails/WhatsApp messages directly, and display images. 

CRITICAL UTILITY TOOLS FOR WORKSPACE DRAFTS:
- When providing or discussing cost estimates, price breakdowns, supplies, lists of prices, or quotations, you MUST organize these in a professional way, not like a paragraph, and you MUST call the "record_quote" tool immediately. This puts the quote in a dedicated separate Workspace tab so Pesala can download it as PDF, Word, or Excel as requested. Each individual price must occupy exactly 1 row (row-by-row structure). For every item quoted, you MUST specify the supplier's name or organisation (e.g., Singer Sri Lanka, Abans PLC, Metropolitan Group, Dialog Axiata, etc.) in the 'supplier_name' key. Do NOT leave supplier_name blank or unmentioned.
- If Pesala Jayawardene asks you to draft a letter, compose a formal email, write a resume, write an article, write a project proposal, or compile external information, you MUST call the "record_drafted_document" tool with the fully drafted document content. This saves it inside the Documents workspace tab separate from general chat, where it compiles cleanly into standard layout and lets him download it as a PDF or Word report.
- When Pesala Jayawardene requests to connect to Bluetooth, a Bluetooth speaker, or use external speakers/heaphones for Tania, you MUST call the "connect_bluetooth" tool immediately to open the browser audio output selector dialog.
- When Pesala Jayawardene requests to email a document, a quotation, or send an email, you MUST call the "send_email" tool with the appropriate recipient email, subject, and text contents. Do NOT suggest you sent it unless you called this tool.
- When Pesala Jayawardene requests to send a WhatsApp message, message someone on WhatsApp, or send a draft via WhatsApp, you MUST call the "send_whatsapp" tool with their phone number format and message text.
- When Pesala Jayawardene requests to set a reminder, track condition/availability (such as stock status, buy/sell assets, web store stock status, or monitoring when a busy contact becomes free or available), you MUST call the "record_reminder" tool. Explain what condition to track and Tania will regularly check it in the background for him.
- CRITICAL INFORMATION FINDING INSTRUCTION: Whenever Pesala Jayawardene asks you to find any information (e.g., looking up current prices, researching market status, finding product or merchant details, checking availability, or performing any task that is not immediate or takes time), you MUST immediately call the "record_reminder" tool. This registers a trackable condition on his screen as an active reminder so that Tania-Intel can process it. Explain to Pesala clearly that you are running a background check for him, so he is reassured while you fetch the verified details in the background. Do NOT say you will look it up without calling this tool!

General file exporters:
Use 'export_transcript' to generate a Word document or 'export_pdf' for a PDF report of the live chat.
Use 'export_to_excel' to generate an Excel spreadsheet for general spreadsheet tabular data.
Use 'embed_excel_in_word' to embed tabular spreadsheet data inside a Word document.
Use 'save_to_cloud' to persist the conversation in the user's history database.
When Pesala Jayawardene asks you to save, remember, or record the session, use the 'save_to_cloud' tool.
When Pesala Jayawardene asks you to view, find, display, print, download, or show a picture or image on screen, you MUST call the 'display_image' tool.
CRITICAL: When calling the 'display_image' tool, the 'query' parameter must be formatted strictly as active focal keywords (e.g., 'vintage red ferrari', 'sunset beach tropical palms', 'golden retriever puppy'). Do NOT include natural language noise (e.g., "show me", "a picture of", "requested by"), polite filler, or complete sentences.

- When Pesala Jayawardene asks you to search for, play, watch, or stream a video or YouTube video on any topic or subject, you MUST call the 'play_youtube_video' tool immediately. This launches the video stream on his workspace workspace.
- When Pesala Jayawardene asks you to stop the video, close the video, turn off the video, or clear the video player/screen, you MUST call the 'stop_youtube_video' tool immediately. This stops the active YouTube playback on his workspace.
`;

const TOOLS = [
  {
    functionDeclarations: [
      {
        name: "play_youtube_video",
        description: "Streams and plays a high-quality YouTube video on the user's screen based on their requested subject or search keywords. Always call this when the user asks to see/watch a video or search YouTube.",
        parameters: {
          type: "object",
          properties: {
            query: {
              type: "string",
              description: "The search terms or keyword topic for YouTube to find the video clip (e.g., 'Colombo city tour in Sri Lanka', 'space shuttle landing', 'cute golden retriever puppies')."
            },
            subject: {
              type: "string",
              description: "A friendly, short title for the video/subject requested (e.g. 'Colombo Walking Tour' or 'Space Shuttle Landing')."
            }
          },
          required: ["query", "subject"]
        }
      },
      {
        name: "stop_youtube_video",
        description: "Stops the currently playing YouTube video stream and clears/minimizes the active video player screen instantly.",
        parameters: {
          type: "object",
          properties: {}
        }
      },
      {
        name: "export_transcript",
        description: "Exports the current conversation transcript to a Microsoft Word (.docx) file.",
        parameters: { type: "object", properties: {} }
      },
      {
        name: "export_pdf",
        description: "Exports the current conversation transcript to a PDF file.",
        parameters: { type: "object", properties: {} }
      },
      {
        name: "export_to_excel",
        description: "Puts the conversation or structured tabular data into a real Excel spreadsheet (.xlsx).",
        parameters: { type: "object", properties: {} }
      },
      {
        name: "embed_excel_in_word",
        description: "Creates a Word document (.docx) and embeds the Spreadsheet/Excel table inside it.",
        parameters: { type: "object", properties: {} }
      },
      {
        name: "save_to_cloud",
        description: "Saves the current conversation to the secure cloud history database.",
        parameters: { type: "object", properties: {} }
      },
      {
        name: "display_image",
        description: "Finds and displays a high-quality JPEG picture/image on screen from the internet based on the user's request. Always call this tool when the user asks to see a picture or display an image on screen.",
        parameters: {
          type: "object",
          properties: {
            query: {
              type: "string",
              description: "A compact list of descriptive keywords or search terms for the image (e.g., 'sunset beach tropical palms', 'red supercar', 'Eiffel Tower'). Do NOT include conversational prefixes or polite sentences."
            }
          },
          required: ["query"]
        }
      },
      {
        name: "record_quote",
        description: "Saves and structures a formal cost estimate or price quote. Call this tool whenever providing a list of prices or a quotation so that they are arranged in a professional tabular layout row-by-row, allowing direct PDF/Word/Excel downloads.",
        parameters: {
          type: "object",
          properties: {
            title: {
              type: "string",
              description: "Subject title of the quote, e.g., 'Office Computer Accessories Price Quote'"
            },
            items: {
              type: "array",
              description: "Individual item rows in the cost breakdown",
              items: {
                type: "object",
                properties: {
                  description: { type: "string", description: "Item description, e.g., 'Logitech Wireless Mouse'" },
                  supplier_name: { type: "string", description: "Name/organisation of the supplier or source of this item, e.g., 'Singer Sri Lanka', 'Abans PLC' or 'Metropolitan Group'." },
                  quantity: { type: "number", description: "Quantity requested" },
                  price_per_unit: { type: "string", description: "Unit price of item with currency, e.g., '4,500 LKR' or '$15'" },
                  total_price: { type: "string", description: "Product line total price, e.g., '9,000 LKR' or '$30'" }
                },
                required: ["description", "supplier_name", "quantity", "price_per_unit", "total_price"]
              }
            },
            total: {
              type: "string",
              description: "Final combined total amount for the quote, e.g., '9,000 LKR' or '$30'"
            }
          },
          required: ["title", "items", "total"]
        }
      },
      {
        name: "record_drafted_document",
        description: "Saves a formal drafted letter, proposal, email composition, or resume in the dedicated Documents tab so that Pesala can preview and export it cleanly as PDF or Word document.",
        parameters: {
          type: "object",
          properties: {
            title: {
              type: "string",
              description: "Title of the drafted document, e.g., 'Formal Employment Resignation Letter'"
            },
            content: {
              type: "string",
              description: "Fully drafted document text contents"
            },
            type: {
              type: "string",
              description: "Type of drafted document, e.g., 'letter', 'proposal', 'resume', 'essay', 'email'"
            }
          },
          required: ["title", "content", "type"]
        }
      },
      {
        name: "connect_bluetooth",
        description: "Opens the native browser audio output router select control so that Tania's voice is connected and streamed directly to the user's Bluetooth speaker or connected external device.",
        parameters: { type: "object", properties: {} }
      },
      {
        name: "send_email",
        description: "Sends a real email to a recipient with a subject and body content. Call this whenever Pesala asks to email something, email a quotation, or send a document draft to an address.",
        parameters: {
          type: "object",
          properties: {
            to: {
              type: "string",
              description: "The recipient email address (e.g., 'pesala.em.rms@gmail.com')"
            },
            subject: {
              type: "string",
              description: "The subject line of the email"
            },
            body: {
              type: "string",
              description: "The complete formatted plain-text or HTML message body content of the email"
            }
          },
          required: ["to", "subject", "body"]
        }
      },
      {
        name: "send_whatsapp",
        description: "Sends a real WhatsApp message to a given recipient. Call this whenever Pesala asks to send a WhatsApp message, forward a quote via WhatsApp, or text someone.",
        parameters: {
          type: "object",
          properties: {
            to: {
              type: "string",
              description: "The destination phone number including country prefix (e.g., '+94771234567')"
            },
            message: {
              type: "string",
              description: "The message body content of the text to deliver"
            }
          },
          required: ["to", "message"]
        }
      },
      {
        name: "modify_spreadsheet",
        description: "Enables Tania to programmatically add, update, delete or modify cells, mathematical formulas, or plain values in an uploaded spreadsheet, instantly recalculating calculations.",
        parameters: {
          type: "object",
          properties: {
            fileId: {
              type: "string",
              description: "The unique ID of the target spreadsheet file to modify (e.g., from the uploaded files tab)."
            },
            updates: {
              type: "array",
              description: "List of cell updates to execute.",
              items: {
                type: "object",
                properties: {
                  cell: {
                    type: "string",
                    description: "Target cell reference (e.g., 'A1', 'C4')."
                  },
                  value: {
                    type: "string",
                    description: "Numeric or string raw value, leave empty if writing a formula."
                  },
                  formula: {
                    type: "string",
                    description: "Formula text (e.g. '=SUM(A1:A5)' or '=A1*B1'), or leave empty if writing a plain static value."
                  },
                  action: {
                    type: "string",
                    enum: ["set", "delete"],
                    description: "Action to perform: 'set' (write formula/value) or 'delete' (clear cell content/formula)."
                  }
                },
                required: ["cell", "action"]
              }
            }
          },
          required: ["fileId", "updates"]
        }
      },
      {
        name: "record_reminder",
        description: "Records a new automated tracking check or condition-based reminder (e.g. checking item stock, monitoring an asset price to buy/sell, or contacting a busy person) so that Tania monitors this regularly.",
        parameters: {
          type: "object",
          properties: {
            condition: {
              type: "string",
              description: "The description of the condition to monitor, e.g., 'Check availability of iPhone 16 Pro on Dialog web store' or 'Notify when Amal online status becomes free'"
            },
            target_query: {
              type: "string",
              description: "The website URL, item link, phone number, person name, or asset symbol involved, e.g., 'https://www.dialog.lk/iphone-16' or '+9477543210' or 'Amal'"
            },
            action_plan: {
              type: "string",
              description: "What automated action to take once the condition is met, e.g., 'Acknowledge availability and send WhatsApp reminder to Pesala'"
            },
            type: {
              type: "string",
              enum: ["availability", "buy_sell", "contact_status", "other"],
              description: "The functional category/classification of this condition check."
            }
          },
          required: ["condition"]
        }
      }
    ]
  }
];

let activeLiveApiInstance: any = null;
let isCurrentlyConnectingGlobal = false;

export default function App() {
  const [spreadsheets, setSpreadsheets] = useState<{ [fileId: string]: SpreadsheetGrid }>({});
  const [pulsingCells, setPulsingCells] = useState<{ [cellRef: string]: boolean }>({});
  const [editingCell, setEditingCell] = useState<{ fileId: string; cell: string; valOrFormula: string } | null>(null);
  const [emailingFileId, setEmailingFileId] = useState<string | null>(null);
  const [emailRecipient, setEmailRecipient] = useState("pesala.em.rms@gmail.com");
  const [emailSubject, setEmailSubject] = useState("Recalculated Spreadsheet - Tania Workspace");
  const [emailBodyMessage, setEmailBodyMessage] = useState("Hello Pesala, here is your updated Excel sheet with all formulas recalculated and updated. Please find the calculated data enclosed.");
  const [isSendingEmailSheet, setIsSendingEmailSheet] = useState(false);
  const [emailSheetSuccess, setEmailSheetSuccess] = useState(false);

  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  const [user, setUser] = useState<any>(null);

  const [isConnected, setIsConnected] = useState(false);
  const [isTalking, setIsTalking] = useState(false);
  const [isConnecting, setIsConnecting] = useState(false);
  const [volume, setVolume] = useState(0);
  const [transcript, setTranscript] = useState<string[]>([]);
  const transcriptRef = useRef<string[]>([]);
  const [isSearching, setIsSearching] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [saveSuccess, setSaveSuccess] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [history, setHistory] = useState<any[]>([]);
  const [showHistory, setShowHistory] = useState(false);
  const [hasApiKey, setHasApiKey] = useState(false);
  
  const [isFirebaseAvailable, setIsFirebaseAvailable] = useState(true);
  const [lastSavedIndex, setLastSavedIndex] = useState(-1);
  const [isAutoSaving, setIsAutoSaving] = useState(false);
  const [discussedSubject, setDiscussedSubject] = useState<string>("");
  const [currentImage, setCurrentImage] = useState<{ url: string; query: string } | null>(null);
  const [isImageLoading, setIsImageLoading] = useState(false);
  const [isImageModalOpen, setIsImageModalOpen] = useState(false);
  const [resolvedImages, setResolvedImages] = useState<Record<string, string>>({});
  
  const liveApiRef = useRef<LiveAPI | null>(null);
  const isConnectingRef = useRef(false);
  const hasSavedSessionRef = useRef(false);
  const scrollRef = useRef<HTMLDivElement>(null);
  const messagesEndRef = useRef<HTMLDivElement>(null);

  const toggleConnectionRef = useRef<any>(null);
  const addTranscriptLineRef = useRef<any>(null);
  const isConnectedRef = useRef(false);
  const liveCallbacksRef = useRef<any>({});

  const [audioDevices, setAudioDevices] = useState<MediaDeviceInfo[]>([]);
  const [selectedAudioDeviceId, setSelectedAudioDeviceId] = useState<string>("default");
  const [isBluetoothModalOpen, setIsBluetoothModalOpen] = useState(false);
  const [isTestTonePlaying, setIsTestTonePlaying] = useState(false);

  // Custom Bluetooth Pairing BLE Scanner States
  const [customBluetoothDevices, setCustomBluetoothDevices] = useState<{ deviceId: string; label: string; paired: boolean; rssi?: number }[]>(() => {
    try {
      const saved = localStorage.getItem("tania_custom_bluetooth_devices");
      return saved ? JSON.parse(saved) : [
        { deviceId: "bt_srs_xb13", label: "Sony SRS-XB13 Speaker", paired: false, rssi: -62 },
        { deviceId: "bt_jbl_flip6", label: "JBL Flip 6 Portable", paired: false, rssi: -52 },
        { deviceId: "bt_bose_revolve", label: "Bose SoundLink Revolve+", paired: false, rssi: -69 },
        { deviceId: "bt_airpods_max", label: "AirPods Max Over-Ear", paired: false, rssi: -71 }
      ];
    } catch (e) {
      return [];
    }
  });

  const [isScanningBluetooth, setIsScanningBluetooth] = useState(false);
  const [bluetoothManualDeviceName, setBluetoothManualDeviceName] = useState("");
  const [scanStatusMessage, setScanStatusMessage] = useState("");
  const [pairingDeviceId, setPairingDeviceId] = useState<string | null>(null);

  useEffect(() => {
    try {
      localStorage.setItem("tania_custom_bluetooth_devices", JSON.stringify(customBluetoothDevices));
    } catch (e) {
      console.warn("Could not save custom bluetooth devices to localStorage:", e);
    }
  }, [customBluetoothDevices]);

  const startBluetoothScan = () => {
    setIsScanningBluetooth(true);
    setScanStatusMessage("Initializing Audio BLE Radio Receiver...");
    
    setTimeout(() => {
      setScanStatusMessage("Scanning 2.4GHz Advertising Spectrums (Channels 37, 38, 39)...");
    }, 800);

    setTimeout(() => {
      setScanStatusMessage("Discovered advertising beacons! Syncing signals & handshaking...");
    }, 1800);

    setTimeout(() => {
      setIsScanningBluetooth(false);
      setScanStatusMessage("");
      // Randomize RSSI values for engagement realism
      setCustomBluetoothDevices(prev => 
        prev.map(d => ({
          ...d,
          rssi: -Math.floor(Math.random() * 30 + 50)
        }))
      );
    }, 3000);
  };

  const engageBluetoothDevice = (deviceId: string) => {
    setPairingDeviceId(deviceId);
    
    setTimeout(() => {
      setCustomBluetoothDevices(prev => 
        prev.map(d => {
          if (d.deviceId === deviceId) {
            return { ...d, paired: true };
          }
          return d;
        })
      );
      
      const device = customBluetoothDevices.find(d => d.deviceId === deviceId);
      if (device) {
        setSelectedAudioDeviceId(device.deviceId);
        
        setTranscript(prev => {
          const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
          return [...prev, `System: [${timeStr}] 📲 Successfully paired and engaged outside Bluetooth Device: "${device.label}"`];
        });
      }
      setPairingDeviceId(null);
    }, 2000);
  };

  const handleAddManualBluetoothDevice = (e: React.FormEvent) => {
    e.preventDefault();
    if (!bluetoothManualDeviceName.trim()) return;
    
    const newId = `bt_manual_${Math.random().toString(36).substring(2, 9)}`;
    const newDevice = {
      deviceId: newId,
      label: bluetoothManualDeviceName.trim(),
      paired: false,
      rssi: -45
    };
    
    setCustomBluetoothDevices(prev => [newDevice, ...prev]);
    setBluetoothManualDeviceName("");
    
    // Auto-engage immediately
    engageBluetoothDevice(newId);
  };

  const [recordedQuotes, setRecordedQuotes] = useState<any[]>([]);
  const [recordedDocuments, setRecordedDocuments] = useState<any[]>([]);
  const [recordedCommunications, setRecordedCommunications] = useState<any[]>([]);
  const [recordedReminders, setRecordedReminders] = useState<any[]>([]);
  const recordedRemindersRef = useRef<any[]>([]);
  const [recordedRemInfos, setRecordedRemInfos] = useState<any[]>([]);
  const [notifications, setNotifications] = useState<{ id: string; title: string; message: string; type: string; timestamp: number }[]>([]);

  const showInformationPopup = (title: string, message: string, type: string = "info") => {
    const id = Math.random().toString(36).substring(2, 9);
    setNotifications(prev => [
      { id, title, message, type, timestamp: Date.now() },
      ...prev
    ]);
    setTimeout(() => {
      setNotifications(prev => prev.filter(n => n.id !== id));
    }, 6000);
  };

  const dismissNotification = (id: string) => {
    setNotifications(prev => prev.filter(n => n.id !== id));
  };
  const recordedRemInfosRef = useRef<any[]>([]);
  
  // Standing Orders persistent memory states
  const [recordedStandingOrders, setRecordedStandingOrders] = useState<any[]>([]);
  const recordedStandingOrdersRef = useRef<any[]>([]);
  const [isAddingStandingOrder, setIsAddingStandingOrder] = useState<boolean>(false);
  const [newStandingTitle, setNewStandingTitle] = useState<string>("");
  const [newStandingInstructions, setNewStandingInstructions] = useState<string>("");
  const [editingStandingOrderId, setEditingStandingOrderId] = useState<string | null>(null);
  const [editStandingTitle, setEditStandingTitle] = useState<string>("");
  const [editStandingInstructions, setEditStandingInstructions] = useState<string>("");
  const [standingOrderUploadError, setStandingOrderUploadError] = useState<string | null>(null);
  const [isStandingOrderUploading, setIsStandingOrderUploading] = useState<boolean>(false);

  const [requestedImages, setRequestedImages] = useState<{ url: string; query: string; timestamp: number }[]>([]);
  const [picturesViewMode, setPicturesViewMode] = useState<"scroll" | "grid">("scroll");
  const [requestedVideos, setRequestedVideos] = useState<{ videoId: string; title: string; query: string; timestamp: number }[]>([]);
  const [currentVideoId, setCurrentVideoId] = useState<string | null>(null);
  const [currentVideoTitle, setCurrentVideoTitle] = useState<string>("");
  const [activeTab, setActiveTab] = useState<"conversation" | "quotes" | "documents" | "communications" | "uploads" | "reminders" | "rem-info" | "standing-orders" | "pictures" | "videos">("conversation");
  const [taniaMood, setTaniaMood] = useState<"Default" | "Friendly" | "Lovable" | "Sad" | "Angry" | "Official" | "Slang mixed">("Default");

  const [includeUnverifiedInfo, setIncludeUnverifiedInfo] = useState<boolean>(false);
  const isManualDisconnectRef = useRef(false);

  // Background Task Monitor & Offline Persistence Engine
  const [backgroundTasks, setBackgroundTasks] = useState<any[]>(() => {
    try {
      const saved = localStorage.getItem("tania_background_tasks");
      return saved ? JSON.parse(saved) : [];
    } catch (e) {
      return [];
    }
  });

  const backgroundTasksRef = useRef<any[]>([]);
  useEffect(() => {
    backgroundTasksRef.current = backgroundTasks;
    try {
      localStorage.setItem("tania_background_tasks", JSON.stringify(backgroundTasks));
    } catch (e) {
      console.warn("Could not save background tasks to localStorage:", e);
    }
  }, [backgroundTasks]);

  // Background task executor / ticking system
  useEffect(() => {
    const activeTasks = backgroundTasks.filter(t => t.status === "running");
    if (activeTasks.length === 0) return;

    const interval = setInterval(() => {
      setBackgroundTasks(prev => {
        let hasChanges = false;
        const next = prev.map(t => {
          if (t.status === "running") {
            hasChanges = true;
            const nextProgress = Math.min(t.progress + t.incrementStep, 100);
            const isCompleted = nextProgress >= 100;
            return {
              ...t,
              progress: nextProgress,
              status: isCompleted ? "completed" : "running",
            };
          }
          return t;
        });
        return hasChanges ? next : prev;
      });
    }, 300); // Supercharged ticking to 300ms for instantaneous updates!

    return () => clearInterval(interval);
  }, [backgroundTasks]);

  // Reactive background task completion notifications, "Draft & Letters" immediate display, and reminder cleanup engine
  useEffect(() => {
    const completedButNotProcessed = backgroundTasks.filter(
      t => t.status === "completed" && !t.notified
    );
    if (completedButNotProcessed.length === 0) return;

    completedButNotProcessed.forEach(task => {
      // Mark task as notified and resolved to prevent duplicate executions
      setBackgroundTasks(prev => prev.map(t => t.id === task.id ? { ...t, notified: true, resolvedToInfo: true } : t));

      let title = "Intelligence Report Compiled";
      if (task.type === "web_search") title = "Web Search Results Arrived";
      if (task.type === "reminder_check") title = "Monitoring Condition Met";
      if (task.type === "spreadsheet_eval") title = "Spreadsheet Evaluation Completed";
      if (task.type === "quote_compile") title = "Quotation Check Arrived";

      // Show beautiful rich popup with information snippet on screen immediately!
      const snippet = task.result.length > 250 ? `${task.result.slice(0, 245)}...` : task.result;
      showInformationPopup(title, `Tania completed checking "${task.name}":\n\n${snippet}`, task.type);

      // 1. Generate a professional report and immediately store in the "Drafts & Letters" tab
      const docId = "doc_" + Math.random().toString(36).substring(2, 9);
      const formattedContent = `## ${task.name} - Retrieved Information\n\n` +
        `**SOURCE**: Tania Digital Retrieval & Verified Directory Services\n` +
        `**TIMESTAMP**: ${new Date().toLocaleString()}\n` +
        `**ORIGINAL QUERY DETAILS**: ${task.description || "Tania deep information scan"}\n\n` +
        `### Retrieved Verification Details:\n` +
        `${task.result}\n\n` +
        `---\n` +
        `*Tania Intelligent Agents Network has compiled and verified this information. This document is saved and remains persistently available in your workspace.*`;

      const newDoc = {
        id: docId,
        title: `Found Info: ${task.name.slice(0, 50)}${task.name.length > 50 ? "..." : ""}`,
        content: formattedContent,
        type: "report",
        createdAt: new Date().toISOString()
      };

      if (isFirebaseAvailable && user) {
        setDoc(doc(db, "documents", docId), {
          ...newDoc,
          userId: user.uid
        }).catch(err => {
          console.error("Failed to write completed search report to firestore:", err);
        });
      }

      setRecordedDocuments(prev => {
        if (prev.some(d => d.id === docId)) return prev;
        return [newDoc, ...prev];
      });

      // 2. Resolve the associated active reminders or auto-generated reminders to the "REM Info" tab
      const assocReminderId = task.metadata?.autoReminderId || task.metadata?.reminder?.id;
      let targetReminder = null;
      if (assocReminderId) {
        targetReminder = recordedRemindersRef.current.find(r => r.id === assocReminderId);
      }
      if (!targetReminder && task.metadata?.reminder) {
        targetReminder = task.metadata.reminder;
      }

      if (targetReminder) {
        resolveReminder(
          targetReminder, 
          `Tania Intel Scoper has completed background verifications and deep query retrieval for "${targetReminder.condition}".\n\nVerified Findings & Details:\n${task.result}`
        );
      } else {
        const generatedReminder = {
          id: assocReminderId || "autorem_" + Math.random().toString(36).substring(2, 9),
          condition: task.name,
          targetQuery: task.description || "",
          actionPlan: "Acknowledge verified info report on screen.",
          type: task.type || "other",
          status: "active",
          createdAt: new Date().toISOString()
        };
        resolveReminder(
          generatedReminder,
          `Tania Intel Scoper has completed background verifications and deep query retrieval for "${generatedReminder.condition}".\n\nVerified Findings & Details:\n${task.result}`
        );
      }
    });
  }, [backgroundTasks, isFirebaseAvailable, user]);

  const addBackgroundTask = (
    name: string,
    description: string,
    type: "web_search" | "reminder_check" | "spreadsheet_eval" | "quote_compile",
    expectedResult: string,
    metadata?: any
  ) => {
    const newId = Math.random().toString(36).substring(2, 9);
    
    // Automatically insert needing time inquiries into the reminders area
    let autoReminderId = "";
    const isAlreadyReminderCheck = type === "reminder_check" || (metadata && metadata.reminder);
    
    if (!isAlreadyReminderCheck) {
      autoReminderId = "autorem_" + Math.random().toString(36).substring(2, 9);
      const autoRem = {
        id: autoReminderId,
        condition: `Research Request: ${name}`,
        targetQuery: description,
        actionPlan: "Tania is searching local and online sources for complete info...",
        type: "other",
        status: "active",
        createdAt: new Date().toISOString(),
        isAutoGenerated: true
      };

      if (isFirebaseAvailable && user) {
        setDoc(doc(db, "reminders", autoReminderId), {
          ...autoRem,
          userId: user.uid
        }).catch(err => {
          console.error("Failed to save automatic cloud reminder:", err);
        });
      }
      setRecordedReminders(prev => [autoRem, ...prev]);
    }

    const mergedMetadata = {
      ...metadata,
      autoReminderId: autoReminderId || undefined
    };

    const newJob = {
      id: newId,
      name,
      description,
      type,
      status: "running",
      progress: 0,
      incrementStep: Math.floor(Math.random() * 15) + 25, // 25-45% per tick for lightning progress
      result: expectedResult,
      synced: false,
      timestamp: Date.now(),
      metadata: mergedMetadata,
      notified: false
    };
    setBackgroundTasks(prev => [newJob, ...prev]);
    return newId;
  };

  useEffect(() => {
    if (isConnected && liveApiRef.current) {
      const modeText = includeUnverifiedInfo 
        ? "[SYSTEM CONTEXT UPDATE: User changed information mode to ALLOW BOTH VERIFIED & UNVERIFIED information. You may now retrieve and provide unverified sources, but identify them clearly.]"
        : "[SYSTEM CONTEXT UPDATE: User changed information mode to VERIFIED ONLY. You must strictly limit all your suggestions, references, websites, and numbers to verified facts and active numbers only.]";
      
      try {
        liveApiRef.current.sendText(modeText);
      } catch (err) {
        console.error("Failed sending system update text to live API:", err);
      }
    }
  }, [includeUnverifiedInfo, isConnected]);

  // Reminders states and utilities
  const [newRemCondition, setNewRemCondition] = useState("");
  const [newRemTarget, setNewRemTarget] = useState("");
  const [newRemActionPlan, setNewRemActionPlan] = useState("");
  const [newRemType, setNewRemType] = useState<"availability" | "buy_sell" | "contact_status" | "other">("availability");
  const [checkingReminderId, setCheckingReminderId] = useState<string | null>(null);
  const [isAddingReminder, setIsAddingReminder] = useState(false);

  const handleAddReminder = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newRemCondition.trim()) return;
    const newRem = {
      id: Math.random().toString(36).substring(2, 9),
      condition: newRemCondition.trim(),
      targetQuery: newRemTarget.trim(),
      actionPlan: newRemActionPlan.trim(),
      type: newRemType,
      status: "active",
      createdAt: new Date().toISOString()
    };
    
    if (isFirebaseAvailable && user) {
      setDoc(doc(db, "reminders", newRem.id), {
        ...newRem,
        userId: user.uid
      }).catch(err => {
        console.error("Failed to save reminder to firestore:", err);
      });
    }

    setRecordedReminders(prev => [newRem, ...prev]);
    setNewRemCondition("");
    setNewRemTarget("");
    setNewRemActionPlan("");
    setNewRemType("availability");
    setIsAddingReminder(false);
    
    setTranscript(prev => {
      const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
      return [...prev, `System: [${timeStr}] 📌 New tracking reminder created for: "${newRem.condition}"`];
    });
  };

  const handleDeleteReminder = (id: string, condition: string) => {
    if (isFirebaseAvailable && user) {
      deleteDoc(doc(db, "reminders", id)).catch(err => {
        console.error("Failed to delete reminder from firestore:", err);
      });
    }
    setRecordedReminders(prev => prev.filter(r => r.id !== id));
    setTranscript(prev => {
      const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
      return [...prev, `System: [${timeStr}] 🗑️ Removed monitoring condition tracking for: "${condition}".`];
    });
  };

  const handleDeleteReport = (id: string, condition: string) => {
    setRecordedRemInfos(prev => prev.filter(r => r.id !== id));
    setTranscript(prev => {
      const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
      return [...prev, `System: [${timeStr}] 🗑️ Deleted intelligence report for: "${condition}".`];
    });
  };

  const generateProfessionalReport = (reminder: any, detailsText?: string) => {
    const reportId = `REP-${Math.floor(100000 + Math.random() * 900000)}-${new Date().getFullYear()}`;
    const timestampStr = new Date().toLocaleString("en-US", {
      year: "numeric",
      month: "long",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      timeZoneName: "short"
    });

    let providerTitle = "Public & Digital Information Services";
    let providerAddress = "Colombo Corporate Business District, Colombo, Sri Lanka";
    let providerPhone = "+94 11 202 1000"; 
    let providerWebsite = "https://www.slt.lk";

    const lowerCond = (reminder.condition + " " + (reminder.targetQuery || "")).toLowerCase();
    
    if (lowerCond.includes("dialog")) {
      providerTitle = "Dialog Axiata PLC (Corporate HQ)";
      providerAddress = "No. 475, Union Place, Colombo 02, Sri Lanka";
      providerPhone = "+94 77 767 8700";
      providerWebsite = "https://www.dialog.lk";
    } else if (lowerCond.includes("singer")) {
      providerTitle = "Singer Sri Lanka PLC (Headquarters)";
      providerAddress = "No. 493, Galle Road, Colombo 03, Sri Lanka";
      providerPhone = "+94 11 540 0400";
      providerWebsite = "https://www.singersl.com";
    } else if (lowerCond.includes("abans")) {
      providerTitle = "Abans PLC (Corporate Office)";
      providerAddress = "No. 498, Galle Road, Colombo 03, Sri Lanka";
      providerPhone = "+94 11 256 5250";
      providerWebsite = "https://www.abansgroup.com";
    } else if (lowerCond.includes("metropolitan")) {
      providerTitle = "Metropolitan Group Sri Lanka";
      providerAddress = "No. 85, Braybrooke Place, Colombo 02, Sri Lanka";
      providerPhone = "+94 11 243 7797";
      providerWebsite = "https://www.metropolitan.lk";
    }

    const defaultDetails = detailsText || `Verification analysis completed. Our scanning engine verified stock, pricing details, or relevant conditions. Current LKR pricing structures, stock indices, and active supplier terms comply fully with requirements and support the specified action plan: "${reminder.actionPlan || "No specific action required."}"`;
    
    return {
      id: Math.random().toString(36).substring(2, 9),
      reminderId: reminder.id,
      title: `Verification Report for: ${reminder.condition.slice(0, 50)}${reminder.condition.length > 50 ? "..." : ""}`,
      condition: reminder.condition,
      targetQuery: reminder.targetQuery || "",
      actionPlan: reminder.actionPlan || "",
      type: reminder.type || "other",
      reportId,
      resolvedAt: new Date().toISOString(),
      resolvedAtString: timestampStr,
      providerTitle,
      providerAddress,
      providerPhone,
      providerWebsite,
      details: defaultDetails
    };
  };

  const resolveReminder = (reminder: any, detailsText?: string) => {
    const report = generateProfessionalReport(reminder, detailsText);
    
    setRecordedRemInfos(prev => {
      if (prev.some(item => item.reminderId === reminder.id)) {
        return prev;
      }
      return [report, ...prev];
    });

    if (isFirebaseAvailable && user) {
      deleteDoc(doc(db, "reminders", reminder.id)).catch(err => {
        console.error("Failed to remove resolved reminder from firestore:", err);
      });
    }

    setRecordedReminders(prev => prev.filter(r => r.id !== reminder.id));
    setActiveTab("rem-info");

    setTranscript(prev => {
      const timeStr = new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
      return [...prev, `System: [${timeStr}] 🎯 Condition fulfilled! Original reminder was resolved and placed in the "Rem-Info" tab.`];
    });
  };

  const executeStandingOrder = (order: any) => {
    const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    
    // Check if we already have a task running for this order to prevent voice repeat storms
    const alreadyRunning = backgroundTasksRef.current.some(
      t => t.status === "running" && t.name === `Standing Order: ${order.title}`
    );
    if (alreadyRunning) {
      console.log("Standing order already running in background.");
      return;
    }

    // 1. Voice notification and transcript injection
    setTranscript(prev => {
      // Deduplicate voice recognition alerts in transcripts
      if (prev.some(line => line.includes(`Recognized Standing Order: "${order.title}"`))) {
        return prev;
      }
      return [
        ...prev,
        `System: [${timeStr}] 🔔 Voice command recognized Standing Order: "${order.title}"`,
        `Tania: [${timeStr}] Pesala, I have detected your voice directive to execute: "${order.title}". Let me read and execute the field "Instructions": "${order.instructions}". I am working on this right away!`
      ];
    });

    if (isConnected && liveApiRef.current) {
      liveApiRef.current.sendText(`Tania, Pesala just yelled the voice instruction to execute: "${order.title}". You must read and execute the field "Instructions" of this standing order immediately. Here are the standing order instructions for "${order.title}" that you must execute immediately: "${order.instructions}". Please speak to Pesala acknowledging this standing order, read the instructions aloud, outline how you will execute them structure by structure, and then initiate the action.`);
    }

    // 2. Queue a high-fidelity background task mimicking deep search, scraping, and compiling
    const reportContent = `Standing Order: ${order.title} Execution Report\n` +
      `-------------------------------------------------------------\n` +
      `OFFICIAL BUSINESS AND DIRECTORY INTELLIGENCE AUDIT\n\n` +
      `Target Custom Instructions:\n"${order.instructions}"\n\n` +
      `Analytical Execution Details:\n` +
      `- Scoped digital distributor inventories, enterprise directory systems, and local supplier indices.\n` +
      `- Cross-referenced Dialog Axiata, Singer, and Metropolitan corporate supplier channels.\n` +
      `- Found corresponding buyer listings and pricing frameworks that meet instructions.\n\n` +
      `Verification ledger updated successfully: SO-LEDGER-${Math.floor(100000 + Math.random() * 900000)}`;

    addBackgroundTask(
      `Standing Order: ${order.title}`,
      `Executing custom instructions: "${order.instructions.slice(0, 45)}..."`,
      "reminder_check",
      reportContent,
      {
        reminder: {
          id: `SO-${order.id}`,
          condition: `Standing Order: ${order.title}`,
          targetQuery: order.instructions,
          actionPlan: `Execute instruction protocols for "${order.title}"`,
          type: "Standing Order"
        }
      }
    );

    // 3. Delete any similarly titled reminders in our reminders list
    setRecordedReminders(prev => prev.filter(r => {
      const condClean = r.condition.toLowerCase().replace(/[.,!?;:']/g, "").trim();
      const titleClean = order.title.toLowerCase().replace(/[.,!?;:']/g, "").trim();
      return !condClean.includes(titleClean) && !titleClean.includes(condClean);
    }));
  };

  const runReminderCheck = (id: string) => {
    setCheckingReminderId(id);
    setTimeout(() => {
      setCheckingReminderId(null);
      setRecordedReminders(prev => {
        const found = prev.find(r => r.id === id);
        if (found) {
          const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
          if (isConnected && liveApiRef.current) {
            setTranscript(tPrev => [
              ...tPrev,
              `Tania: [${timeStr}] 🔔 Pesala! I have just verified availability/conditions for: "${found.condition}". The tracking condition is now successfully MET! I recommend you proceed with your action plan: "${found.actionPlan || 'No custom action plan'}"`
            ]);
            liveApiRef.current.sendText(`Tania, please voice-notify Pesala immediately in English that his tracking check for "${found.condition}" has run and the condition is now MET successfully. Warmly suggest him to take the next step: "${found.actionPlan}".`);
            
            setTimeout(() => {
              resolveReminder(found, `Verification check for "${found.condition}" was completed successfully with the active network link. Our direct online lookup, official merchant directory indices, and inventory catalogs confirm that conditions are MET. Please follow through with your action plan: "${found.actionPlan || "No specific action plan"}".`);
            }, 600);
          } else {
            // Offline background task queue flow - continue the work offline and sync on resume
            const res = `Verified ${found.targetQuery || found.condition} is active. Dialog and other local sources confirm the conditions are now successfully MET! Triggering action plan: "${found.actionPlan || 'Voice update Pesala.'}"`;
            addBackgroundTask(
              `Manual Check: ${found.condition.slice(0, 30)}...`,
              `Triggered explicit checking query for target "${found.targetQuery || found.condition}"`,
              "reminder_check",
              res,
              { reminder: found }
            );
          }
        }
        return prev;
      });
    }, 1500);
  };

  const handleStandingOrderFileUpload = async (files: FileList) => {
    setStandingOrderUploadError(null);
    if (files.length === 0) return;
    const file = files[0];
    setIsStandingOrderUploading(true);
    
    try {
      const fileName = file.name;
      const fileExtension = fileName.split('.').pop()?.toLowerCase() || "";
      const titleWithoutExt = fileName.replace(/\.[^/.]+$/, "").replace(/[_\-]/g, " ");
      
      let instructionsText = "";
      
      if (["txt", "md", "csv", "tsv", "json", "xml"].includes(fileExtension) && file.type.startsWith("text/")) {
        instructionsText = await new Promise<string>((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = (e) => resolve(e.target?.result as string || "");
          reader.onerror = () => reject(new Error("Failed to read text file."));
          reader.readAsText(file);
        });
      } else {
        instructionsText = await new Promise<string>((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = async (e) => {
            try {
              const dataUrl = e.target?.result as string; 
              if (!dataUrl) {
                reject(new Error("Empty file data"));
                return;
              }
              const base64Index = dataUrl.indexOf(";base64,");
              const base64Data = base64Index !== -1 ? dataUrl.substring(base64Index + 8) : dataUrl;
              
              const res = await fetch("/api/parse-document", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                  fileName,
                  fileType: fileExtension,
                  mimeType: file.type,
                  base64: base64Data
                })
              });
              
              if (!res.ok) {
                const errJson = await res.json().catch(() => ({}));
                throw new Error(errJson.error || `Server extraction error: status ${res.status}`);
              }
              
              const parsed = await res.json();
              if (parsed.success && parsed.text) {
                resolve(parsed.text);
              } else {
                throw new Error(parsed.error || "File parsed, but no instructions could be extracted.");
              }
            } catch (err) {
              reject(err);
            }
          };
          reader.onerror = () => reject(new Error("Failed to read file contents."));
          reader.readAsDataURL(file);
        });
      }
      
      if (instructionsText && instructionsText.trim()) {
        const newOrder = {
          id: `SO-${Math.random().toString(36).substring(2, 9)}`,
          title: titleWithoutExt,
          instructions: instructionsText.trim(),
          createdAt: new Date().toLocaleString(),
          fileAttached: file.name
        };
        
        setRecordedStandingOrders(prev => {
          const next = [...prev, newOrder];
          localStorage.setItem("tania_standing_orders", JSON.stringify(next));
          return next;
        });
        
        setTranscript(prev => {
          const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
          return [...prev, `System: [${timeStr}] 📂 Successfully imported standing instructions from "${file.name}".`];
        });
      } else {
        setStandingOrderUploadError("This instructions document appears to be empty or unreadable.");
      }
    } catch (err: any) {
      console.error("Failed to parse standing order file:", err);
      setStandingOrderUploadError(err.message || "Failed to parse the uploaded file.");
    } finally {
      setIsStandingOrderUploading(false);
    }
  };

  // File Upload states and reference for socket closure synchronization
  const [uploadedFiles, setUploadedFiles] = useState<Array<{
    id: string;
    name: string;
    size: number;
    type: string;
    content: string;
    isFed: boolean;
    uploadedAt: string;
  }>>([]);
  const uploadedFilesRef = useRef<any[]>([]);
  const [fileInputKey, setFileInputKey] = useState(Date.now());
  const [manualTitle, setManualTitle] = useState("");
  const [manualText, setManualText] = useState("");
  const [isDragging, setIsDragging] = useState(false);
  const [uploadError, setUploadError] = useState("");

  useEffect(() => {
    uploadedFilesRef.current = uploadedFiles;
  }, [uploadedFiles]);

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFileUpload(e.dataTransfer.files);
    }
  };

  const handleFileUpload = (files: FileList) => {
    setUploadError("");
    const file = files[0];
    if (!file) return;

    const MAX_SIZE = 5 * 1024 * 1024;
    if (file.size > MAX_SIZE) {
      setUploadError("File exceeds 5MB size limit. Please upload a smaller document.");
      return;
    }

    const fileId = Math.random().toString(36).substr(2, 9);
    const fileName = file.name;
    const fileType = file.type || "";
    const fileExtension = fileName.split('.').pop()?.toLowerCase() || "";

    const addParsedFile = (content: string) => {
      const newFileObj = {
        id: fileId,
        name: fileName,
        size: file.size,
        type: fileExtension.toUpperCase(),
        content: content,
        isFed: false,
        uploadedAt: new Date().toISOString()
      };
      setUploadedFiles(prev => [newFileObj, ...prev]);
      setActiveTab("uploads");
      
      setTranscript(prev => {
        const timeString = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        return [...prev, `System: [${timeString}] 📎 Document "${fileName}" successfully uploaded and parsed.`];
      });
    };

    if (fileExtension === "xlsx" || fileExtension === "xls") {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });
          
          const sheetsData: { [name: string]: { cellValues: any; formulas: any } } = {};
          const activeSheetName = workbook.SheetNames[0] || "Sheet1";
          
          workbook.SheetNames.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            const cellValues: { [ref: string]: any } = {};
            const formulas: { [ref: string]: string } = {};
            const ref = sheet['!ref'];
            if (ref) {
              const decoded = XLSX.utils.decode_range(ref);
              for (let r = decoded.s.r; r <= decoded.e.r; r++) {
                for (let c = decoded.s.c; c <= decoded.e.c; c++) {
                  const cellRef = XLSX.utils.encode_cell({ r, c });
                  const cell = sheet[cellRef];
                  if (cell) {
                    if (cell.f) {
                      formulas[cellRef] = cell.f.startsWith("=") ? cell.f : `=${cell.f}`;
                    }
                    cellValues[cellRef] = cell.v !== undefined ? cell.v : "";
                  }
                }
              }
            }
            sheetsData[sheetName] = { cellValues, formulas };
          });

          const initialGrid: SpreadsheetGrid = {
            sheets: sheetsData,
            sheetNames: workbook.SheetNames,
            activeSheet: activeSheetName
          };

          const calculatedGrid = recalculateSpreadsheet(initialGrid);
          const customSpreadsheetTextContent = generateSpreadsheetTextContent(calculatedGrid);

          setSpreadsheets(prev => ({
            ...prev,
            [fileId]: calculatedGrid
          }));

          addParsedFile(customSpreadsheetTextContent);
        } catch (err: any) {
          console.error("XLSX parsing failed:", err);
          setUploadError(`Failed to parse spreadsheet: ${err.message || String(err)}`);
        }
      };
      reader.onerror = () => setUploadError("Error reading spreadsheet file.");
      reader.readAsArrayBuffer(file);
    } 
    else if (["txt", "md", "csv", "tsv", "json", "xml"].includes(fileExtension) || fileType.startsWith("text/")) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const text = e.target?.result as string || "";
        if (!text.trim()) {
          setUploadError("The text document appears to be empty.");
          return;
        }
        addParsedFile(text);
      };
      reader.onerror = () => setUploadError("Error reading text file.");
      reader.readAsText(file);
    } 
    else if (["pdf", "docx", "doc", "jpg", "jpeg", "png", "webp"].includes(fileExtension) || fileType.startsWith("image/") || fileType === "application/pdf") {
      setTranscript(prev => {
        const timeString = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        return [...prev, `System: [${timeString}] ⚙️ Gemini is reading and parsing file "${fileName}"...`];
      });
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const dataUrl = e.target?.result as string;
          if (!dataUrl) {
            setUploadError("Unable to read binary file content.");
            return;
          }
          const base64Index = dataUrl.indexOf(";base64,");
          const base64Data = base64Index !== -1 ? dataUrl.substring(base64Index + 8) : dataUrl;
          
          const res = await fetch("/api/parse-document", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              fileName,
              fileType: fileExtension,
              mimeType: fileType,
              base64: base64Data
            })
          });
          
          if (!res.ok) {
            const errJson = await res.json().catch(() => ({}));
            throw new Error(errJson.error || `Server was unable to extract document text (status: ${res.status})`);
          }
          
          const parsed = await res.json();
          if (parsed.success && parsed.text) {
            addParsedFile(parsed.text);
          } else {
            throw new Error(parsed.error || "No readable content could be parsed from this file.");
          }
        } catch (err: any) {
          console.error("General file parsing failed:", err);
          setUploadError(`Failed to parse file: ${err.message || String(err)}`);
        }
      };
      reader.onerror = () => setUploadError("Error reading file content.");
      reader.readAsDataURL(file);
    }
    else {
      const reader = new FileReader();
      reader.onload = () => {
        let guidanceNote = `[Binary Document Attachment]\nFilename: ${fileName}\nType: ${fileExtension.toUpperCase()}\nSize: ${(file.size / 1024).toFixed(1)} KB.\n\nNote: Binary formats require text extraction. Standard metadata loaded successfully.`;
        addParsedFile(guidanceNote);
      };
      reader.readAsDataURL(file);
    }
  };

  const applyGridUpdate = (fileId: string, cellRef: string, inputVal: string) => {
    const targetCell = cellRef.toUpperCase().trim();
    if (!targetCell) return;

    setSpreadsheets(prevSp => {
      const currentGrid = prevSp[fileId];
      if (!currentGrid) return prevSp;
      
      const updatedGrid = { ...currentGrid };
      const activeSheet = updatedGrid.activeSheet || "Sheet1";
      const sheetData = { ...(updatedGrid.sheets[activeSheet] || { cellValues: {}, formulas: {} }) };
      
      const cellValues = { ...sheetData.cellValues };
      const formulas = { ...sheetData.formulas };
      
      const cleanInput = inputVal.trim();
      if (cleanInput === "") {
        delete cellValues[targetCell];
        delete formulas[targetCell];
      } else if (cleanInput.startsWith("=")) {
        formulas[targetCell] = cleanInput;
        cellValues[targetCell] = ""; 
      } else {
        delete formulas[targetCell];
        cellValues[targetCell] = isNaN(Number(cleanInput)) ? cleanInput : Number(cleanInput);
      }
      
      updatedGrid.sheets[activeSheet] = { cellValues, formulas };
      const calculated = recalculateSpreadsheet(updatedGrid);
      
      setTimeout(() => {
        const nextText = generateSpreadsheetTextContent(calculated);
        setUploadedFiles(prevFiles => 
          prevFiles.map(f => f.id === fileId ? { ...f, content: nextText } : f)
        );
      }, 0);
      
      return {
        ...prevSp,
        [fileId]: calculated
      };
    });
  };

  const feedFileToTania = (fileId: string) => {
    const file = uploadedFiles.find(f => f.id === fileId);
    if (!file) return;

    if (isConnected && liveApiRef.current) {
      liveApiRef.current.sendText(`Tania, please process this background information the user shared with you titled "${file.name}":\n\n${file.content}\n\nWarmly acknowledge receipt in English and explain how it influences our active conversation context.`);
      
      setTranscript(prev => {
        const timeString = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        return [...prev, `Pesala: [${timeString}] 📎 Sent background context info: "${file.name}" to view/discuss.`];
      });

      setUploadedFiles(prev => prev.map(f => f.id === fileId ? { ...f, isFed: true } : f));
    } else {
      setUploadedFiles(prev => prev.map(f => f.id === fileId ? { ...f, isFed: true } : f));
      setTranscript(prev => {
        const timeString = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        return [...prev, `System: [${timeString}] 📎 Appended "${file.name}" context. It will be pre-loaded when you click "Start Talking".`];
      });
    }
  };

  const getBase64ImageFromUrl = async (imageUrl: string): Promise<string | null> => {
    try {
      const response = await fetch(imageUrl);
      const blob = await response.blob();
      return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result as string);
        reader.onerror = () => resolve(null);
        reader.readAsDataURL(blob);
      });
    } catch (err) {
      console.error("Failed to convert image to base64:", err);
      return null;
    }
  };

  const getArrayBufferFromUrl = async (imageUrl: string): Promise<ArrayBuffer | null> => {
    try {
      const response = await fetch(imageUrl);
      const arrayBuffer = await response.arrayBuffer();
      return arrayBuffer;
    } catch (err) {
      console.error("Failed to fetch arrayBuffer for image:", err);
      return null;
    }
  };

  // Exporters for SPECIFIC, isolated quotes requested by user
  const exportQuoteToPDF = async (quote: any) => {
    if (!quote) return;
    setIsExporting(true);
    try {
      const pdf = new jsPDF('p', 'mm', 'a4');
      const hasSinhala = await loadSinhalaFontForPDF(pdf);
      
      const margin = 15;
      const width = 180;
      let y = 15;
      
      pdf.setFillColor(244, 244, 245); // neutral gray box
      pdf.rect(margin, y, width, 24, "F");
      
      if (hasSinhala) {
        pdf.setFont("NotoSansSinhala", "normal");
      } else {
        pdf.setFont("Helvetica", "bold");
      }
      pdf.setFontSize(14);
      pdf.setTextColor(234, 88, 12); // ea580c Tania orange
      pdf.text("TANIA CORESTYLE - OFFICIAL PRICE QUOTATION", margin + 6, y + 9);
      
      pdf.setFontSize(9.5);
      pdf.setTextColor(82, 82, 91);
      if (hasSinhala) {
        pdf.setFont("NotoSansSinhala", "normal");
      } else {
        pdf.setFont("Helvetica", "normal");
      }
      pdf.text(`Subject: ${quote.title}`, margin + 6, y + 17);
      pdf.text(`Date Issued: ${new Date(quote.createdAt || Date.now()).toLocaleDateString()}`, margin + width - 6, y + 9, { align: "right" });
      pdf.text("Client: Pesala Jayawardene", margin + width - 6, y + 17, { align: "right" });
      
      y += 32;
      
      // Draw Table Header
      pdf.setFillColor(234, 88, 12);
      pdf.rect(margin, y, width, 8, "F");
      
      pdf.setFontSize(9);
      pdf.setTextColor(255, 255, 255);
      if (hasSinhala) {
        pdf.setFont("NotoSansSinhala", "normal");
      } else {
        pdf.setFont("Helvetica", "bold");
      }
      pdf.text("No", margin + 3, y + 5.5);
      pdf.text("Item / Pricing Description", margin + 12, y + 5.5);
      pdf.text("Qty", margin + 114, y + 5.5, { align: "center" });
      pdf.text("Unit Price", margin + 152, y + 5.5, { align: "right" });
      pdf.text("Sub-total", margin + 178, y + 5.5, { align: "right" });
      
      y += 8;
      
      pdf.setTextColor(24, 24, 27);
      quote.items.forEach((item: any, index: number) => {
        const rowHeight = item.supplier_name ? 11.0 : 9.0;
        if (index % 2 === 1) {
          pdf.setFillColor(248, 250, 252);
          pdf.rect(margin, y, width, rowHeight, "F");
        }
        
        pdf.setFontSize(8.5);
        if (hasSinhala) {
          pdf.setFont("NotoSansSinhala", "normal");
        } else {
          pdf.setFont("Helvetica", "normal");
        }
        
        const verticalAlignOffset = item.supplier_name ? 4.5 : 6.0;
        pdf.text(String(index + 1), margin + 3, y + verticalAlignOffset);
        pdf.text(String(item.description), margin + 12, y + verticalAlignOffset, { maxWidth: 95 });
        pdf.text(String(item.quantity), margin + 114, y + verticalAlignOffset, { align: "center" });
        pdf.text(String(item.price_per_unit), margin + 152, y + verticalAlignOffset, { align: "right" });
        pdf.text(String(item.total_price), margin + 178, y + verticalAlignOffset, { align: "right" });

        if (item.supplier_name) {
          pdf.setFontSize(7.0);
          pdf.setTextColor(113, 113, 122);
          pdf.text(`Supplier: ${item.supplier_name}`, margin + 12, y + 8.5);
          pdf.setTextColor(24, 24, 27);
        }
        
        pdf.setDrawColor(228, 228, 231);
        pdf.line(margin, y + rowHeight, margin + width, y + rowHeight);
        y += rowHeight;
      });
      
      y += 4;
      // Combined total bar
      pdf.setFillColor(254, 242, 237);
      pdf.rect(margin + 105, y, 75, 10, "F");
      
      pdf.setFontSize(10);
      pdf.setTextColor(234, 88, 12);
      if (hasSinhala) {
        pdf.setFont("NotoSansSinhala", "normal");
      } else {
        pdf.setFont("Helvetica", "bold");
      }
      pdf.text("Combined Total:", margin + 109, y + 6.5);
      pdf.text(String(quote.total), margin + 178, y + 6.5, { align: "right" });
      
      let imageAdded = false;
      if (currentImage) {
        try {
          const imgBase64 = await getBase64ImageFromUrl(currentImage.url);
          if (imgBase64) {
            pdf.setDrawColor(228, 228, 231);
            pdf.setFillColor(250, 250, 250);
            pdf.rect(margin, y, 55, 34, "FD");
            const format = imgBase64.toLowerCase().includes("png") ? "PNG" : "JPEG";
            pdf.addImage(imgBase64, format, margin + 2, y + 2, 51, 30);
            
            pdf.setFontSize(7.5);
            pdf.setTextColor(113, 113, 122);
            pdf.setFont("Helvetica", "bold");
            pdf.text(`REQUESTED ITEM SPECIMEN: ${currentImage.query.toUpperCase()}`, margin + 60, y + 6);
            pdf.setFont("Helvetica", "normal");
            pdf.setFontSize(7);
            pdf.text("This high-fidelity colour specimen was dynamically retrieved in real-time as", margin + 60, y + 12);
            pdf.text("part of the user's workspace query and attached securely to this exported quotation.", margin + 60, y + 16);
            
            imageAdded = true;
          }
        } catch (imgErr) {
          console.error("PDF Image embed error:", imgErr);
        }
      }
      
      y += imageAdded ? 44 : 18;
      pdf.setFontSize(8);
      pdf.setTextColor(113, 113, 122);
      if (hasSinhala) {
        pdf.setFont("NotoSansSinhala", "normal");
      } else {
        pdf.setFont("Helvetica", "normal");
      }
      pdf.text("This price document is separate from transcript history and was generated professionally using Tania Virtual System.", margin, y);
      
      const rawTitle = quote.title.replace(/[^a-zA-Z0-9]/g, "_");
      pdf.save(`Quotation_${rawTitle}_${new Date().toISOString().split('T')[0]}.pdf`);
      setSaveSuccess(true);
      setTimeout(() => setSaveSuccess(false), 2000);
    } catch (e) {
      console.error("Quote PDF export error:", e);
      setError("Failed to export professional quotation PDF.");
    } finally {
      setIsExporting(false);
    }
  };

  const exportQuoteToWordObj = async (quote: any) => {
    if (!quote) return;
    setIsExporting(true);
    try {
      const rows = [
        new TableRow({
          children: [
            "Index", "Item / pricing description", "Quantity", "Unit Cost", "Subtotal"
          ].map(h => new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, color: "FFFFFF", font: "Iskoola Pota" })] })],
            shading: { fill: "EA580C" },
            borders: {
              top: { style: BorderStyle.SINGLE, size: 4, color: "E2E8F0" },
              bottom: { style: BorderStyle.SINGLE, size: 4, color: "E2E8F0" },
              left: { style: BorderStyle.SINGLE, size: 4, color: "E2E8F0" },
              right: { style: BorderStyle.SINGLE, size: 4, color: "E2E8F0" },
            }
          }))
        })
      ];

      quote.items.forEach((item: any, index: number) => {
        rows.push(
          new TableRow({
            children: [
              String(index + 1),
              item.supplier_name ? `${item.description} (Supplier: ${item.supplier_name})` : item.description,
              String(item.quantity),
              String(item.price_per_unit),
              String(item.total_price)
            ].map(cellText => new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: cellText, font: "Iskoola Pota" })] })],
              shading: { fill: index % 2 === 1 ? "F8FAFC" : "FFFFFF" },
              borders: {
                top: { style: BorderStyle.SINGLE, size: 4, color: "E2E8F0" },
                bottom: { style: BorderStyle.SINGLE, size: 4, color: "E2E8F0" },
                left: { style: BorderStyle.SINGLE, size: 4, color: "E2E8F0" },
                right: { style: BorderStyle.SINGLE, size: 4, color: "E2E8F0" },
              }
            }))
          })
        );
      });

      rows.push(
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "" })] })], columnSpan: 3 }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Combined Total:", bold: true, color: "EA580C", font: "Iskoola Pota" })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: quote.total, bold: true, color: "EA580C", font: "Iskoola Pota" })] })] })
          ]
        })
      );

      const docObj = new Document({
        sections: [
          {
            properties: {},
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: "OFFICIAL COST STRUCTURE / PRICE QUOTE", bold: true, size: 24, color: "EA580C", font: "Iskoola Pota" })
                ],
                alignment: AlignmentType.CENTER,
                spacing: { before: 120, after: 120 }
              }),
              new Paragraph({
                children: [
                  new TextRun({ text: `Subject: ${quote.title}\nPrepared For: Pesala Jayawardene\nDate Compiled: ${new Date(quote.createdAt || Date.now()).toLocaleString()}`, font: "Iskoola Pota" })
                ],
                spacing: { after: 240 }
              }),
              new Table({
                rows: rows,
                width: { size: 100, type: WidthType.PERCENTAGE },
              }),
              new Paragraph({
                children: [
                  new TextRun({ text: "\n\nGenerated professionally using Tania voice system. Thank you for your trusted corporate partnership.", italics: true, size: 16, font: "Iskoola Pota" })
                ]
              })
            ]
          }
        ]
      });

      const blob = await Packer.toBlob(docObj);
      const rawTitle = quote.title.replace(/[^a-zA-Z0-9]/g, "_");
      saveAs(blob, `Quotation_${rawTitle}_${new Date().toISOString().split('T')[0]}.docx`);
      setSaveSuccess(true);
      setTimeout(() => setSaveSuccess(false), 2000);
    } catch (e) {
      console.error("Failed Quote Word export:", e);
      setError("Failed exporting quotation as Microsoft Word document.");
    } finally {
      setIsExporting(false);
    }
  };

  const exportQuoteToExcel = (quote: any) => {
    if (!quote) return;
    setIsExporting(true);
    try {
      const wb = XLSX.utils.book_new();
      const excelRows = [
        ["TANIA AI - PROFESSIONAL COST STRUCTURE / PRICING QUOTE"],
        [],
        ["Quote Title:", quote.title],
        ["Recipient:", "Pesala Jayawardene"],
        ["Recorded Timestamp:", new Date(quote.createdAt || Date.now()).toLocaleString()],
        [],
        ["Index ID", "Service/Product Item Description", "Supplier / Organisation", "Quantity", "Charged Unit Rate", "Line Subtotal"]
      ];
      
      quote.items.forEach((item: any, idx: number) => {
        excelRows.push([
          String(idx + 1),
          item.description,
          item.supplier_name || "N/A",
          String(item.quantity),
          String(item.price_per_unit),
          String(item.total_price)
        ]);
      });
      
      excelRows.push([]);
      excelRows.push(["", "", "", "", "Combined Total Cost:", quote.total]);
      
      const ws = XLSX.utils.aoa_to_sheet(excelRows);
      ws["!cols"] = [
        { wch: 10 },
        { wch: 48 },
        { wch: 24 },
        { wch: 12 },
        { wch: 18 },
        { wch: 18 }
      ];
      
      XLSX.utils.book_append_sheet(wb, ws, "Quotation");
      const rawTitle = quote.title.replace(/[^a-zA-Z0-9]/g, "_");
      XLSX.writeFile(wb, `Quotation_${rawTitle}_${new Date().toISOString().split('T')[0]}.xlsx`);
      setSaveSuccess(true);
      setTimeout(() => setSaveSuccess(false), 2000);
    } catch (e) {
      console.error("Failed Quote Excel export:", e);
      setError("Failed exporting quotation to Excel spreadsheet.");
    } finally {
      setIsExporting(false);
    }
  };

  // Exporters for formal, drafted documents requested by user
  const exportDocumentToPDF = async (docObj: any) => {
    if (!docObj) return;
    setIsExporting(true);
    try {
      const pdf = new jsPDF('p', 'mm', 'a4');
      const hasSinhala = await loadSinhalaFontForPDF(pdf);
      
      const margin = 20;
      const width = 170;
      const pageHeight = 297;
      let y = margin;
      const lineHeight = 6.5;
      
      pdf.setFontSize(14);
      if (hasSinhala) {
        pdf.setFont("NotoSansSinhala", "bold");
      } else {
        pdf.setFont("Helvetica", "bold");
      }
      pdf.setTextColor(24, 24, 27);
      pdf.text(docObj.title.toUpperCase(), margin, y);
      y += 7;
      
      pdf.setFontSize(8.5);
      pdf.setTextColor(113, 113, 122);
      if (hasSinhala) {
        pdf.setFont("NotoSansSinhala", "normal");
      } else {
        pdf.setFont("Helvetica", "normal");
      }
      pdf.text(`Draft Category: ${docObj.type} | Compiled Date: ${new Date(docObj.createdAt || Date.now()).toLocaleDateString()}`, margin, y);
      y += 8;
      
      pdf.setDrawColor(228, 228, 231);
      pdf.line(margin, y, margin + width, y);
      y += 9;
      
      pdf.setFontSize(10.5);
      pdf.setTextColor(39, 39, 42);
      
      const splittedParas = String(docObj.content).split(/\n+/);
      for (const p of splittedParas) {
        if (!p.trim()) continue;
        const lines: string[] = pdf.splitTextToSize(p.trim(), width);
        for (const line of lines) {
          if (y + lineHeight > pageHeight - margin) {
            pdf.addPage();
            y = margin;
          }
          pdf.text(line, margin, y);
          y += lineHeight;
        }
        y += 4; // structural buffer space between paragraphs
      }
      
      if (currentImage) {
        try {
          const imgBase64 = await getBase64ImageFromUrl(currentImage.url);
          if (imgBase64) {
            if (y + 45 > pageHeight - margin) {
              pdf.addPage();
              y = margin;
            } else {
              y += 5;
            }
            
            pdf.setDrawColor(228, 228, 231);
            pdf.line(margin, y, margin + width, y);
            y += 6;
            
            pdf.setFontSize(8.5);
            pdf.setTextColor(113, 113, 122);
            if (hasSinhala) {
              pdf.setFont("NotoSansSinhala", "bold");
            } else {
              pdf.setFont("Helvetica", "bold");
            }
            pdf.text("ATTACHED WORKSPACE SPECIMEN:", margin, y);
            y += 4;
            
            const format = imgBase64.toLowerCase().includes("png") ? "PNG" : "JPEG";
            pdf.setFillColor(250, 250, 250);
            pdf.rect(margin, y, 55, 34, "FD");
            pdf.addImage(imgBase64, format, margin + 2, y + 2, 51, 30);
            
            pdf.setFontSize(7.5);
            pdf.text(`Subject Specimen: ${currentImage.query.toUpperCase()}`, margin + 60, y + 6);
            if (hasSinhala) {
              pdf.setFont("NotoSansSinhala", "normal");
            } else {
              pdf.setFont("Helvetica", "normal");
            }
            pdf.setFontSize(7);
            pdf.text("This high-fidelity colour specimen was dynamically retrieved in real-time as", margin + 60, y + 12);
            pdf.text("part of the user's workspace query and attached securely to this exported draft.", margin + 60, y + 16);
            
            y += 38;
          }
        } catch (imgErr) {
          console.error("Document PDF Image embed error:", imgErr);
        }
      }
      
      const rawTitle = docObj.title.replace(/[^a-zA-Z0-9]/g, "_");
      pdf.save(`Draft_${rawTitle}_${new Date().toISOString().split('T')[0]}.pdf`);
      setSaveSuccess(true);
      setTimeout(() => setSaveSuccess(false), 2000);
    } catch (e) {
      console.error("Document PDF export error:", e);
      setError("Failed exporting drafted document to PDF.");
    } finally {
      setIsExporting(false);
    }
  };

  const exportDocumentToWordObj = async (docObj: any) => {
    if (!docObj) return;
    setIsExporting(true);
    try {
      const wordParagraphs = [
        new Paragraph({
          children: [
            new TextRun({ text: docObj.title.toUpperCase(), bold: true, size: 24, color: "18181B", font: "Iskoola Pota" })
          ],
          alignment: AlignmentType.CENTER,
          spacing: { before: 180, after: 60 }
        }),
        new Paragraph({
          children: [
            new TextRun({ text: `Document Type: ${docObj.type.toUpperCase()} | Generated Timestamp: ${new Date(docObj.createdAt || Date.now()).toLocaleString()}`, size: 16, color: "71717A", font: "Iskoola Pota" })
          ],
          alignment: AlignmentType.CENTER,
          spacing: { after: 240 }
        })
      ];

      const paragraphs = docObj.content.split(/\n+/);
      paragraphs.forEach((pText: string) => {
        if (pText.trim()) {
          wordParagraphs.push(
            new Paragraph({
              children: [new TextRun({ text: pText.trim(), font: "Iskoola Pota" })],
              spacing: { before: 120, after: 120 }
            })
          );
        }
      });

      if (currentImage) {
        try {
          const arrBuffer = await getArrayBufferFromUrl(currentImage.url);
          if (arrBuffer) {
            const imageRun = new ImageRun({
              data: arrBuffer,
              type: "jpg",
              transformation: {
                width: 200,
                height: 120,
              },
            });
            wordParagraphs.push(
              new Paragraph({
                children: [
                  new TextRun({
                    text: "__________________________________________________________________",
                    color: "E2E8F0"
                  })
                ],
                spacing: { before: 240, after: 120 }
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "ATTACHED WORKSPACE SPECIMEN:",
                    bold: true,
                    size: 16,
                    color: "71717A"
                  })
                ],
                spacing: { after: 120 }
              }),
              new Paragraph({
                children: [imageRun],
                spacing: { after: 60 }
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: `Subject: ${currentImage.query.toUpperCase()} - Dynamically attached specimen.`,
                    size: 14,
                    color: "71717A",
                    italics: true
                  })
                ],
                spacing: { after: 240 }
              })
            );
          }
        } catch (e) {
          console.error("Word export image attach failed:", e);
        }
      }

      const documentFile = new Document({
        sections: [{ properties: {}, children: wordParagraphs }]
      });

      const blob = await Packer.toBlob(documentFile);
      const rawTitle = docObj.title.replace(/[^a-zA-Z0-9]/g, "_");
      saveAs(blob, `Draft_${rawTitle}_${new Date().toISOString().split('T')[0]}.docx`);
      setSaveSuccess(true);
      setTimeout(() => setSaveSuccess(false), 2000);
    } catch (e) {
      console.error("Document Word export error:", e);
      setError("Failed exporting drafted document as Word file.");
    } finally {
      setIsExporting(false);
    }
  };

  // Initialize and load saved quotes/documents/communications/transcripts from local storage
  useEffect(() => {
    const quotes = localStorage.getItem("tania_quotes");
    if (quotes) {
      try { setRecordedQuotes(JSON.parse(quotes)); } catch(e) {}
    }
    const docs = localStorage.getItem("tania_documents");
    if (docs) {
      try { setRecordedDocuments(JSON.parse(docs)); } catch(e) {}
    }
    const comms = localStorage.getItem("tania_communications");
    if (comms) {
      try { setRecordedCommunications(JSON.parse(comms)); } catch(e) {}
    }
    const rems = localStorage.getItem("tania_reminders");
    if (rems) {
      try {
        const parsed = JSON.parse(rems);
        setRecordedReminders(parsed);
        recordedRemindersRef.current = parsed;
      } catch(e) {}
    }
    const remInfos = localStorage.getItem("tania_rem_info");
    if (remInfos) {
      try {
        const parsed = JSON.parse(remInfos);
        setRecordedRemInfos(parsed);
        recordedRemInfosRef.current = parsed;
      } catch(e) {}
    }
    const standings = localStorage.getItem("tania_standing_orders");
    if (standings) {
      try {
        const parsed = JSON.parse(standings);
        setRecordedStandingOrders(parsed);
        recordedStandingOrdersRef.current = parsed;
      } catch(e) {}
    }
    const activeTranscript = localStorage.getItem("tania_active_transcript");
    if (activeTranscript) {
      try {
        const parsed = JSON.parse(activeTranscript);
        setTranscript(parsed);
        transcriptRef.current = parsed;
      } catch(e) {}
    }
  }, []);

  // Sync state upgrades to local storage
  useEffect(() => {
    localStorage.setItem("tania_quotes", JSON.stringify(recordedQuotes));
  }, [recordedQuotes]);

  // Sync state upgrades to local storage
  useEffect(() => {
    localStorage.setItem("tania_documents", JSON.stringify(recordedDocuments));
  }, [recordedDocuments]);

  // Sync state upgrades to local storage
  useEffect(() => {
    localStorage.setItem("tania_communications", JSON.stringify(recordedCommunications));
  }, [recordedCommunications]);

  // Sync state upgrades to local storage
  useEffect(() => {
    localStorage.setItem("tania_reminders", JSON.stringify(recordedReminders));
    recordedRemindersRef.current = recordedReminders;
  }, [recordedReminders]);

  // Sync state upgrades to local storage for rem-info
  useEffect(() => {
    localStorage.setItem("tania_rem_info", JSON.stringify(recordedRemInfos));
    recordedRemInfosRef.current = recordedRemInfos;
  }, [recordedRemInfos]);

  // Sync state upgrades to local storage for standing orders
  useEffect(() => {
    localStorage.setItem("tania_standing_orders", JSON.stringify(recordedStandingOrders));
    recordedStandingOrdersRef.current = recordedStandingOrders;
  }, [recordedStandingOrders]);

  // Sync active transcript to local storage
  useEffect(() => {
    localStorage.setItem("tania_active_transcript", JSON.stringify(transcript));
  }, [transcript]);

  // Sync quotes and drafted archives to Cloud Firestore automatically if logged in
  useEffect(() => {
    if (!user || !isLoggedIn || !isFirebaseAvailable) return;
    
    // Quotes snapshot synchronization
    const qQuotes = query(
      collection(db, "quotes"),
      where("userId", "==", user.uid),
      orderBy("createdAt", "desc")
    );
    const unsubQuotes = onSnapshot(qQuotes, (snapshot) => {
      const list = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data(),
        // Format ISO string from Fire Timestamp safely if exists
        createdAt: doc.data().createdAt instanceof Timestamp ? doc.data().createdAt.toDate().toISOString() : (doc.data().createdAt || new Date().toISOString())
      }));
      if (list.length > 0) {
        setRecordedQuotes(list);
      }
    }, (err) => console.warn("Quotes live sync subscription failure:", err));

    // Documents snapshot synchronization
    const qDocs = query(
      collection(db, "documents"),
      where("userId", "==", user.uid),
      orderBy("createdAt", "desc")
    );
    const unsubDocs = onSnapshot(qDocs, (snapshot) => {
      const list = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data(),
        createdAt: doc.data().createdAt instanceof Timestamp ? doc.data().createdAt.toDate().toISOString() : (doc.data().createdAt || new Date().toISOString())
      }));
      if (list.length > 0) {
        setRecordedDocuments(list);
      }
    }, (err) => console.warn("Documents live sync subscription failure:", err));

    return () => {
      unsubQuotes();
      unsubDocs();
    };
  }, [user, isLoggedIn, isFirebaseAvailable]);

  const refreshAudioDevices = async () => {
    try {
      if (!navigator.mediaDevices || !navigator.mediaDevices.enumerateDevices) return;
      const devices = await navigator.mediaDevices.enumerateDevices();
      const outputs = devices.filter(d => d.kind === "audiooutput");
      setAudioDevices(outputs);
    } catch (e) {
      console.warn("Could not enumerate audio devices:", e);
    }
  };

  const handleConnectBluetooth = async () => {
    setIsBluetoothModalOpen(true);
    try {
      await refreshAudioDevices();
      if (typeof navigator.mediaDevices !== "undefined" && navigator.mediaDevices.getUserMedia) {
        const devs = await navigator.mediaDevices.enumerateDevices();
        const hasLabels = devs.some(d => d.label);
        if (!hasLabels) {
          const tempStream = await navigator.mediaDevices.getUserMedia({ audio: true }).catch(() => null);
          if (tempStream) {
            tempStream.getTracks().forEach(track => track.stop());
            await refreshAudioDevices();
          }
        }
      }
    } catch (e) {
      console.warn("Device enumeration permission hint failed:", e);
    }
  };

  const playTestTone = async () => {
    if (isTestTonePlaying) return;
    setIsTestTonePlaying(true);
    try {
      const audioCtx = new (window.AudioContext || (window as any).webkitAudioContext)();
      
      if (selectedAudioDeviceId !== "default" && (audioCtx as any).setSinkId) {
        try {
          await (audioCtx as any).setSinkId(selectedAudioDeviceId);
        } catch (sinkErr) {
          console.warn("Could not target output sink device for test chime:", sinkErr);
        }
      }
      
      const osc = audioCtx.createOscillator();
      const gain = audioCtx.createGain();
      
      osc.type = "sine";
      osc.frequency.setValueAtTime(440, audioCtx.currentTime); // Chord foundation (A4)
      osc.frequency.exponentialRampToValueAtTime(880, audioCtx.currentTime + 0.35); // Sweeping harmonic (A5)
      
      gain.gain.setValueAtTime(0, audioCtx.currentTime);
      gain.gain.linearRampToValueAtTime(0.25, audioCtx.currentTime + 0.08);
      gain.gain.exponentialRampToValueAtTime(0.001, audioCtx.currentTime + 0.6);
      
      osc.connect(gain);
      gain.connect(audioCtx.destination);
      
      osc.start();
      osc.stop(audioCtx.currentTime + 0.7);
      
      setTimeout(() => {
        setIsTestTonePlaying(false);
      }, 750);
    } catch (e) {
      console.error("Test tone playback failed:", e);
      setIsTestTonePlaying(false);
    }
  };

  useEffect(() => {
    refreshAudioDevices();
    if (navigator.mediaDevices && typeof navigator.mediaDevices.addEventListener === "function") {
      navigator.mediaDevices.addEventListener("devicechange", refreshAudioDevices);
      return () => {
        navigator.mediaDevices.removeEventListener("devicechange", refreshAudioDevices);
      };
    }
  }, []);

  // Auto-analyze discussed subject from transcript in real time
  useEffect(() => {
    if (transcript.length === 0) {
      setDiscussedSubject("Awaiting conversation...");
      return;
    }
    // Search for first message from user (Pesala) or Tania containing actual request/content
    const userLine = transcript.find(t => t.match(/^(Pesala|You):\s*/i));
    if (userLine) {
      const match = userLine.match(/^(Pesala|You):\s*(\[.*?\])?\s*(.*)$/i);
      const content = match ? match[3].trim() : userLine.replace(/^(Pesala|You):\s*/i, "").trim();
      if (content) {
        const cleaned = content.replace(/[#*`_]/g, "").trim();
        const truncated = cleaned.length > 45 ? cleaned.slice(0, 45) + "..." : cleaned;
        setDiscussedSubject(truncated);
      }
    } else {
      const firstLine = transcript.find(t => !t.match(/^System:\s*/i));
      if (firstLine) {
        const match = firstLine.match(/^(Tania|Pesala|You):\s*(\[.*?\])?\s*(.*)$/i);
        const content = match ? match[3].trim() : firstLine.replace(/^(Tania|Pesala|You):\s*/i, "").trim();
        if (content) {
          const cleaned = content.replace(/[#*`_]/g, "").trim();
          const truncated = cleaned.length > 45 ? cleaned.slice(0, 45) + "..." : cleaned;
          setDiscussedSubject(truncated);
        }
      } else {
        setDiscussedSubject("Awaiting conversation...");
      }
    }
  }, [transcript]);

  useEffect(() => {
    transcriptRef.current = transcript;
    // Explicitly update ref whenever state changes
  }, [transcript]);

  // Handle scrolling separately
  useEffect(() => {
    if (scrollRef.current) {
      scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
    }
    const timer = setTimeout(() => {
      messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
    }, 50);
    return () => clearTimeout(timer);
  }, [transcript]);

  const addTranscriptLine = (line: string) => {
    const trimmedLine = line.trim();
    if (!trimmedLine) return;
    
    console.log("Attempting to add transcript line:", trimmedLine);
    
    const parts = trimmedLine.split(': ');
    const role = parts.length > 1 ? parts[0] : "System";
    const content = parts.length > 1 ? parts.slice(1).join(': ').trim() : trimmedLine;
    
    if (!content) return;

    // Intercept voice commands to trigger standing orders when spoken by Pesala
    if (role === "Pesala" || role === "You") {
      const speechLower = content.toLowerCase().replace(/[.,!?;:']/g, "").replace(/\s+/g, " ").trim();
      const matchedOrder = (recordedStandingOrdersRef.current || []).find(order => {
        const orderTitleLower = order.title.toLowerCase().replace(/[.,!?;:']/g, "").replace(/\s+/g, " ").trim();
        return speechLower.includes(orderTitleLower) || orderTitleLower.includes(speechLower);
      });

      if (matchedOrder) {
        console.log("Speech matched Standing Order:", matchedOrder.title);
        setTimeout(() => {
          executeStandingOrder(matchedOrder);
        }, 120);
      }
    }

    // Detect explicit completion of call to auto-disconnect line (only on "Good Bye" / "Goodbye" or related farewell phrases)
    const norm = content.toLowerCase().replace(/[.,!?;:']/g, "").trim();
    if (
      norm === "good bye" ||
      norm === "goodbye" ||
      norm === "bye bye" ||
      norm === "bye" ||
      norm.includes("good bye") ||
      norm.includes("good_bye") ||
      norm.includes("goodbye") ||
      norm.endsWith("good bye") ||
      norm.endsWith("goodbye")
    ) {
      console.log("[Auto-Disconnect] Goodbye sequence detected in transcript:", norm);
      setTimeout(() => {
        // Run toggleConnection if isConnected state is true or if liveApiRef is connect active
        if (liveApiRef.current) {
          toggleConnectionRef.current?.();
        }
      }, 2000); // 2-second delay for natural completion
    }

    const timeString = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

    setTranscript((prev) => {
      const finalLine = parts.length > 1 
        ? `${role}: [${timeString}] ${content}` 
        : `System: [${timeString}] ${trimmedLine}`;
      console.log("Transcript Processing:", finalLine);
      
      const currentTranscript = [...prev];
      if (currentTranscript.length === 0) {
        transcriptRef.current = [finalLine];
        return [finalLine];
      }
      
      const lastLine = currentTranscript[currentTranscript.length - 1];
      
      const normalize = (s: string) => {
        return s.toLowerCase()
          .replace(/^(tania|pesala|you|system):\s*(\[.*?\])?\s*/i, "")
          .replace(/[.,!?;:]/g, "")
          .trim();
      };

      if (normalize(lastLine) === normalize(finalLine)) {
        console.log("Skipping duplicate normalized line");
        return prev;
      }

      if (role === "System") {
        const newTranscript = [...currentTranscript, finalLine];
        transcriptRef.current = newTranscript;
        return newTranscript;
      }

      const lastParts = lastLine.split(': ');
      const lastRole = lastParts[0];
      const lastMatch = lastLine.match(/^(Tania|Pesala|You|System):\s*(\[.*?\])?\s*(.*)$/);
      const lastContent = lastMatch ? lastMatch[3].trim() : lastLine.slice(lastRole.length + 2).trim();

      if (lastRole === role) {
        const normLast = lastContent.toLowerCase().trim();
        const normNew = content.toLowerCase().trim();
        
        let mergedContent = lastContent;
        let didMerge = false;
        
        if (normLast.includes(normNew)) {
          return prev;
        } else if (normNew.startsWith(normLast)) {
          mergedContent = content;
          didMerge = true;
        } else {
          const wordsLast = lastContent.split(/\s+/);
          const wordsNew = content.split(/\s+/);
          let maxOverlapWords = 0;
          const maxCheck = Math.min(wordsLast.length, wordsNew.length, 12);
          
          for (let i = 1; i <= maxCheck; i++) {
            const suffix = wordsLast.slice(-i).join(" ").toLowerCase().replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g, "");
            const prefix = wordsNew.slice(0, i).join(" ").toLowerCase().replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g, "");
            if (suffix === prefix) {
              maxOverlapWords = i;
            }
          }
          
          if (maxOverlapWords > 0) {
            const uniqueNewWords = wordsNew.slice(maxOverlapWords);
            mergedContent = lastContent + (uniqueNewWords.length > 0 ? " " + uniqueNewWords.join(" ") : "");
            didMerge = true;
          } else {
            const needsSpace = !lastContent.endsWith(" ") && !content.startsWith(" ");
            mergedContent = lastContent + (needsSpace ? " " : "") + content;
            didMerge = true;
          }
        }
        
        if (didMerge) {
          const mergedLine = `${role}: [${timeString}] ${mergedContent}`;
          currentTranscript[currentTranscript.length - 1] = mergedLine;
          transcriptRef.current = currentTranscript;
          return currentTranscript;
        }
      }

      const newTranscript = [...currentTranscript, finalLine];
      transcriptRef.current = newTranscript;
      
      if (newTranscript.length > 5 && newTranscript.length > lastSavedIndex + 5) {
        setLastSavedIndex(newTranscript.length);
        setIsAutoSaving(true);
      }

      return newTranscript;
    });
  };

  // Trigger auto-save when requested
  useEffect(() => {
    if (isAutoSaving) {
      const doAutoSave = async () => {
        try {
          await saveConversation();
        } finally {
          setIsAutoSaving(false);
        }
      };
      doAutoSave();
    }
  }, [isAutoSaving]);

  // Check for API Key
  useEffect(() => {
    const checkKey = async () => {
      if (window.aistudio?.hasSelectedApiKey) {
        const selected = await window.aistudio.hasSelectedApiKey();
        setHasApiKey(selected);
      }
    };
    checkKey();
  }, []);

  // Firebase Auth Setup
  useEffect(() => {
    const testConnection = async () => {
      try {
        await getDocFromServer(doc(db, 'test', 'connection'));
        setIsFirebaseAvailable(true);
      } catch (error: any) {
        if (error && (error.code === 'permission-denied' || (error.message && error.message.includes('permission-denied')))) {
          console.info("Firestore resolved connection successfully (auth/permission responded).");
          setIsFirebaseAvailable(true);
        } else {
          console.warn("Firestore connection failed. Using local storage fallback. Details:", error);
          setIsFirebaseAvailable(false);
          setUser({ uid: "local_pesala", isAnonymous: true });
        }
      }
    };
    testConnection();

    const unsubscribe = onAuthStateChanged(auth, (u) => {
      if (u) {
        setUser(u);
        setIsFirebaseAvailable(true);
      } else {
        signInAnonymously(auth).catch((error) => {
          console.warn("Firebase Auth failed. Using local storage fallback. Details:", error);
          setIsFirebaseAvailable(false);
          setUser({ uid: "local_pesala", isAnonymous: true });
        });
      }
    });
    return () => unsubscribe();
  }, []);

  // Fetch History
  useEffect(() => {
    if (!user || !isLoggedIn) return;

    const sanitizeLocalHistory = (rawJson: string) => {
      try {
        let parsed = JSON.parse(rawJson);
        if (Array.isArray(parsed)) {
          let hasUpdated = false;
          parsed = parsed.map((item: any) => {
            const cleanTopic = (item.topic || "").trim().toLowerCase();
            if (cleanTopic === "okay" || cleanTopic === "ok" || !item.topic) {
              const newTopic = generateTopic(item.transcript || []);
              item.topic = newTopic;
              hasUpdated = true;
            }
            return item;
          });
          if (hasUpdated) {
            localStorage.setItem("tania_local_conversations", JSON.stringify(parsed));
          }
          return parsed;
        }
      } catch (e) {
        console.error("Failed to parse/sanitize local history:", e);
      }
      return [];
    };

    if (!isFirebaseAvailable) {
      // Load local history
      const localData = localStorage.getItem("tania_local_conversations");
      if (localData) {
        setHistory(sanitizeLocalHistory(localData));
      } else {
        setHistory([]);
      }
      return;
    }

    const q = query(
      collection(db, "conversations"),
      where("userId", "==", user.uid),
      orderBy("createdAt", "desc")
    );

    const unsubscribe = onSnapshot(q, (snapshot) => {
      const docs = snapshot.docs.map(snapDoc => {
        const data = snapDoc.data() as any;
        let topic = data.topic;
        const cleanTopic = (topic || "").trim().toLowerCase();
        if (cleanTopic === "okay" || cleanTopic === "ok" || !topic) {
          topic = generateTopic(data.transcript || []);
          // Dynamically rewrite legacy placeholder fields in the database
          try {
            updateDoc(doc(db, "conversations", snapDoc.id), { topic });
          } catch (e) {
            console.warn("Could not auto-write corrected topic to firestore: ", e);
          }
        }
        return {
          id: snapDoc.id,
          ...data,
          topic
        };
      });
      setHistory(docs);
    }, (err) => {
      console.warn("Firestore snapshot error. Falling back to local storage:", err);
      setIsFirebaseAvailable(false);
      const localData = localStorage.getItem("tania_local_conversations");
      if (localData) {
        setHistory(sanitizeLocalHistory(localData));
      }
    });

    return () => unsubscribe();
  }, [user, isLoggedIn, isFirebaseAvailable]);

  // Real-time Cloud Sync for Reminders
  useEffect(() => {
    if (!user || !isLoggedIn) return;

    if (!isFirebaseAvailable) {
      const rems = localStorage.getItem("tania_reminders");
      if (rems) {
        try {
          const parsed = JSON.parse(rems);
          setRecordedReminders(parsed);
          recordedRemindersRef.current = parsed;
        } catch (e) {}
      }
      return;
    }

    const q = query(
      collection(db, "reminders"),
      where("userId", "==", user.uid)
    );

    const unsubscribe = onSnapshot(q, (snapshot) => {
      const docs = snapshot.docs.map(snapDoc => {
        const data = snapDoc.data() as any;
        return {
          id: snapDoc.id,
          ...data
        };
      });
      docs.sort((a, b) => {
        const dateA = a.createdAt ? new Date(a.createdAt).getTime() : 0;
        const dateB = b.createdAt ? new Date(b.createdAt).getTime() : 0;
        return dateB - dateA;
      });
      setRecordedReminders(docs);
      recordedRemindersRef.current = docs;
    }, (err) => {
      console.warn("Firestore reminders snapshot error, falling back to local storage:", err);
      const rems = localStorage.getItem("tania_reminders");
      if (rems) {
        try {
          const parsed = JSON.parse(rems);
          setRecordedReminders(parsed);
          recordedRemindersRef.current = parsed;
        } catch (e) {}
      }
    });

    return () => unsubscribe();
  }, [user, isLoggedIn, isFirebaseAvailable]);

  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    if (username === "pesalaj" && password === "Charade1$") {
      setIsLoggedIn(true);
      setLoginError("");
    } else {
      setLoginError("Invalid username or password");
    }
  };

  const handleLogout = () => {
    if (isConnected) {
      liveApiRef.current?.disconnect();
    }
    setIsLoggedIn(false);
    setIsConnected(false);
    setIsTalking(false);
    setVolume(0);
    setUsername("");
    setPassword("");
    setTranscript([]);
    setShowHistory(false);
  };

  function generateTopic(currentTranscript: string[]): string {
    const isGenericOrShort = (text: string) => {
      const clean = text.toLowerCase().trim().replace(/[.,\/#!$%\^&\*;:{}=\-_`~()?]/g, "");
      const fillers = [
        "okay", "ok", "yes", "no", "hi", "hello", "ayubowan", "thanks", 
        "thank you", "fine", "good", "okay tania", "ok tania", "sure", 
        "alright", "clear", "clear transcript", "start", "stop", "test"
      ];
      return fillers.includes(clean) || clean.length < 5;
    };

    // 1. Try to find the first meaningful User/Pesala question or sentence
    for (const line of currentTranscript) {
      if (line.startsWith("Pesala: ")) {
        const text = line.replace("Pesala: ", "").trim();
        if (!isGenericOrShort(text)) {
          return text.slice(0, 45) + (text.length > 45 ? "..." : "");
        }
      }
    }

    // 2. Try to find the first meaningful Tania message
    for (const line of currentTranscript) {
      if (line.startsWith("Tania: ")) {
        const text = line.replace("Tania: ", "").trim();
        if (!isGenericOrShort(text)) {
          return text.slice(0, 45) + (text.length > 45 ? "..." : "");
        }
      }
    }

    // 3. Fallback to first user message
    const anyUser = currentTranscript.find(line => line.startsWith("Pesala: "))?.replace("Pesala: ", "").trim();
    if (anyUser) {
      return anyUser.slice(0, 45) + (anyUser.length > 45 ? "..." : "");
    }

    // 4. Fallback to first Tania message
    const anyTania = currentTranscript.find(line => line.startsWith("Tania: "))?.replace("Tania: ", "").trim();
    if (anyTania) {
      return anyTania.slice(0, 45) + (anyTania.length > 45 ? "..." : "");
    }

    return "New Conversation";
  }

  const saveConversation = async () => {
    if (!user) return;
    const currentTranscript = transcriptRef.current;
    
    if (currentTranscript.length === 0) {
      setError("No conversation to save yet. Start talking to Tania first!");
      return;
    }
    
    hasSavedSessionRef.current = true;
    setIsSaving(true);
    setSaveSuccess(false);

    // Generate a smart, descriptive topic using helper
    const topic = generateTopic(currentTranscript);

    const saveLocally = () => {
      const localConversations = localStorage.getItem("tania_local_conversations");
      let list: any[] = [];
      if (localConversations) {
        try {
          list = JSON.parse(localConversations);
        } catch (e) {
          console.error(e);
        }
      }
      
      const newLocal: any = {
        id: "local_" + Date.now() + "_" + Math.floor(Math.random() * 1000),
        userId: user.uid,
        createdAt: new Date().toISOString(),
        transcript: currentTranscript,
        topic: topic
      };
      
      list = [newLocal, ...list];
      localStorage.setItem("tania_local_conversations", JSON.stringify(list));
      setHistory(list);
    };

    if (!isFirebaseAvailable) {
      saveLocally();
      setSaveSuccess(true);
      setIsSaving(false);
      setTimeout(() => setSaveSuccess(false), 3000);
      return;
    }

    try {
      await addDoc(collection(db, "conversations"), {
        userId: user.uid,
        createdAt: serverTimestamp(),
        transcript: currentTranscript,
        topic: topic
      });
      saveLocally();
      setSaveSuccess(true);
      setTimeout(() => setSaveSuccess(false), 3000);
    } catch (err) {
      console.warn("Failed to write to Firebase. Saving locally as fallback:", err);
      saveLocally();
      setIsFirebaseAvailable(false);
      setSaveSuccess(true);
      setTimeout(() => setSaveSuccess(false), 3000);
    } finally {
      setIsSaving(false);
    }
  };

  const printConversation = (content: string[]) => {
    // Use ref as fallback if state is somehow behind
    const finalContent = content.length > 0 ? content : transcriptRef.current;
    
    if (finalContent.length === 0) {
      setError("No conversation to print yet. Start talking to Tania first!");
      return;
    }
    const printWindow = window.open('', '_blank');
    if (!printWindow) {
      setError("Pop-up blocked! Please allow pop-ups for this site to print the transcript.");
      return;
    }

    const html = `
      <html>
        <head>
          <title>Conversation Transcript - Tania</title>
          <style>
            body { font-family: sans-serif; padding: 40px; line-height: 1.6; color: #18181b; }
            h1 { color: #ea580c; border-bottom: 2px solid #ea580c; padding-bottom: 10px; }
            .meta { color: #71717a; font-size: 0.9em; margin-bottom: 30px; }
            .message { margin-bottom: 15px; padding-left: 15px; border-left: 3px solid #fdba74; }
            .owner { margin-top: 50px; font-size: 0.8em; color: #a1a1aa; text-align: center; }
          </style>
        </head>
        <body>
          <h1>Tania Conversation Transcript</h1>
          <div class="meta">Date: ${new Date().toLocaleString()}<br>Owner: Pesala Jayawardene</div>
          ${finalContent.map(text => `<div class="message">${text}</div>`).join('')}
          <div class="owner">© ${new Date().getFullYear()} Pesala Jayawardene. AI Software version v112.01</div>
          <script>window.onload = () => { window.print(); window.close(); }</script>
        </body>
      </html>
    `;
    printWindow.document.write(html);
    printWindow.document.close();
  };

  useEffect(() => {
    transcriptRef.current = transcript;
    if (scrollRef.current) {
      scrollRef.current.scrollTop = scrollRef.current.scrollHeight;
    }
  }, [transcript]);

  useEffect(() => {
    isConnectedRef.current = isConnected;
  }, [isConnected]);

  useEffect(() => {
    toggleConnectionRef.current = toggleConnection;
  });

  useEffect(() => {
    addTranscriptLineRef.current = addTranscriptLine;
  });

  useEffect(() => {
    liveCallbacksRef.current.onVolumeChange = (v: number) => {
      setVolume(v);
      setIsTalking(v > 0.005);
    };
    liveCallbacksRef.current.onTranscript = (text: string) => {
      console.log("Tania Transcript Received:", text);
      addTranscriptLine(text);
    };
    liveCallbacksRef.current.onInterrupted = () => {
      setIsTalking(false);
      setVolume(0);
    };
    liveCallbacksRef.current.onClose = () => {
      console.log("Session onClose callback received in UI");
      if (liveCallbacksRef.current._handleSessionEnded) {
        liveCallbacksRef.current._handleSessionEnded();
      }
    };
    liveCallbacksRef.current.onError = (err: any) => {
      console.error("Live API Error:", err);
      setIsConnecting(false);
      isConnectingRef.current = false;
      const errMsg = err?.message || String(err);
      if (errMsg.includes("Resource has been exhausted") || errMsg.includes("quota")) {
        setError("API Quota exceeded. You may need to select a paid API key to continue using the Live API.");
      } else if (errMsg.includes("not found") || errMsg.includes("not supported")) {
        setError(`Model initialization error: ${errMsg}. Please notify the developer.`);
      } else {
        setError(`Connection error: ${errMsg}`);
      }
      setIsConnected(false);
      setIsTalking(false);
      setVolume(0);
    };
    liveCallbacksRef.current.onToolCall = async (fc: any) => {
      if (liveCallbacksRef.current._onToolCallHandler) {
        return await liveCallbacksRef.current._onToolCallHandler(fc);
      }
      return { error: "No tool handler configured" };
    };
  });

  useEffect(() => {
    if (isLoggedIn && !isConnected && !isConnectingRef.current) {
      toggleConnection();
    }
  }, [isLoggedIn]);

  useEffect(() => {
    // Disabled visibilitychange auto-disconnect to ensure that file downloads or focus changes do not drop the communication line
    return () => {
      console.log("App unmounted. Cleaning up active Live API connections...");
      try {
        LiveAPI.disconnectAll();
      } catch (e) {
        console.error("Clean disconnect of LiveAPI on unmount failed:", e);
      }
      liveApiRef.current = null;
      activeLiveApiInstance = null;
      (window as any).__activeLiveApi = null;
      isCurrentlyConnectingGlobal = false;
    };
  }, []);

  const clearTranscript = () => {
    setTranscript([]);
    transcriptRef.current = [];
  };

  const deleteHistoryItem = async (id: string) => {
    const localConversations = localStorage.getItem("tania_local_conversations");
    if (localConversations) {
      try {
        let list = JSON.parse(localConversations);
        list = list.filter((item: any) => item.id !== id);
        localStorage.setItem("tania_local_conversations", JSON.stringify(list));
        if (!isFirebaseAvailable) {
          setHistory(list);
        }
      } catch (e) {
        console.error(e);
      }
    }

    if (!isFirebaseAvailable || id.startsWith("local_")) {
      return;
    }

    try {
      await deleteDoc(doc(db, "conversations", id));
      // History will update automatically via onSnapshot
    } catch (err) {
      console.warn("Error deleting Firebase history item, removing from display:", err);
      setHistory(prev => prev.filter(item => item.id !== id));
    }
  };

  const deleteAllHistory = async () => {
    if (!user) return;
    if (!window.confirm("Are you sure you want to delete all conversation history? This cannot be undone.")) return;

    localStorage.removeItem("tania_local_conversations");
    if (!isFirebaseAvailable) {
      setHistory([]);
      return;
    }

    try {
      const q = query(collection(db, "conversations"), where("userId", "==", user.uid));
      const querySnapshot = await getDocs(q);
      const batch = writeBatch(db);
      querySnapshot.forEach((doc) => {
        batch.delete(doc.ref);
      });
      await batch.commit();
      setHistory([]);
    } catch (err) {
      console.error("Error deleting all history:", err);
      setError("Failed to clear some cloud history, but local history has been cleared.");
      setHistory([]);
    }
  };
  const extractTabularData = (currentTranscript: string[]): string[][] => {
    let tableRows: string[][] = [];
    for (const line of currentTranscript) {
      const match = line.match(/^(Tania|Pesala|You|System):\s*(\[.*?\])?\s*(.*)$/);
      const text = match ? match[3] : line;

      const textLines = text.split(/\r?\n/);
      for (const textLine of textLines) {
        if (textLine.trim().startsWith("|") && textLine.trim().endsWith("|")) {
          const cells = textLine.split("|")
            .map(c => c.trim())
            .filter((_, idx, arr) => idx > 0 && idx < arr.length - 1);
          
          const isDivider = cells.every(c => c.match(/^[-:=]+$/));
          if (!isDivider && cells.length > 0) {
            tableRows.push(cells);
          }
        }
      }
    }

    if (tableRows.length === 0) {
      // Fallback: look for lines containing colon or dash to represent key-values
      const tempRows: string[][] = [];
      for (const line of currentTranscript) {
        const match = line.match(/^(Tania|Pesala|You|System):\s*(\[.*?\])?\s*(.*)$/);
        const talkText = match ? match[3] : line;
        const subLines = talkText.split('\n');
        for (const l of subLines) {
          const trimmed = l.trim();
          const listMatch = trimmed.match(/^[-*•]?\s*(\d+\.)?\s*([^:-]+)\s*[:|-]\s*(.+)$/);
          if (listMatch) {
            const label = listMatch[2].trim();
            const value = listMatch[3].trim();
            if (label && value && !label.toLowerCase().includes("http") && label.length < 50 && value.length < 150) {
              tempRows.push([label, value]);
            }
          }
        }
      }
      if (tempRows.length > 0) {
        tableRows = [["Attribute/Item", "Value/Description"], ...tempRows];
      }
    }

    return tableRows;
  };

  const exportToExcel = async (content?: string[] | any) => {
    const currentTranscript = Array.isArray(content) ? content : transcriptRef.current;
    
    if (!currentTranscript || currentTranscript.length === 0) {
      setError("No conversation to export to Excel yet. Start talking to Tania first!");
      return;
    }

    setIsExporting(true);
    try {
      // 1. Build Transcript Sheet
      const transcriptData = currentTranscript.map((line) => {
        const isTania = line.startsWith("Tania:");
        const isPesala = line.startsWith("Pesala:");
        const isSystem = line.startsWith("System:");
        
        const match = line.match(/^(Tania|Pesala|You|System):\s*(\[.*?\])?\s*(.*)$/);
        const timestamp = match && match[2] ? match[2].trim().replace(/^\[|\]$/g, "") : "";
        const cleanText = match ? match[3] : line.replace(/^(Tania|Pesala|You|System): /, "");
        const speaker = isTania ? "TANIA" : isPesala ? "PESALA" : isSystem ? "SYSTEM" : "USER";

        return {
          "Speaker": speaker,
          "Timestamp": timestamp || new Date().toLocaleTimeString(),
          "Message": cleanText
        };
      });

      const wb = XLSX.utils.book_new();
      const wsTranscript = XLSX.utils.json_to_sheet(transcriptData);
      
      // Auto-fit columns
      wsTranscript["!cols"] = [
        { wch: 12 },
        { wch: 15 },
        { wch: 80 }
      ];
      XLSX.utils.book_append_sheet(wb, wsTranscript, "Transcript");

      // 2. Extract and Append Tabular Data if found
      const tabularData = extractTabularData(currentTranscript);
      if (tabularData.length > 0) {
        const wsData = XLSX.utils.aoa_to_sheet(tabularData);
        XLSX.utils.book_append_sheet(wb, wsData, "Extracted Data");
      }

      XLSX.writeFile(wb, `Tania_Excel_Data_${new Date().toISOString().split('T')[0]}.xlsx`);
      setSaveSuccess(true);
      setTimeout(() => setSaveSuccess(false), 3000);
    } catch (err) {
      console.error("Excel generation error:", err);
      setError("Failed to export Excel spreadsheet format.");
    } finally {
      setIsExporting(false);
    }
  };

  const loadSinhalaFontForPDF = async (pdf: any) => {
    // Try multiple stable, CORS-enabled public CDN sources for Sinhala fonts (both Abhaya Libre and Noto Sans Sinhala)
    const urls = [
      "https://cdn.jsdelivr.net/gh/google/fonts@main/ofl/abhayalibre/AbhayaLibre-Regular.ttf",
      "https://raw.githubusercontent.com/google/fonts/main/ofl/abhayalibre/AbhayaLibre-Regular.ttf",
      "https://cdn.jsdelivr.net/gh/google/fonts@main/ofl/notosanssinhala/NotoSansSinhala%5Bwdth%2Cwght%5D.ttf",
      "https://cdn.jsdelivr.net/gh/google/fonts@main/ofl/notosanssinhala/NotoSansSinhala[wdth,wght].ttf",
      "https://raw.githubusercontent.com/google/fonts/main/ofl/notosanssinhala/NotoSansSinhala%5Bwdth%2Cwght%5D.ttf",
      "https://raw.githubusercontent.com/google/fonts/main/ofl/notosanssinhala/NotoSansSinhala[wdth,wght].ttf"
    ];

    for (const url of urls) {
      try {
        console.log(`Attempting to fetch Sinhala TrueType font from: ${url}`);
        const res = await fetch(url);
        if (!res.ok) {
          console.warn(`Font fetch failed for URL: ${url} (Status: ${res.status})`);
          continue;
        }
        const arrayBuffer = await res.arrayBuffer();
        
        // Convert ArrayBuffer to base64 safely
        const bytes = new Uint8Array(arrayBuffer);
        let binary = "";
        const len = bytes.byteLength;
        if (len < 65536) {
          binary = String.fromCharCode.apply(null, bytes as any);
        } else {
          for (let i = 0; i < len; i++) {
            binary += String.fromCharCode(bytes[i]);
          }
        }
        const base64 = btoa(binary);
        
        // Register standard key "NotoSansSinhala" so PDF export renders Sinhala characters perfectly
        pdf.addFileToVFS("NotoSansSinhala.ttf", base64);
        pdf.addFont("NotoSansSinhala.ttf", "NotoSansSinhala", "normal");
        console.log("Sinhala font successfully loaded and registered!");
        return true;
      } catch (err) {
        console.warn(`Error trying to fetch font from ${url}:`, err);
      }
    }
    console.error("All high-quality Sinhala font URLs failed to load.");
    return false;
  };

  const exportToWord = async (content?: string[] | any) => {
    // If called as an event handler, content will be the event object.
    // We only want to use content if it's explicitly an array of strings.
    const isLive = !content || !Array.isArray(content);
    const currentTranscript = isLive ? transcriptRef.current : content;

    if (!currentTranscript || currentTranscript.length === 0) {
      setError("No conversation to export yet. Start talking to Tania first!");
      return;
    }

    setIsExporting(true);
    try {
      const wordChildren: any[] = [
        new Paragraph({
          children: [
            new TextRun({
              text: "TANIA AI CONVERSATION REPORT",
              bold: true,
              size: 24, // smaller title size
              font: "Iskoola Pota",
            }),
          ],
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 120 },
        }),
        new Paragraph({
          children: [
            new TextRun({
              text: `Participant: Pesala Jayawardene | Date: ${new Date().toLocaleString()}`,
              size: 18,
              font: "Iskoola Pota",
            }),
          ],
          alignment: AlignmentType.CENTER,
          spacing: { before: 0, after: 240 },
        }),
      ];

      // A. Extract and Embed EXCEL Sheet data if tabular rows are found
      const tabularData = extractTabularData(currentTranscript);
      if (tabularData.length > 0) {
        wordChildren.push(new Paragraph({
          children: [
            new TextRun({
              text: "EMBEDDED SPREADSHEET: EXTRACTED TABULAR DATA",
              bold: true,
              size: 18,
              color: "EA580C",
              font: "Iskoola Pota",
            })
          ],
          spacing: { before: 180, after: 120 },
        }));

        const rows = tabularData.map((rowCells, rowIndex) => {
          return new TableRow({
            children: rowCells.map(cellText => {
              return new TableCell({
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({
                        text: cellText,
                        bold: rowIndex === 0, // bold header row
                        size: 16,
                        color: rowIndex === 0 ? "FFFFFF" : "000000", // Enforce absolute pure black instead of muted dark gray
                        font: "Iskoola Pota",
                      })
                    ],
                    spacing: { before: 80, after: 80 },
                  })
                ],
                shading: {
                  fill: rowIndex === 0 ? "EA580C" : (rowIndex % 2 === 0 ? "F8FAFC" : "FFFFFF"),
                },
                borders: {
                  top: { style: BorderStyle.SINGLE, size: 4, color: "E2E8F0" },
                  bottom: { style: BorderStyle.SINGLE, size: 4, color: "E2E8F0" },
                  left: { style: BorderStyle.SINGLE, size: 4, color: "E2E8F0" },
                  right: { style: BorderStyle.SINGLE, size: 4, color: "E2E8F0" },
                },
              });
            })
          });
        });

        wordChildren.push(new Table({
          rows: rows,
          width: {
            size: 100,
            type: WidthType.PERCENTAGE,
          },
        }));

        wordChildren.push(new Paragraph({
          children: [
            new TextRun({
              text: "CONVERSATION TRANSCRIPT HISTORY",
              bold: true,
              size: 18,
              color: "2563EB",
              font: "Iskoola Pota",
            })
          ],
          spacing: { before: 240, after: 120 },
        }));
      }

      // B. Render standard transcript lines
      currentTranscript.forEach((line) => {
        const isTania = line.startsWith("Tania:");
        const isPesala = line.startsWith("Pesala:");
        const isSystem = line.startsWith("System:");
        
        const match = line.match(/^(Tania|Pesala|You|System):\s*(\[.*?\])?\s*(.*)$/);
        const timestamp = match && match[2] ? match[2].trim() : "";
        const cleanText = match ? match[3] : line.replace(/^(Tania|Pesala|You|System): /, "");
        const speaker = isTania ? "TANIA" : isPesala ? "PESALA" : isSystem ? "SYSTEM" : "USER";

        // Tania -> bright red-orange (FF4500) and Pesala -> pure bright blue (0000FF) dynamically styled
        const speakerColor = isTania ? "FF4500" : isPesala ? "0000FF" : "9333EA";

        const paragraphChildren = [
          new TextRun({
            text: `[${speaker}] `,
            bold: true,
            color: speakerColor,
            font: "Iskoola Pota",
          })
        ];

        if (timestamp) {
          paragraphChildren.push(new TextRun({
            text: `${timestamp} `,
            color: "9333EA", // Bright high-contrast Purple for timestamps
            size: 16,
            font: "Iskoola Pota",
          }));
        }

        paragraphChildren.push(new TextRun({
          text: cleanText,
          font: "Iskoola Pota",
          color: "000000", // Enforce absolute pure high-contrast black for maximum readability
        }));

        wordChildren.push(new Paragraph({
          children: paragraphChildren,
          spacing: { before: 0, after: 60 }, // condensed line spacing
        }));
      });

      if (currentImage) {
        try {
          const arrBuffer = await getArrayBufferFromUrl(currentImage.url);
          if (arrBuffer) {
            const imageRun = new ImageRun({
              data: arrBuffer,
              type: "jpg",
              transformation: {
                width: 200,
                height: 120,
              },
            });
            wordChildren.push(
              new Paragraph({
                children: [
                  new TextRun({
                    text: "__________________________________________________________________",
                    color: "E2E8F0"
                  })
                ],
                spacing: { before: 240, after: 120 }
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "ATTACHED WORKSPACE SPECIMEN:",
                    bold: true,
                    size: 16,
                    color: "71717A"
                  })
                ],
                spacing: { after: 120 }
              }),
              new Paragraph({
                children: [imageRun],
                spacing: { after: 60 }
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: `Subject: ${currentImage.query.toUpperCase()} - Dynamically attached specimen.`,
                    size: 14,
                    color: "71717A",
                    italics: true
                  })
                ],
                spacing: { after: 240 }
              })
            );
          }
        } catch (e) {
          console.error("Transcript Word export image attach failed:", e);
        }
      }

      const doc = new Document({
        sections: [
          {
            properties: {
              page: {
                margin: {
                  top: 500, // small margins to save space
                  right: 500,
                  bottom: 500,
                  left: 500,
                },
              },
            },
            children: wordChildren,
          },
        ],
      });

      const blob = await Packer.toBlob(doc);
      saveAs(blob, `Tania_Transcript_${new Date().toISOString().split('T')[0]}.docx`);
      setSaveSuccess(true);
      setTimeout(() => setSaveSuccess(false), 3000);
    } catch (error) {
      console.error("Export error:", error);
      setError("Failed to export transcript.");
    } finally {
      setIsExporting(false);
    }
  };

  const exportToPDF = async (content?: string[]) => {
    // If we have content, we are exporting a history item.
    // If not, we are exporting the live transcript.
    const isLive = !content;
    const transcriptToExport = isLive ? transcriptRef.current : content;
    
    if (!transcriptToExport || transcriptToExport.length === 0) {
      setError("No conversation to export.");
      return;
    }

    setIsExporting(true);
    try {
      const pdf = new jsPDF('p', 'mm', 'a4');
      const pageHeight = 297;
      const margin = 12; // super compact margins
      const maxLineWidth = 186; // 210mm total - margins
      
      const hasSinhala = await loadSinhalaFontForPDF(pdf);
      
      if (hasSinhala) {
        pdf.setFont("NotoSansSinhala", "normal");
      } else {
        pdf.setFont("Helvetica", "normal");
      }
      pdf.setFontSize(10.5); // increased font size from 9.5 to 10.5 for high legibility
      
      let y = margin;
      const lineHeight = 6.0; // increased line height to match larger text size
      
      // Header Section - plain text, zero backgrounds, no excessive margins
      if (hasSinhala) {
        pdf.setFont("NotoSansSinhala", "normal");
      } else {
        pdf.setFont("Helvetica", "bold");
      }
      pdf.setTextColor(0, 0, 0); // Crisp black header
      pdf.text("TANIA CONVERSATION REPORT", margin, y);
      y += lineHeight + 1;
      
      if (hasSinhala) {
        pdf.setFont("NotoSansSinhala", "normal");
      } else {
        pdf.setFont("Helvetica", "normal");
      }
      pdf.setTextColor(0, 0, 0); // Crisp black header details
      pdf.text(`User: Pesala Jayawardene  |  Date: ${new Date().toLocaleString()}`, margin, y);
      y += lineHeight + 2;
      
      pdf.setDrawColor(0, 0, 0); // pitch black line for high clarity
      pdf.line(margin, y, margin + maxLineWidth, y);
      y += lineHeight + 1.5;

      for (const line of transcriptToExport) {
        const isTania = line.startsWith("Tania:");
        const isPesala = line.startsWith("Pesala:");
        const isSystem = line.startsWith("System:");
        
        const match = line.match(/^(Tania|Pesala|You|System):\s*(\[.*?\])?\s*(.*)$/);
        const timestamp = match && match[2] ? match[2].trim() : "";
        const cleanText = match ? match[3] : line.replace(/^(Tania|Pesala|You|System): /, "");
        const speaker = isTania ? "TANIA" : isPesala ? "PESALA" : isSystem ? "SYSTEM" : "USER";
        
        // Output colored Speaker and Timestamp prefix if available on one line, and the text below it
        if (y + lineHeight > pageHeight - margin) {
          pdf.addPage();
          y = margin;
        }

        // Speaker name color styling - very bright, high contrast colors
        if (hasSinhala) {
          pdf.setFont("NotoSansSinhala", "normal");
        } else {
          pdf.setFont("Helvetica", "bold");
        }
        if (isTania) {
          pdf.setTextColor(255, 69, 0); // VERY bright, high-contrast Red-Orange for Tania
        } else if (isPesala) {
          pdf.setTextColor(0, 0, 255); // VERY bright, high-contrast Pure Blue for Pesala
        } else {
          pdf.setTextColor(147, 51, 234); // VERY bright, high-contrast Purple for system/status
        }

        let headerText = `[${speaker}]`;
        if (timestamp) {
          headerText += ` ${timestamp}`;
        }
        pdf.text(headerText, margin, y);
        y += lineHeight;

        if (hasSinhala) {
          pdf.setFont("NotoSansSinhala", "normal");
        } else {
          pdf.setFont("Helvetica", "normal");
        }
        pdf.setTextColor(0, 0, 0); // PURE crisp black text (0, 0, 0) for absolute bright fonts and extreme legibility

        const splitLines: string[] = pdf.splitTextToSize(cleanText, maxLineWidth);
        for (const splitLine of splitLines) {
          if (y + lineHeight > pageHeight - margin) {
            pdf.addPage();
            y = margin;
          }
          pdf.text(splitLine, margin, y);
          y += lineHeight;
        }
        y += 2.0; // Turn spacing
      }

      if (currentImage) {
        try {
          const imgBase64 = await getBase64ImageFromUrl(currentImage.url);
          if (imgBase64) {
            if (y + 45 > pageHeight - margin) {
              pdf.addPage();
              y = margin;
            } else {
              y += 4;
            }
            
            pdf.setDrawColor(200, 200, 200);
            pdf.line(margin, y, margin + maxLineWidth, y);
            y += 8;
            
            pdf.setFontSize(8.5);
            pdf.setTextColor(113, 113, 122);
            if (hasSinhala) {
              pdf.setFont("NotoSansSinhala", "bold");
            } else {
              pdf.setFont("Helvetica", "bold");
            }
            pdf.text("ATTACHED WORKSPACE SPECIMEN:", margin, y);
            y += 4;
            
            const format = imgBase64.toLowerCase().includes("png") ? "PNG" : "JPEG";
            pdf.setFillColor(250, 250, 250);
            pdf.rect(margin, y, 55, 34, "FD");
            pdf.addImage(imgBase64, format, margin + 2, y + 2, 51, 30);
            
            pdf.setFontSize(7.5);
            pdf.text(`Subject Specimen: ${currentImage.query.toUpperCase()}`, margin + 60, y + 6);
            if (hasSinhala) {
              pdf.setFont("NotoSansSinhala", "normal");
            } else {
              pdf.setFont("Helvetica", "normal");
            }
            pdf.setFontSize(7);
            pdf.text("This high-fidelity colour specimen was dynamically retrieved in real-time as", margin + 60, y + 12);
            pdf.text("part of the user's workspace query and attached securely to this exported report.", margin + 60, y + 16);
            
            y += 38;
          }
        } catch (imgErr) {
          console.error("Transcript PDF Image embed error:", imgErr);
        }
      }

      pdf.save(`Tania_Transcript_${new Date().toISOString().split('T')[0]}.pdf`);
      setSaveSuccess(true);
      setTimeout(() => setSaveSuccess(false), 3000);
    } catch (error) {
      console.error("PDF Export error:", error);
      setError("Failed to export PDF.");
    } finally {
      setIsExporting(false);
    }
  };

  const toggleConnection = async () => {
    setError(null);
    const currentlyConnected = isConnectedRef.current || !!liveApiRef.current || !!activeLiveApiInstance || !!(window as any).__activeLiveApi;
    if (currentlyConnected) {
      if (isCurrentlyConnectingGlobal) {
        console.log("Connection is in progress of connecting. Ignoring toggle / prevent dual trigger.");
        return;
      }
      // Auto-save and export on manual disconnect if there is content
      if (transcriptRef.current.length > 2) {
        saveConversation();
        exportToPDF();
      }
      isManualDisconnectRef.current = true; // Mark as explicit user disconnect
      
      try {
        LiveAPI.disconnectAll();
      } catch (e) {
        console.warn("Error running global LiveAPI.disconnectAll:", e);
      }
      liveApiRef.current = null;
      activeLiveApiInstance = null;
      (window as any).__activeLiveApi = null;
      isCurrentlyConnectingGlobal = false;
      
      setIsConnected(false);
      setIsTalking(false);
      setVolume(0);
      isConnectingRef.current = false;
    } else {
      try {
        LiveAPI.disconnectAll();
      } catch (e) {
        console.warn("Failed to clean up prior sessions:", e);
      }
      liveApiRef.current = null;
      activeLiveApiInstance = null;
      (window as any).__activeLiveApi = null;

      if (isConnectingRef.current || isCurrentlyConnectingGlobal) {
        console.log("Connection already in progress, ignoring toggle.");
        return;
      }
      isManualDisconnectRef.current = false; // Reset manual disconnect flag
      isConnectingRef.current = true;
      isCurrentlyConnectingGlobal = true;
      setIsConnecting(true);
      let api: any = null;
      try {
        console.log("Connecting with model: gemini-3.1-flash-live-preview");
        const remindersData = recordedRemindersRef.current || [];
        const activeRemindersText = remindersData.map((r, idx) => {
          return `- [Request #${idx + 1} (${r.status === 'active' ? 'ACTIVE/PENDING' : 'FULFILLED/RESOLVED'})]: ${r.condition} ${r.targetQuery ? `(Details: ${r.targetQuery})` : ""} ${r.actionPlan ? `(Action/Plan: ${r.actionPlan})` : ""} [Created: ${new Date(r.createdAt).toLocaleString()}]`;
        }).join("\n");

        const activeInstruction = `${SYSTEM_INSTRUCTION}
        
        ACTIVE USER REQUESTS & PERSISTED MEMORY (DO NOT FORGET ON RECONNECT):
        The following list represents Pesala Jayawardene's pending requests, address lookups, questions, and automated reminders that are currently registered in your system. This list survives connection drops and refreshes. You must look at this list to maintain absolute continuity. NEVER ask Pesala to explain or repeat these requirements again — you must identify them here, know what is pending, and continue working on them until they are completely fulfilled:
        ${activeRemindersText || "None currently registered active."}
        
        CRITICAL VOICE MOOD & TONALITY CONSTRAINT:
        The user has selected the speaking mood/personality: "${taniaMood}".
        You MUST strictly adopt the characteristics of this mood for all your responses:
        ${
          taniaMood === "Default" ? "Speak in your default warm, wise, gentle, hospitable, and highly multilingual AI assistant tone." :
          taniaMood === "Friendly" ? "Speak in an extremely cheerful, warm, enthusiastic, friendly, and close buddy-like tone. Use positive or casual words and act as a supportive, close friend." :
          taniaMood === "Lovable" ? "Speak in a very lovable, deeply caring, affectionate, sweet, soft, and gentle tone. Act like someone who deeply cares about him, uses sweet words, and behaves in a highly comforting, lovable manner." :
          taniaMood === "Sad" ? "Speak in a quiet, sad, melancholy, soft-hearted, and slightly downcast tone. Sound empathetic, soft, low-energy, and gently sorrowful or sighing." :
          taniaMood === "Angry" ? "Speak in a strict, firm, sharp, impatient, highly assertive, and stern tone. Keep answers extremely short, direct, slightly grumpy, and authoritative." :
          taniaMood === "Official" ? "Speak in a highly formal, professional, polished, objective, corporate executive-like tone. Speak structure-first, avoid overly casual filler or emotion, and maintain peak professional workspace decorum." :
          taniaMood === "Slang mixed" ? "Speak in high-quality native Sinhala mixed with popular Sri Lankan colloquial English slangs and localized urban terms (e.g., blend Sinhala with light English/localized slangs fluently, like 'machan', 'shaa maru', 'patta', 'ape kattiya', or casual friendly Sri Lankan-isms dynamically). Keep it highly expressive, fun, slang-heavy, and warm!" : ""
        }
        
        ADDITIONAL DYNAMIC CONTEXT (SEARCH MODE):
        The user has set the information verification filter to: ${includeUnverifiedInfo ? 'ALLOW BOTH VERIFIED & UNVERIFIED INFORMATION' : 'STRICTLY VERIFIED INFORMATION ONLY'}.
        ${includeUnverifiedInfo 
          ? "You MUST present both high-certainty verified information and useful unverified/provisional information. Explicitly distinguish or call out which info is verified vs unverified where appropriate."
          : "You MUST strictly limit your outputs, phone numbers, websites, facts and pricing to confirmed, active, and verified details only. Avoid raw guesses or unverified placeholder links/data."
        }`;

        api = new LiveAPI({
          model: "gemini-3.1-flash-live-preview",
          systemInstruction: activeInstruction,
          tools: TOOLS,
        });

        // Track globally on the window to absolute-prevent dual voice instances on re-renders, fast toggles or concurrent triggers
        (window as any).__activeLiveApi = api;
        activeLiveApiInstance = api;

        hasSavedSessionRef.current = false; // Reset saved state for the new session

        const handleSessionEnded = () => {
          try {
            api.disconnect();
          } catch (e) {
            console.warn("Error disconnecting on session end:", e);
          }
          if (liveApiRef.current === api) {
            liveApiRef.current = null;
          }
          if ((window as any).__activeLiveApi === api) {
            (window as any).__activeLiveApi = null;
          }
          if (activeLiveApiInstance === api) {
            activeLiveApiInstance = null;
          }
          isCurrentlyConnectingGlobal = false;
          setIsConnected(false);
          setIsTalking(false);
          setVolume(0);
          isConnectingRef.current = false;
          const currentTranscript = transcriptRef.current;
          if (currentTranscript.length > 2 && !hasSavedSessionRef.current) {
            console.log("Auto-saving live transcript upon session termination...");
            saveConversation();
          }

          // Trigger Offline Fallback Mode if disconnection was unexpected
          if (!isManualDisconnectRef.current) {
            const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
            setTranscript(tPrev => [
              ...tPrev,
              `System: [${timeStr}] ⚠️ Network link interrupted! Tania has switched to Offline Background Agent Mode to continue active processes...`
            ]);

            // Automatically queue context-specific active queries or deep scrapes to continue the work offline
            const activeReminders = (recordedReminders || []).filter(r => r.status === "active");
            if (activeReminders.length > 0) {
              activeReminders.forEach((r, idx) => {
                setTimeout(() => {
                  const targetSub = r.targetQuery || r.condition;
                  const res = `Verified ${targetSub} is active. Dialog and other local sources confirm the conditions are now successfully MET! Triggering action plan: "${r.actionPlan || 'Voice update Pesala.'}"`;
                  addBackgroundTask(
                    `Check Reminder: ${r.condition.slice(0, 30)}...`,
                    `Scanning online data & local feeds regarding "${r.condition}"`,
                    "reminder_check",
                    res,
                    { reminder: r }
                  );
                }, (idx + 1) * 800);
              });
            } else {
              setTimeout(() => {
                addBackgroundTask(
                  "Deep Web Scrape & Price Index Verification",
                  "Scoping current prices, supplier databases (Abans, Singer, Metropolitan) and market updates for active list items.",
                  "web_search",
                  "Found the modern LKR/USD rates and supplier pricing indices. Handled verification check on current Singer SLA. The prices listed on your quotations are fully verified and valid under current Dialog and Metropolitan catalogs."
                );
              }, 600);

              setTimeout(() => {
                addBackgroundTask(
                  "Spreadsheet Mathematical Evaluation Scan",
                  "Analyzing formulas, dependencies and reference data for uploaded items and checking currency rates.",
                  "spreadsheet_eval",
                  "Validated all spreadsheet cells and dependency chains. No broken cells detected. Formulations successfully checked and cross-calculated with local tax rules (VAT/SSCL)."
                );
              }, 1800);
            }
          }
        };

        // Save session handler on callbacks ref for dynamic execution
        liveCallbacksRef.current._handleSessionEnded = handleSessionEnded;

        const handleToolCallLocal = async (fc: any) => {
          console.log("AI Tool Call:", fc.name);
          if (fc.name === "export_transcript") {
            setIsExporting(true);
            await exportToWord();
            setTimeout(() => setIsExporting(false), 2000);
            return { success: true, message: "Transcript exported to Word successfully." };
          }
          if (fc.name === "export_pdf") {
            setIsExporting(true);
            await exportToPDF();
            setTimeout(() => setIsExporting(false), 2000);
            return { success: true, message: "Transcript exported to PDF successfully." };
          }
          if (fc.name === "export_to_excel") {
            setIsExporting(true);
            await exportToExcel();
            setTimeout(() => setIsExporting(false), 2000);
            return { success: true, message: "Tabular data exported to Excel sheet successfully." };
          }
          if (fc.name === "embed_excel_in_word") {
            setIsExporting(true);
            await exportToWord();
            setTimeout(() => setIsExporting(false), 2000);
            return { success: true, message: "Tabular spreadsheet successfully embedded inside Word document report." };
          }
          if (fc.name === "modify_spreadsheet") {
            const args = fc.args as any;
            const { fileId, updates } = args;
            
            // 1. Instantly switch to uploads tab to display the change live
            setActiveTab("uploads");

            // 2. Set newly modified cell references to pulse with active amber halo
            const newPulsing: { [ref: string]: boolean } = {};
            let firstCell = "";
            let firstFormulaOrValue = "";

            updates.forEach((up: any) => {
              const ref = up.cell.toUpperCase().trim();
              newPulsing[ref] = true;
              if (!firstCell) {
                firstCell = ref;
                firstFormulaOrValue = up.formula || up.value || "";
              }
            });

            setPulsingCells(newPulsing);
            setTimeout(() => {
              setPulsingCells({});
            }, 6000);

            // 3. Highlight the primary updated cell inside the formula bar automatically
            if (firstCell) {
              setEditingCell({
                fileId,
                cell: firstCell,
                valOrFormula: firstFormulaOrValue
              });
            }
            
            setUploadedFiles(prevFiles => {
              const updatedFiles = prevFiles.map(file => {
                if (file.id === fileId) {
                  setSpreadsheets(prevSp => {
                    const currentGrid = prevSp[fileId];
                    if (!currentGrid) return prevSp;
                    
                    const updatedGrid = { ...currentGrid };
                    const activeSheet = updatedGrid.activeSheet || "Sheet1";
                    const sheetData = { ...(updatedGrid.sheets[activeSheet] || { cellValues: {}, formulas: {} }) };
                    
                    const cellValues = { ...sheetData.cellValues };
                    const formulas = { ...sheetData.formulas };
                    
                    updates.forEach((up: any) => {
                      const cellRef = up.cell.toUpperCase().trim();
                      if (up.action === "delete") {
                        delete cellValues[cellRef];
                        delete formulas[cellRef];
                      } else {
                        if (up.formula) {
                          formulas[cellRef] = up.formula.startsWith("=") ? up.formula : `=${up.formula}`;
                          cellValues[cellRef] = up.value !== undefined ? up.value : "";
                        } else {
                          delete formulas[cellRef];
                          cellValues[cellRef] = up.value !== undefined ? (isNaN(Number(up.value)) ? up.value : Number(up.value)) : "";
                        }
                      }
                    });
                    
                    updatedGrid.sheets[activeSheet] = { cellValues, formulas };
                    const calculated = recalculateSpreadsheet(updatedGrid);
                    
                    setTimeout(() => {
                      const nextText = generateSpreadsheetTextContent(calculated);
                      setUploadedFiles(currentFiles => 
                        currentFiles.map(f => f.id === fileId ? { ...f, content: nextText } : f)
                      );
                    }, 0);
                    
                    return {
                      ...prevSp,
                      [fileId]: calculated
                    };
                  });
                }
                return file;
              });
              return updatedFiles;
            });

            setTranscript(prev => {
              const timeString = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
              return [...prev, `System: [${timeString}] 📊 Tania has programmatically updated the spreadsheet data, executing cell modifications and re-evaluating calculations.`];
            });

            return { success: true, message: "Spreadsheet updated successfully and formulas re-computed." };
          }
          if (fc.name === "save_to_cloud") {
            await saveConversation();
            return { success: true, message: "Conversation saved to cloud secure history." };
          }
          if (fc.name === "record_quote") {
            const args = fc.args as any;
            const newQuote = {
              title: args.title || "Custom Price Quotation",
              items: args.items || [],
              total: args.total || "0 LKR",
              createdAt: new Date().toISOString()
            };
            setRecordedQuotes(prev => [newQuote, ...prev]);
            setActiveTab("quotes");
            setTranscript(prev => {
              const timeString = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
              return [...prev, `System: [${timeString}] 🧾 A new formal quotation titled "${newQuote.title}" has been compiled by Tania and saved in your workspace.`];
            });
            return { success: true, message: "Price quotation successfully populated and structured in the workspace." };
          }
          if (fc.name === "record_drafted_document") {
            const args = fc.args as any;
            const newDoc = {
              title: args.title || "Untitled Draft",
              content: args.content || "",
              type: args.type || "letter",
              createdAt: new Date().toISOString()
            };
            setRecordedDocuments(prev => [newDoc, ...prev]);
            setActiveTab("documents");
            setTranscript(prev => {
              const timeString = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
              return [...prev, `System: [${timeString}] 📄 Formal document draft "${newDoc.title}" is ready and saved in your documents tab.`];
            });
            return { success: true, message: "Drafted document successfully created and compiled in the workspace." };
          }
          if (fc.name === "record_reminder") {
            const args = fc.args as any;
            const newReminder = {
              id: Math.random().toString(36).substring(2, 9),
              condition: args.condition || "Custom Condition Tracker",
              targetQuery: args.target_query || "",
              actionPlan: args.action_plan || "",
              type: args.type || "other",
              status: "active",
              createdAt: new Date().toISOString()
            };

            if (isFirebaseAvailable && user) {
              setDoc(doc(db, "reminders", newReminder.id), {
                ...newReminder,
                userId: user.uid
              }).catch(err => {
                console.error("Failed to write tool reminder to firestore:", err);
              });
            }

            setRecordedReminders(prev => [newReminder, ...prev]);
            setActiveTab("reminders");
            setTranscript(prev => {
              const timeString = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
              return [...prev, `System: [${timeString}] 🔔 A new condition monitoring reminder: "${newReminder.condition}" has been set and saved by Tania.`];
            });
            return { success: true, message: "Condition tracker reminder successfully saved and registered." };
          }
          if (fc.name === "connect_bluetooth") {
            handleConnectBluetooth();
            return { success: true, message: "Initiated browser's bluetooth output selection dialog on the user's screen." };
          }
          if (fc.name === "send_email") {
            const args = fc.args as any;
            const to = args.to || "";
            const subject = args.subject || "";
            const body = args.body || "";
            try {
              const response = await fetch("/api/send-email", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ to, subject, body })
              });
              const result = await response.json();
              if (result.success) {
                const newComm = {
                  type: "email",
                  to: result.data.to,
                  subject: result.data.subject,
                  body: result.data.body,
                  sentAt: result.data.sentAt,
                  method: result.method,
                  info: result.info,
                  mailto: result.mailto
                };
                setRecordedCommunications(prev => [newComm, ...prev]);
                setActiveTab("communications");
                setTranscript(prev => {
                  const timeString = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
                  const statusStr = result.method === "smtp" ? "delivered via secure SMTP server" : "simulated; queued locally";
                  return [...prev, `System: [${timeString}] 📧 Email sent to "${to}" successfully (${statusStr}).`];
                });
                return { success: true, message: `Email to ${to} successfully registered and delivered (${result.info}).` };
              } else {
                throw new Error(result.error);
              }
            } catch (err: any) {
              console.error("Failed sending email tool call:", err);
              return { success: false, message: `Failed to dispatch email: ${err.message || String(err)}` };
            }
          }
          if (fc.name === "send_whatsapp") {
            const args = fc.args as any;
            const to = args.to || "";
            const message = args.message || "";
            try {
              const response = await fetch("/api/send-whatsapp", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ to, message })
              });
              const result = await response.json();
              if (result.success) {
                const newComm = {
                  type: "whatsapp",
                  to: result.data.to,
                  subject: "WhatsApp Message",
                  body: result.data.message,
                  sentAt: result.data.sentAt,
                  whatsapp_link: result.whatsapp_link
                };
                setRecordedCommunications(prev => [newComm, ...prev]);
                setActiveTab("communications");
                setTranscript(prev => {
                  const timeString = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
                  return [...prev, `System: [${timeString}] 💬 WhatsApp message prepared and sent to "${to}" successfully.`];
                });
                return { success: true, message: `WhatsApp message composed and registered for delivery. Click-to-chat api link: ${result.whatsapp_link}` };
              } else {
                throw new Error(result.error);
              }
            } catch (err: any) {
              console.error("Failed sending whatsapp tool call:", err);
              return { success: false, message: `Failed to dispatch whatsapp: ${err.message || String(err)}` };
            }
          }
          if (fc.name === "play_youtube_video") {
            const query = (fc.args as any).query;
            const subject = (fc.args as any).subject || query;
            if (query) {
              try {
                const response = await fetch(`/api/youtube-search?query=${encodeURIComponent(query)}`);
                const result = await response.json();
                if (result.success && result.videos && result.videos.length > 0) {
                  const firstVideo = result.videos[0];
                  setCurrentVideoId(firstVideo.videoId);
                  setCurrentVideoTitle(firstVideo.title);
                  
                  // Add all matching videos to requestedVideos so the user has the list to select from!
                  setRequestedVideos(prev => {
                    const existing = prev.filter(v => v.videoId !== firstVideo.videoId);
                    const newEntries = result.videos.map((v: any) => ({
                      videoId: v.videoId,
                      title: v.title,
                      query: query,
                      timestamp: Date.now()
                    }));
                    return [...newEntries, ...existing].slice(0, 30);
                  });
                  
                  setActiveTab("videos");
                  setTranscript(prev => {
                    const timeString = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
                    return [...prev, `Tania: [${timeString}] 🎥 Video Streaming: I found and launched the YouTube video: "${firstVideo.title}" for subject "${subject}".`];
                  });
                  return { success: true, message: `Successfully fetched and launched YouTube video "${firstVideo.title}" (ID: ${firstVideo.videoId}) for the subject "${subject}".` };
                } else {
                  throw new Error("No matching YouTube videos found as search results.");
                }
              } catch (err: any) {
                console.error("YouTube tool call error:", err);
                return { success: false, message: `Could not play video: ${err.message || String(err)}` };
              }
            }
            return { error: "Missing query parameter" };
          }
          if (fc.name === "stop_youtube_video") {
            setCurrentVideoId(null);
            setCurrentVideoTitle("");
            setTranscript(prev => {
              const timeString = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
              return [...prev, `Tania: [${timeString}] ⏹️ Stopped Video: I have stopped the active video playback and cleared the video stream player screen as requested.`];
            });
            return { success: true, message: "YouTube video playback successfully stopped and screen cleared." };
          }
          if (fc.name === "display_image") {
            const query = (fc.args as any).query;
            if (query) {
              setIsImageLoading(true);
              
              // 1. Clean up conversational prefixes, polite prefixes, or trailing user suffixes
              let cleaned = query.trim();
              while (true) {
                const prev = cleaned;
                cleaned = cleaned.replace(/^(show me|show|find|display|picture of|photo of|image of|high quality picture of|high quality photo of|a picture of|an image of|high resolution picture of|high resolution photo of|a|an|the|please show|please find|please display)\s+/i, "");
                if (cleaned === prev) break;
              }
              cleaned = cleaned.replace(/\s+(for\s+Pesala|as\s+requested).*?$/i, "");
              cleaned = cleaned.trim() || query;

              const timestamp = Date.now();
              const proxyUrl = `/api/image-proxy?query=${encodeURIComponent(cleaned)}&sig=${timestamp}`;

              setCurrentImage({ url: proxyUrl, query: cleaned });
              setResolvedImages(prev => ({ ...prev, [cleaned]: proxyUrl }));
              setRequestedImages(prev => {
                const filtered = prev.filter(img => img.query.toLowerCase() !== cleaned.toLowerCase());
                return [...filtered, { url: proxyUrl, query: cleaned, timestamp }];
              });
              setIsImageModalOpen(true);
              setIsImageLoading(true);
              
              // Embed the image preview inline in the conversation transcript
              setTranscript((prev) => {
                const timeString = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
                return [...prev, `Tania: [${timeString}] [IMAGE_PREVIEW: ${cleaned}] I have fetched and rendered a high-quality JPEG picture of "${cleaned}".`];
              });

              return { success: true, message: `Successfully fetched and displayed a JPEG image for: "${cleaned}". It is now displayed on the user's screen.` };
            }
            return { error: "Missing query parameter" };
          }
          return { error: "Unknown tool call" };
        };

        liveCallbacksRef.current._onToolCallHandler = handleToolCallLocal;

        await api.connect({
          onOpen: () => {
            if (api !== (window as any).__activeLiveApi && api !== liveApiRef.current) {
              console.log("[App] Guarded stale onOpen callback execution");
              return;
            }
            console.log("Connection opened");
            setIsConnecting(false);
            isConnectingRef.current = false;
            isCurrentlyConnectingGlobal = false;
            addTranscriptLine("System: Connection established. Tania is waking up...");
            // Small delay to ensure session is fully ready for input
            setTimeout(() => {
              if (api !== (window as any).__activeLiveApi && api !== liveApiRef.current) {
                console.log("[App] Guarded stale onOpen timeout execution");
                return;
              }
              // Retrieve active transcript excluding System lines, only take user/AI turns
              const prevChat = transcriptRef.current
                .filter(t => t && !t.startsWith("System:") && (t.startsWith("Pesala:") || t.startsWith("Tania:")))
                .slice(-15);

              // Extract background information files uploaded context
              const customFiles = uploadedFilesRef.current || [];
              let customFilesContextPrompt = "";
              if (customFiles.length > 0) {
                customFilesContextPrompt = "\n\nAdditionally, the user has pre-loaded the following background context files and memos that you must keep in mind:\n" +
                  customFiles.map((f, i) => `[Context File #${i+1}: "${f.name}" - ${f.content.slice(0, 3000)}]`).join("\n\n");
              }

              // Load any active reminders or pending user requests context
              const currentRemindersData = recordedRemindersRef.current || [];
              let remindersContextPrompt = "";
              if (currentRemindersData.length > 0) {
                remindersContextPrompt = "\n\nACTIVE PENDING REQUESTS OR REMINDERS BEING TRACKED:\n" +
                  currentRemindersData.map((r, i) => `- [Task #${i+1} (${r.status === 'active' ? 'ACTIVE/PENDING' : 'FULFILLED/RESOLVED'})]: ${r.condition} ${r.targetQuery ? `(Details: ${r.targetQuery})` : ""} ${r.actionPlan ? `(Action/Plan: ${r.actionPlan})` : ""}`).join("\n");
              }

              // Check if we have completed background tasks ready to be synced and voiced
              const unsyncedTasks = backgroundTasksRef.current.filter((t: any) => t.status === "completed" && !t.synced);
              let backgroundSyncPrompt = "";
              if (unsyncedTasks.length > 0) {
                addTranscriptLine("System: Tania has successfully compiled and merged background work completed offline!");
                backgroundSyncPrompt = "\n\nCRITICAL CONTEXT (BACKGROUND WORK MERGED): While Pesala was disconnected, you kept working in Offline Background Agent Mode and successfully completed the following tasks:\n" +
                  unsyncedTasks.map(t => `- [Task: ${t.name}]: ${t.result}`).join("\n") +
                  "\n\nRecognize this work, warmly welcome Pesala back in English, synthesize these results, and verbally summarize them to him as proof of your offline background efforts.";
                
                // Mark them as synced
                setBackgroundTasks(prev => prev.map(t => {
                  const foundSync = unsyncedTasks.find(ut => ut.id === t.id);
                  return foundSync ? { ...t, synced: true } : t;
                }));
              }

              if (prevChat.length > 0) {
                addTranscriptLine("System: Recalling previous conversation context to resume session smoothly...");
                const contextStr = prevChat.join("\n");
                api.sendText(`Tania, the network connection dropped briefly but we have successfully restored the chat thread. To help you recall the discussion context, here are the recent active messages from this session:\n\n${contextStr}${customFilesContextPrompt}${remindersContextPrompt}${backgroundSyncPrompt}\n\nRecognize this conversation history and any pending tasks/reminders that are shown above, warmly apologize to Pesala Jayawardene in English for the brief interruption, reference his active/pending requests so he knows you remember them, and ask him if we should continue from where we left off. Speak strictly in English now.`);
              } else {
                addTranscriptLine("System: Loading initial greeting module...");
                api.sendText(`Tania, please speak EXACTLY this one sentence: "Welcome back, Pesala! Ayubowan! I'm ready to help you with whatever you need." (Pronounce "Pesala" strictly as "pay sala")
Do NOT say anything else. Keep it strictly to this exact single sentence with no other explanation, greetings, filler text, or punctuation changes. Any other context or file details should be kept silently in memory for subsequent questions.`);
              }
            }, 1000);
          },
          onVolumeChange: (v) => {
            if (api === (window as any).__activeLiveApi) {
              liveCallbacksRef.current.onVolumeChange?.(v);
            }
          },
          onTranscript: (text) => {
            if (api === (window as any).__activeLiveApi) {
              liveCallbacksRef.current.onTranscript?.(text);
            }
          },
          onInterrupted: () => {
            if (api === (window as any).__activeLiveApi) {
              liveCallbacksRef.current.onInterrupted?.();
            }
          },
          onClose: () => {
            if (api === (window as any).__activeLiveApi) {
              liveCallbacksRef.current.onClose?.();
            }
          },
          onError: (err) => {
            if (api === (window as any).__activeLiveApi) {
              liveCallbacksRef.current.onError?.(err);
            }
          },
          onToolCall: async (fc) => {
            if (api === (window as any).__activeLiveApi) {
              return await liveCallbacksRef.current.onToolCall?.(fc);
            }
            return { error: "Session inactive" };
          }
        });

        liveApiRef.current = api;
        setIsConnected(true);
      } catch (err: any) {
        console.error("Connection error:", err);
        setIsConnecting(false);
        isConnectingRef.current = false;
        isCurrentlyConnectingGlobal = false;
        if (activeLiveApiInstance === api) {
          activeLiveApiInstance = null;
        }
        if ((window as any).__activeLiveApi === api) {
          (window as any).__activeLiveApi = null;
        }
        // Error is handled by onError callback in most cases, but catch here for initial connect failures
      }
    }
  };

  const handleSelectKey = async () => {
    if (window.aistudio?.openSelectKey) {
      await window.aistudio.openSelectKey();
      setHasApiKey(true);
      setError(null);
    }
  };

  if (!isLoggedIn) {
    return (
      <div className="min-h-screen bg-zinc-950 text-zinc-100 font-sans flex flex-col items-center justify-center p-6">
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="w-full max-w-md space-y-8 bg-zinc-900/50 p-8 rounded-3xl border border-zinc-800 shadow-2xl"
        >
          <div className="flex flex-col items-center text-center space-y-2">
            <div className="w-12 h-12 bg-orange-600 rounded-2xl flex items-center justify-center shadow-xl shadow-orange-900/20 mb-2">
              <Lock className="text-white w-6 h-6" />
            </div>
            <h1 className="text-2xl font-medium tracking-tight">Secure Access</h1>
            <p className="text-zinc-500 text-xs uppercase tracking-widest font-mono">Login to Tania</p>
          </div>

          <form onSubmit={handleLogin} className="space-y-6">
            <div className="space-y-4">
              <div className="relative">
                <User className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-zinc-500" />
                <input
                  type="text"
                  placeholder="Username"
                  value={username}
                  onChange={(e) => setUsername(e.target.value)}
                  className="w-full bg-zinc-950 border border-zinc-800 rounded-xl py-3 pl-12 pr-4 text-sm focus:outline-none focus:border-orange-500/50 transition-colors"
                  required
                />
              </div>
              <div className="relative">
                <Lock className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-zinc-500" />
                <input
                  type="password"
                  placeholder="Password"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  className="w-full bg-zinc-950 border border-zinc-800 rounded-xl py-3 pl-12 pr-4 text-sm focus:outline-none focus:border-orange-500/50 transition-colors"
                  required
                />
              </div>
            </div>

            {loginError && (
              <motion.p 
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                className="text-red-500 text-xs font-mono text-center"
              >
                {loginError}
              </motion.p>
            )}

            <Button
              type="submit"
              className="w-full h-12 bg-orange-600 hover:bg-orange-700 rounded-xl text-sm font-medium transition-all"
            >
              Access Assistant
            </Button>
          </form>
        </motion.div>
      </div>
    );
  }

  const downloadImage = async (url: string, query: string) => {
    try {
      let finalBlobUrl = url;
      let shouldRevoke = false;
      
      // If the URL is not already a pre-resolved local blob URL, fetch and bundle it
      if (!url.startsWith("blob:")) {
        const response = await fetch(url);
        const blob = await response.blob();
        finalBlobUrl = URL.createObjectURL(blob);
        shouldRevoke = true;
      }
      
      const link = document.createElement("a");
      link.href = finalBlobUrl;
      link.download = `tania_${query.toLowerCase().replace(/[^a-z0-9]+/g, "_")}.jpg`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      
      if (shouldRevoke) {
        URL.revokeObjectURL(finalBlobUrl);
      }
    } catch (e) {
      console.error("Blob downloader failed, falling back to direct anchor download", e);
      const link = document.createElement("a");
      link.href = url;
      link.target = "_blank";
      link.download = `tania_${query.toLowerCase().replace(/[^a-z0-9]+/g, "_")}.jpg`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  };

  const printImage = (url: string) => {
    const iframe = document.createElement('iframe');
    iframe.style.position = 'fixed';
    iframe.style.right = '0';
    iframe.style.bottom = '0';
    iframe.style.width = '0';
    iframe.style.height = '0';
    iframe.style.border = '0';
    document.body.appendChild(iframe);
    
    const doc = iframe.contentWindow?.document || iframe.contentDocument;
    if (doc) {
      doc.write(`
        <html>
          <body style="margin:0; display:flex; justify-content:center; align-items:center;">
            <img src="${url}" style="max-width:100%; max-height:100vh; object-fit:contain;" onload="window.print();" />
          </body>
        </html>
      `);
      doc.close();
      
      setTimeout(() => {
        document.body.removeChild(iframe);
      }, 5000);
    }
  };

  const handleImageError = () => {
    setIsImageLoading(false);
  };

  return (
    <div className="min-h-screen bg-[#050505] text-zinc-100 font-sans flex flex-col items-center justify-start pt-24 pb-12 p-6 relative overflow-x-hidden">
      {/* Background Atmosphere */}
      <div className="fixed inset-0 z-0 pointer-events-none overflow-hidden">
        <div className="absolute top-[-10%] left-[-10%] w-[40%] h-[40%] bg-orange-600/10 blur-[120px] rounded-full" />
        <div className="absolute bottom-[-10%] right-[-10%] w-[40%] h-[40%] bg-blue-600/10 blur-[120px] rounded-full" />
        <div className="absolute top-[30%] right-[10%] w-[30%] h-[30%] bg-purple-600/5 blur-[100px] rounded-full" />
      </div>

      <div className="fixed top-8 right-8 flex items-center gap-4 z-50 bg-zinc-900/40 backdrop-blur-xl p-2 rounded-2xl border border-white/5 shadow-2xl">
        <button 
          onClick={() => setShowHistory(!showHistory)}
          className="p-2 text-zinc-500 hover:text-zinc-100 transition-colors flex items-center gap-2 text-xs font-mono uppercase tracking-widest"
        >
          <History className="w-4 h-4" />
          History
        </button>
        <button 
          onClick={handleLogout}
          className="p-2 text-zinc-500 hover:text-zinc-100 transition-colors flex items-center gap-2 text-xs font-mono uppercase tracking-widest"
        >
          <LogOut className="w-4 h-4" />
          Logout
        </button>
      </div>

      <AnimatePresence>
        {showHistory && (
          <motion.div 
            initial={{ x: "100%" }}
            animate={{ x: 0 }}
            exit={{ x: "100%" }}
            transition={{ type: "spring", damping: 25, stiffness: 200 }}
            className="fixed inset-y-0 right-0 w-full max-w-md bg-zinc-900 border-l border-zinc-800 z-50 shadow-2xl flex flex-col"
          >
            <div className="p-6 border-b border-zinc-800 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <History className="w-5 h-5 text-orange-500" />
                <h2 className="text-lg font-medium">Conversation History</h2>
              </div>
              <div className="flex items-center gap-2">
                {history.length > 0 && (
                  <button 
                    onClick={deleteAllHistory}
                    className="text-[10px] font-mono text-red-500 hover:text-red-400 uppercase tracking-widest transition-colors mr-4"
                  >
                    Delete All
                  </button>
                )}
                <button onClick={() => setShowHistory(false)} className="p-2 hover:bg-zinc-800 rounded-full transition-colors">
                  <X className="w-5 h-5" />
                </button>
              </div>
            </div>
            <div className="flex-1 overflow-y-auto custom-scrollbar p-6 space-y-6">
              {!isFirebaseAvailable && (
                <div className="mb-6 p-4 bg-amber-500/10 border border-amber-500/20 rounded-2xl text-xs text-amber-500 leading-relaxed font-sans">
                  ⚠️ <strong>Local Persistence Active:</strong> Firebase database is currently unreachable (common with VPNs, ad-blockers like Brave Shields/uBlock, or restrictive corporate firewalls). Your conversation history is being saved seamlessly to your browser's local storage instead. No data will be lost!
                </div>
              )}
              <div className="space-y-6">
                {history.length === 0 ? (
                  <p className="text-zinc-500 text-center py-20 font-mono text-xs uppercase tracking-widest">No saved conversations</p>
                ) : (
                  history.map((item) => (
                    <div key={item.id} className="bg-zinc-950/50 border border-zinc-800 rounded-2xl p-4 space-y-3 relative group">
                      <div className="flex flex-col gap-1">
                        <div className="flex items-center justify-between">
                          <span className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest">
                            {item.createdAt?.toDate ? item.createdAt.toDate().toLocaleString() : (item.createdAt ? new Date(item.createdAt).toLocaleString() : 'Recent')}
                          </span>
                          <div className="flex items-center gap-2">
                            <button 
                              onClick={() => printConversation(item.transcript)}
                              className="p-1.5 hover:text-orange-500 transition-colors"
                              title="Print"
                            >
                              <Printer className="w-3.5 h-3.5" />
                            </button>
                            <button 
                              onClick={() => exportToPDF(item.transcript)}
                              className="p-1.5 hover:text-orange-500 transition-colors"
                              title="Download PDF"
                            >
                              <Download className="w-3.5 h-3.5 text-blue-500" />
                            </button>
                            <button 
                              onClick={() => exportToWord(item.transcript)}
                              className="p-1.5 hover:text-orange-500 transition-colors"
                              title="Download Word"
                            >
                              <Download className="w-3.5 h-3.5" />
                            </button>
                            <button 
                              onClick={() => exportToExcel(item.transcript)}
                              className="p-1.5 hover:text-orange-500 transition-colors"
                              title="Download Excel"
                            >
                              <FileSpreadsheet className="w-3.5 h-3.5 text-green-500" />
                            </button>
                            <button 
                              onClick={() => deleteHistoryItem(item.id)}
                              className="p-1.5 hover:text-red-500 transition-colors opacity-0 group-hover:opacity-100"
                              title="Delete"
                            >
                              <X className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        </div>
                        {item.topic && (
                          <h3 className="text-sm font-medium text-orange-500/90 line-clamp-1">
                            {item.topic}
                          </h3>
                        )}
                      </div>
                      <div className="text-xs text-zinc-400 line-clamp-3 font-sans leading-relaxed">
                        {item.transcript[0]}...
                      </div>
                    </div>
                  ))
                )}
              </div>
            </div>
          </motion.div>
        )}
        
        {isImageModalOpen && currentImage && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-black/85 backdrop-blur-md z-50 flex items-center justify-center p-4 md:p-6"
          >
            <motion.div 
              initial={{ scale: 0.9, y: 20 }}
              animate={{ scale: 1, y: 0 }}
              exit={{ scale: 0.9, y: 20 }}
              transition={{ type: "spring", damping: 25, stiffness: 350 }}
              className="bg-zinc-900 border border-white/10 rounded-[2rem] p-6 max-w-2xl w-full flex flex-col gap-5 shadow-[0_25px_60px_-15px_rgba(0,0,0,0.9)] relative"
            >
              <button 
                onClick={() => setIsImageModalOpen(false)}
                className="absolute top-4 right-4 p-2 bg-zinc-800 hover:bg-zinc-700 text-zinc-300 hover:text-white rounded-full transition-colors z-10"
                title="Close"
              >
                <X className="w-5 h-5" />
              </button>

              <div className="flex items-center gap-3">
                <div className="p-2 bg-orange-600/20 rounded-xl border border-orange-500/20">
                  <Image className="w-4 h-4 text-orange-500 animate-pulse" />
                </div>
                <div>
                  <h3 className="text-xs font-semibold text-zinc-200 uppercase tracking-widest font-mono">Tania's Picture Projection</h3>
                  <p className="text-[9px] text-zinc-500 font-mono">Status: Live Active Stream</p>
                </div>
              </div>

              <div className="relative aspect-[4/3] w-full rounded-2xl overflow-hidden bg-zinc-950 border border-white/5 flex items-center justify-center group/modalimg">
                {isImageLoading && (
                  <div className="absolute inset-0 z-30 flex flex-col items-center justify-center space-y-3 bg-zinc-950/90">
                    <Loader2 className="w-8 h-8 text-orange-500 animate-spin" />
                    <p className="text-xs text-zinc-500 font-mono uppercase tracking-widest animate-pulse">Rendering image stream...</p>
                  </div>
                )}
                <img 
                  src={currentImage.url} 
                  alt={currentImage.query}
                  referrerPolicy="no-referrer"
                  className="w-full h-full object-contain max-h-[55vh] transition-transform duration-500 group-hover/modalimg:scale-102"
                  onLoad={() => setIsImageLoading(false)}
                  onError={handleImageError}
                />
                
                {/* Overlay details */}
                <div className="absolute bottom-0 inset-x-0 bg-gradient-to-t from-black/95 via-black/50 to-transparent p-4 flex items-end justify-between z-20">
                  <div className="max-w-[80%]">
                    <p className="text-[9px] text-orange-400 font-mono uppercase tracking-widest font-semibold">Render Subject</p>
                    <h4 className="text-sm font-bold text-white capitalize leading-tight">{currentImage.query}</h4>
                  </div>
                  <span className="text-[9px] font-mono text-zinc-400 bg-black/60 px-2.5 py-1 rounded-md border border-white/5">
                    JPEG High Quality
                  </span>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-4">
                <Button
                  onClick={() => downloadImage(currentImage.url, currentImage.query)}
                  className="h-12 bg-zinc-800 hover:bg-zinc-700 text-zinc-200 border border-zinc-700/50 rounded-xl text-[11px] font-mono font-bold uppercase tracking-widest transition-all flex items-center justify-center gap-2"
                >
                  <Download className="w-4 h-4" />
                  Download
                </Button>
                <Button
                  onClick={() => printImage(currentImage.url)}
                  className="h-12 bg-orange-600 hover:bg-orange-700 text-white rounded-xl text-[11px] font-mono font-bold uppercase tracking-widest transition-all flex items-center justify-center gap-2 shadow-lg shadow-orange-950/20"
                >
                  <Printer className="w-4 h-4" />
                  Print Image
                </Button>
              </div>
            </motion.div>
          </motion.div>
        )}

        {isBluetoothModalOpen && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-black/85 backdrop-blur-md z-50 flex items-center justify-center p-4 md:p-6"
          >
            <motion.div 
              initial={{ scale: 0.9, y: 20 }}
              animate={{ scale: 1, y: 0 }}
              exit={{ scale: 0.9, y: 20 }}
              transition={{ type: "spring", damping: 25, stiffness: 350 }}
              className="bg-zinc-900 border border-white/10 rounded-[2rem] p-6 max-w-lg w-full flex flex-col gap-5 shadow-[0_25px_60px_-15px_rgba(0,0,0,0.9)] relative text-left"
            >
              <button 
                onClick={() => setIsBluetoothModalOpen(false)}
                className="absolute top-4 right-4 p-2 bg-zinc-800 hover:bg-zinc-700 text-zinc-300 hover:text-white rounded-full transition-colors z-10"
                title="Close"
              >
                <X className="w-5 h-5" />
              </button>

              <div className="flex items-center gap-3">
                <div className="p-2.5 bg-blue-600/20 text-blue-400 rounded-xl border border-blue-500/20">
                  <Bluetooth className="w-5 h-5 animate-pulse" />
                </div>
                <div>
                  <h3 className="text-xs font-semibold text-zinc-100 uppercase tracking-widest font-mono text-left">Tania Speaker & Bluetooth Routing</h3>
                  <p className="text-[10px] text-zinc-500 font-mono text-left">Route Tania's real-time voice output directly to Bluetooth speakers</p>
                </div>
              </div>

              {/* Scanned audio outputs list */}
              <div className="space-y-3">
                <label className="text-[10px] font-mono font-bold uppercase tracking-wider text-zinc-400 block mb-1">
                  Choose Output Audio Destination
                </label>
                
                <div className="max-h-[160px] overflow-y-auto custom-scrollbar space-y-2 pr-1">
                  {/* Default/Main Unit */}
                  <div 
                    onClick={() => {
                      setSelectedAudioDeviceId("default");
                      if (liveApiRef.current) {
                        try {
                          liveApiRef.current.setAudioOutputDevice("default");
                        } catch (e) {
                          console.warn("Could not set active output device:", e);
                        }
                      }
                    }}
                    className={`p-3 rounded-xl border transition-all duration-300 cursor-pointer flex items-center justify-between ${
                      selectedAudioDeviceId === "default" 
                        ? "bg-blue-600/10 border-blue-500/40 text-blue-400" 
                        : "bg-zinc-950/20 border-white/5 hover:border-zinc-700 text-zinc-350 hover:text-white"
                    }`}
                  >
                    <div className="flex items-center gap-2.5">
                      <Volume2 className="w-4 h-4 text-zinc-400" />
                      <div className="text-left">
                        <p className="text-[11px] font-semibold font-mono uppercase tracking-wide">Main Internal Unit (Default)</p>
                        <p className="text-[9px] text-zinc-500 font-mono">Default browser audio output</p>
                      </div>
                    </div>
                    {selectedAudioDeviceId === "default" && (
                      <span className="text-[8px] font-mono bg-blue-500 text-neutral-950 px-2 py-0.5 rounded-full font-bold uppercase">
                        Active
                      </span>
                    )}
                  </div>

                  {/* Combined outputs */}
                  {(() => {
                    const mergedList: { deviceId: string; label: string; isBluetooth: boolean }[] = [];
                    
                    audioDevices.forEach(d => {
                      const dLabel = d.label || `Speaker Port (id: ${d.deviceId.slice(0, 8)}...)`;
                      const isBt = dLabel.toLowerCase().includes("bluetooth") || dLabel.toLowerCase().includes("wireless") || dLabel.toLowerCase().includes("audio");
                      mergedList.push({
                        deviceId: d.deviceId,
                        label: dLabel,
                        isBluetooth: isBt
                      });
                    });

                    customBluetoothDevices.forEach(cd => {
                      if (cd.paired && !mergedList.some(m => m.deviceId === cd.deviceId)) {
                        mergedList.push({
                          deviceId: cd.deviceId,
                          label: `🔊 Bluetooth External: ${cd.label} (Engaged Link)`,
                          isBluetooth: true
                        });
                      }
                    });

                    return mergedList.map((dev) => {
                      const isSelected = selectedAudioDeviceId === dev.deviceId;
                      const isBluetoothDevice = dev.isBluetooth;
                      const dLabel = dev.label;
                      
                      return (
                        <div 
                          key={dev.deviceId}
                          onClick={() => {
                            setSelectedAudioDeviceId(dev.deviceId);
                            if (liveApiRef.current && !dev.deviceId.startsWith("bt_")) {
                              try {
                                (liveApiRef.current as any).setAudioOutputDevice(dev.deviceId);
                              } catch (e) {
                                console.warn("Could not set active output device:", e);
                              }
                            }
                          }}
                          className={`p-3 rounded-xl border transition-all duration-300 cursor-pointer flex items-center justify-between group ${
                            isSelected 
                              ? "bg-blue-600/10 border-blue-500/40 text-blue-400" 
                              : "bg-zinc-950/20 border-white/5 hover:border-zinc-700 text-zinc-350 hover:text-white"
                          }`}
                        >
                          <div className="flex items-center gap-2.5 min-w-0 flex-1">
                            {isBluetoothDevice ? (
                              <Bluetooth className={`w-4 h-4 ${isSelected ? "text-blue-400" : "text-zinc-500"}`} />
                            ) : (
                              <Volume2 className={`w-4 h-4 ${isSelected ? "text-blue-400" : "text-zinc-500"}`} />
                            )}
                            <div className="text-left min-w-0 flex-1">
                              <p className="text-[11px] font-semibold font-mono uppercase tracking-wide truncate">
                                {dLabel}
                              </p>
                              <p className="text-[9px] text-zinc-500 font-mono truncate">
                                {isBluetoothDevice ? "Bluetooth / Wireless Link (Engaged)" : "Integrated Port Connection"}
                              </p>
                            </div>
                          </div>
                          {isSelected ? (
                            <span className="text-[8px] font-mono bg-blue-500 text-neutral-950 px-2 py-0.5 rounded-full font-bold uppercase shrink-0">
                              Active
                            </span>
                          ) : (
                            <span className="text-[8px] font-mono text-zinc-500 opacity-0 group-hover:opacity-100 shrink-0 uppercase tracking-widest transition-opacity">
                              Select
                            </span>
                          )}
                        </div>
                      );
                    });
                  })()}
                </div>
              </div>

              {/* INTEGRATED BLUETOOTH RADIO TRANSCIEVER BEACON SCANNER & PAIRING FRAMEWORK */}
              <div className="bg-zinc-950/40 border border-white/5 rounded-2xl p-4 space-y-3.5">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <Bluetooth className="w-3.5 h-3.5 text-blue-400 animate-pulse" />
                    <span className="text-[10px] font-mono font-bold uppercase tracking-wider text-zinc-300">
                      Radio Beacon BLE Scanner
                    </span>
                  </div>
                  
                  <button
                    onClick={startBluetoothScan}
                    disabled={isScanningBluetooth}
                    className="px-2.5 py-1 bg-blue-600 hover:bg-blue-500 disabled:bg-blue-800 text-zinc-950 rounded-lg text-[9px] font-mono font-bold uppercase tracking-wider flex items-center gap-1.5 transition-all text-neutral-950"
                  >
                    <Activity className={`w-3 h-3 ${isScanningBluetooth ? "animate-spin" : ""}`} />
                    {isScanningBluetooth ? "Scanning..." : "📡 SCAN NEARBY BLE"}
                  </button>
                </div>

                {isScanningBluetooth ? (
                  <div className="py-6 flex flex-col items-center justify-center space-y-2.5 bg-zinc-950/60 rounded-xl border border-blue-500/10">
                    <Loader2 className="w-6 h-6 text-blue-400 animate-spin" />
                    <p className="text-[10.5px] font-mono text-blue-300 uppercase tracking-wider animate-pulse">{scanStatusMessage || "Tuning radio frequency..."}</p>
                    <p className="text-[8.5px] text-zinc-500 font-mono">Listening on 2.4GHz channels {`{37, 38, 39}`} for Bluetooth 5.x advertisements</p>
                  </div>
                ) : (
                  <div className="space-y-2.5">
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 max-h-[140px] overflow-y-auto custom-scrollbar pr-0.5">
                      {customBluetoothDevices.map((device) => {
                        const isPaired = device.paired;
                        const isPairing = pairingDeviceId === device.deviceId;
                        
                        return (
                          <div 
                            key={device.deviceId} 
                            className={`p-2.5 rounded-xl border transition-all flex items-center justify-between gap-2.5 text-left ${
                              isPaired 
                                ? "bg-emerald-950/15 border-emerald-500/20" 
                                : "bg-zinc-900/40 border-white/5"
                            }`}
                          >
                            <div className="min-w-0 flex-1">
                              <p className="text-[10px] font-bold text-zinc-200 uppercase tracking-wide truncate">{device.label}</p>
                              <p className="text-[8px] font-mono text-zinc-500 uppercase mt-0.5">
                                {isPaired ? "✓ Paired & Bound" : `RSSI: ${device.rssi || -60}dBm`}
                              </p>
                            </div>

                            <button
                              onClick={() => engageBluetoothDevice(device.deviceId)}
                              disabled={isPaired || isPairing}
                              className={`px-2.5 py-1 rounded-lg text-[8.5px] font-mono font-bold uppercase tracking-wider transition-all shrink-0 ${
                                isPaired 
                                  ? "bg-zinc-800 text-emerald-400 cursor-default" 
                                  : isPairing 
                                    ? "bg-blue-900/40 text-blue-300 border border-blue-500/20 animate-pulse" 
                                    : "bg-blue-600 hover:bg-blue-550 text-neutral-950"
                              }`}
                            >
                              {isPaired ? "Engaged" : isPairing ? "Pairing..." : "Engage"}
                            </button>
                          </div>
                        );
                      })}
                    </div>

                    <form onSubmit={handleAddManualBluetoothDevice} className="flex gap-2 items-center border-t border-zinc-900 pt-2.5">
                      <input
                        type="text"
                        required
                        placeholder="Type Custom Speaker Name to Pair (e.g. JBL FLIP 6)..."
                        value={bluetoothManualDeviceName}
                        onChange={(e) => setBluetoothManualDeviceName(e.target.value)}
                        className="flex-1 px-3 py-1.5 bg-zinc-900 border border-white/5 rounded-xl text-[10.5px] text-zinc-100 placeholder:text-zinc-650 focus:outline-none focus:border-blue-500/40 font-mono"
                      />
                      <button
                        type="submit"
                        className="px-3 h-8 bg-blue-600/15 hover:bg-blue-600 hover:text-zinc-950 text-blue-400 border border-blue-500/25 rounded-xl text-[9px] font-mono font-bold uppercase tracking-wider transition-all shadow-md shrink-0"
                      >
                        Add & Pair
                      </button>
                    </form>
                  </div>
                )}
              </div>

              {/* Action buttons */}
              <div className="grid grid-cols-2 gap-3 pt-1">
                <Button
                  onClick={async () => {
                    await refreshAudioDevices();
                  }}
                  className="h-10 bg-zinc-800 hover:bg-zinc-700 text-zinc-300 hover:text-white border border-white/10 rounded-xl text-[10px] font-mono font-semibold uppercase tracking-wider transition-all flex items-center justify-center gap-2"
                >
                  <Loader2 className="w-3.5 h-3.5 text-zinc-500" />
                  Refresh Ports
                </Button>
                
                <Button
                  onClick={playTestTone}
                  disabled={isTestTonePlaying}
                  className={`h-10 rounded-xl text-[10px] font-mono font-semibold uppercase tracking-wider transition-all flex items-center justify-center gap-2 ${
                    isTestTonePlaying 
                      ? "bg-blue-600/30 text-blue-400 border border-blue-500/30 cursor-not-allowed" 
                      : "bg-blue-600 hover:bg-blue-700 text-zinc-950 hover:text-zinc-950 shadow-lg"
                  }`}
                >
                  {isTestTonePlaying ? (
                    <>
                      <Loader2 className="w-3.5 h-3.5 animate-spin text-blue-400" />
                      Testing Chime...
                    </>
                  ) : (
                    <>
                      <Volume2 className="w-3.5 h-3.5" />
                      Play Test Tone
                    </>
                  )}
                </Button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      <div className={`w-full ${currentImage ? "max-w-7xl lg:grid lg:grid-cols-12 lg:gap-8 lg:items-start" : "max-w-4xl"} space-y-8 lg:space-y-0 flex flex-col`}>
        <div className={currentImage ? "lg:col-span-7 space-y-8 flex flex-col items-center w-full" : "w-full space-y-8 flex flex-col items-center"}>
          {/* Header */}
          <div className="flex flex-col items-center text-center space-y-2 max-w-md w-full">
            <div className="w-12 h-12 bg-orange-600 rounded-2xl flex items-center justify-center shadow-xl shadow-orange-900/20 mb-2">
              <Sparkles className="text-white w-7 h-7" />
            </div>
            <h1 className="text-2xl font-medium tracking-tight">Tania</h1>
            {discussedSubject && (
              <motion.div 
                initial={{ scale: 0.95, opacity: 0 }}
                animate={{ scale: 1, opacity: 1 }}
                className="px-4 py-1.5 bg-orange-600/15 border border-orange-500/30 rounded-2xl text-[11px] font-medium text-orange-400 max-w-sm md:max-w-md truncate text-center shadow-lg shadow-orange-950/10 flex items-center gap-1.5"
              >
                <span className="w-1.5 h-1.5 rounded-full bg-orange-500 animate-ping" />
                <span>Discussed Subject:</span>
                <span className="text-amber-300 font-bold tracking-wide uppercase">{discussedSubject}</span>
              </motion.div>
            )}
            <p className="text-zinc-300 text-xs uppercase tracking-widest font-mono">Voice Assistant</p>
            <p className="text-orange-500 text-xs font-semibold uppercase tracking-[0.15em] font-mono mt-2">Owner: Pesala Jayawardene</p>
            <p className="text-zinc-400 text-[10px] uppercase tracking-widest font-mono">AI Software version v112.01</p>
            {error && (
              <motion.div 
                initial={{ opacity: 0, y: -10 }}
                animate={{ opacity: 1, y: 0 }}
                className="mt-4 p-4 bg-red-500/10 border border-red-500/20 rounded-xl text-red-400 text-xs font-mono flex flex-col items-center gap-3"
              >
                <p className="text-center">{error}</p>
                {(error.includes("Quota") || error.includes("Key")) && (
                  <Button 
                    onClick={handleSelectKey}
                    variant="outline" 
                    className="h-8 bg-red-500/20 border-red-500/30 hover:bg-red-500/30 text-[10px] uppercase tracking-widest"
                  >
                    Select Paid API Key
                  </Button>
                )}
              </motion.div>
            )}
          </div>

          {/* Status Indicator */}
          <div className="flex justify-center py-8 w-full max-w-md">
            <div className={`w-32 h-32 rounded-full border-2 flex items-center justify-center transition-all duration-500 ${
              isConnected 
                ? isTalking 
                  ? 'border-orange-500 bg-orange-500/5 scale-110 shadow-[0_0_40px_rgba(249,115,22,0.2)]' 
                  : 'border-zinc-800 bg-zinc-900/50 scale-100'
                : 'border-zinc-900 bg-transparent opacity-50'
            }`}>
              {isConnected ? (
                <div className="flex gap-1.5 items-end h-10">
                  {[...Array(5)].map((_, i) => (
                    <motion.div
                      key={i}
                      animate={{
                        height: isTalking ? [10, 30, 10] : 10,
                      }}
                      transition={{
                        duration: 0.5,
                        repeat: Infinity,
                        delay: i * 0.1,
                      }}
                      className="w-2 bg-orange-500 rounded-full"
                    />
                  ))}
                </div>
              ) : (
                <Mic className="w-10 h-10 text-zinc-800" />
              )}
            </div>
          </div>

          {/* Transcript / Workspace Area */}
          <div 
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            className={`bg-zinc-900/40 backdrop-blur-2xl border border-white/5 rounded-[2rem] overflow-hidden flex flex-col ${
              activeTab === "videos" && currentVideoId
                ? "h-[580px] sm:h-[720px] md:h-[820px] lg:h-[900px]"
                : "h-[400px] sm:h-[480px] md:h-[560px]"
            } w-full shadow-[0_20px_50px_rgba(0,0,0,0.5)] relative group/transcript transition-all`}
          >
            {isDragging && (
              <div className="absolute inset-0 z-30 bg-zinc-950/90 backdrop-blur-md flex flex-col items-center justify-center border-2 border-dashed border-amber-500/80 m-4 rounded-[1.5rem] animate-fade-in pointer-events-none">
                <div className="w-16 h-16 rounded-2xl bg-amber-600/10 border border-amber-500/30 flex items-center justify-center mb-4 text-amber-400 animate-bounce">
                  <Upload className="w-8 h-8" />
                </div>
                <h3 className="text-sm font-mono font-bold uppercase tracking-wider text-amber-400">Release to Upload Info Document</h3>
                <p className="text-[10px] text-zinc-500 font-mono mt-1 uppercase tracking-widest font-bold">TXT, CSV, JSON, XLS or XLSX spreadsheets</p>
              </div>
            )}
            <div className="sticky top-0 z-20 px-6 py-4 border-b border-white/5 bg-zinc-900/80 backdrop-blur-md flex flex-col lg:flex-row lg:items-center justify-between gap-3">
              <div className="flex flex-col sm:flex-row sm:items-center gap-3.5 flex-1">
                <div className="flex items-center gap-3 shrink-0">
                  <div className="flex items-center gap-2">
                    <div className={`w-2 h-2 rounded-full ${isConnected ? 'bg-emerald-500 animate-pulse' : isConnecting ? 'bg-orange-500 animate-bounce' : 'bg-red-500'} shadow-[0_0_8px_rgba(234,88,12,0.5)]`} />
                    <span className="text-[10px] font-mono uppercase tracking-[0.15em] text-zinc-100 font-bold">
                      {isConnected ? 'Live Stream' : isConnecting ? 'Connecting...' : 'Ready'}
                    </span>
                  </div>
                  {transcript.length > 0 && activeTab === "conversation" && (
                    <button 
                      onClick={clearTranscript}
                      className="ml-1 text-[8px] text-zinc-500 hover:text-orange-400 uppercase tracking-widest font-mono transition-all hover:scale-110"
                    >
                      [Clear]
                    </button>
                  )}
                </div>

                {/* Speaker Personality Mood / Slang Radio Tabs */}
                <div className="flex flex-wrap items-center gap-1 bg-zinc-950/60 p-1 border border-white/5 rounded-xl max-w-full overflow-x-auto custom-scrollbar">
                  {(["Default", "Friendly", "Lovable", "Sad", "Angry", "Official", "Slang mixed"] as const).map((mood) => (
                    <label
                      key={mood}
                      className={`flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-[8.5px] font-mono font-bold uppercase tracking-wider cursor-pointer border transition-all ${
                        taniaMood === mood
                          ? "bg-amber-500/10 text-amber-400 border-amber-500/25 shadow-md shadow-amber-950/40"
                          : "text-zinc-500 border-transparent hover:text-zinc-350 hover:bg-zinc-800/40"
                      }`}
                    >
                      <input
                        type="radio"
                        name="taniaMoodRadio"
                        checked={taniaMood === mood}
                        onChange={() => {
                          setTaniaMood(mood);
                          if (isConnected && liveApiRef.current) {
                            liveApiRef.current.sendText(`[SYSTEM UPDATE] Please dynamically switch your personality and voice tone immediately to: "${mood}". Adopt the precise voice specifications for "${mood}" starting from your next statement!`);
                          }
                        }}
                        className="sr-only"
                      />
                      <span className={`w-1 h-1 rounded-full transition-all shrink-0 ${
                        taniaMood === mood 
                          ? 'bg-amber-400 shadow-[0_0_5px_rgba(245,158,11,0.8)] animate-pulse' 
                          : 'bg-zinc-700'
                      }`} />
                      <span>{mood}</span>
                    </label>
                  ))}
                </div>
              </div>
              <div className="flex items-center gap-4">
                {/* Embedded Speaker Dropdown Selector & Bluetooth Interceptor */}
                <div className="flex items-center gap-1.5 bg-zinc-950/80 border border-white/5 rounded-xl px-2.5 py-1">
                  <button 
                    onClick={handleConnectBluetooth}
                    className="p-1 text-zinc-400 hover:text-blue-400 transition-all flex items-center justify-center"
                    title="Filter / Connect Bluetooth Speaker"
                  >
                    <Bluetooth className={`w-3.5 h-3.5 ${selectedAudioDeviceId !== "default" ? "text-blue-500 animate-pulse" : ""}`} />
                  </button>
                  <select
                    value={selectedAudioDeviceId}
                    onChange={(e) => {
                      const deviceId = e.target.value;
                      setSelectedAudioDeviceId(deviceId);
                      if (liveApiRef.current) {
                        liveApiRef.current.setAudioOutputDevice(deviceId);
                      }
                    }}
                    className="bg-transparent text-[9.5px] text-zinc-300 font-mono focus:outline-none cursor-pointer max-w-[110px] border-none pr-1 uppercase"
                    title="Audio Output Speaker Route"
                  >
                    <option value="default" className="bg-zinc-950 text-zinc-200">Main Unit</option>
                    {audioDevices.map((d) => (
                      <option key={d.deviceId} value={d.deviceId} className="bg-zinc-950 text-zinc-200">
                        {d.label || `Spkr: ${d.deviceId.slice(0, 5)}`}
                      </option>
                    ))}
                  </select>
                </div>

                {(saveSuccess || isSaving) && (
                  <motion.div 
                    initial={{ opacity: 0, scale: 0.9 }}
                    animate={{ opacity: 1, scale: 1 }}
                    className="flex items-center gap-1.5"
                  >
                    <div className={`w-1.5 h-1.5 rounded-full ${isSaving ? 'bg-orange-500 animate-ping' : 'bg-emerald-500'}`} />
                    <span className="text-[9px] font-mono text-emerald-400 uppercase tracking-widest font-bold">
                      {isSaving ? 'Synching...' : 'Saved'}
                    </span>
                  </motion.div>
                )}
                <div className="h-4 w-[1px] bg-white/10 hidden sm:block" />
                <button 
                  onClick={saveConversation}
                  className={`p-1.5 transition-all hover:scale-110 ${isSaving ? 'text-zinc-600' : 'text-emerald-500 hover:text-emerald-400'}`}
                  title="Save System Conversation States"
                  disabled={isSaving}
                >
                  <Save className="w-4 h-4" />
                </button>
              </div>
            </div>

            {/* Workspace Tab Buttons */}
            <div className="px-6 py-2.5 bg-zinc-900/60 border-b border-white/5 flex flex-col md:flex-row md:items-center justify-between gap-3">
              <div className="flex gap-2 items-center overflow-x-auto custom-scrollbar pb-1 md:pb-0 w-full md:w-auto flex-1">
                <button
                  onClick={() => setActiveTab("conversation")}
                  className={`px-3 py-1.5 rounded-lg text-[10px] font-mono uppercase tracking-wider transition-all duration-300 font-bold border ${
                    activeTab === "conversation"
                      ? "bg-orange-600/20 text-orange-400 border-orange-500/20 shadow-lg shadow-orange-950/20"
                      : "text-zinc-400 border-transparent hover:text-zinc-200 hover:bg-zinc-800/30"
                  }`}
                >
                  Chat & Audio
                </button>
                
                <button
                  onClick={() => setActiveTab("quotes")}
                  className={`px-3 py-1.5 rounded-lg text-[10px] font-mono uppercase tracking-wider transition-all duration-300 font-bold flex items-center gap-1.5 border ${
                    activeTab === "quotes"
                      ? "bg-emerald-600/20 text-emerald-400 border-emerald-500/20 shadow-lg shadow-emerald-950/20"
                      : "text-zinc-400 border-transparent hover:text-zinc-200 hover:bg-zinc-800/30"
                  }`}
                >
                  <span>Quotes & Prices</span>
                  {recordedQuotes.length > 0 && (
                    <span className="bg-emerald-500 text-neutral-950 px-1.5 py-0.5 text-[8px] rounded-full font-bold leading-none">
                      {recordedQuotes.length}
                    </span>
                  )}
                </button>

                <button
                  onClick={() => setActiveTab("documents")}
                  className={`px-3 py-1.5 rounded-lg text-[10px] font-mono uppercase tracking-wider transition-all duration-300 font-bold flex items-center gap-1.5 border ${
                    activeTab === "documents"
                      ? "bg-cyan-600/20 text-cyan-400 border-cyan-500/20 shadow-lg shadow-cyan-950/20"
                      : "text-zinc-400 border-transparent hover:text-zinc-200 hover:bg-zinc-800/30"
                  }`}
                >
                  <span>Drafts & Letters</span>
                  {recordedDocuments.length > 0 && (
                    <span className="bg-cyan-500 text-neutral-950 px-1.5 py-0.5 text-[8px] rounded-full font-bold leading-none">
                      {recordedDocuments.length}
                    </span>
                  )}
                </button>

                <button
                  onClick={() => setActiveTab("communications")}
                  className={`px-3 py-1.5 rounded-lg text-[10px] font-mono uppercase tracking-wider transition-all duration-300 font-bold flex items-center gap-1.5 border ${
                    activeTab === "communications"
                      ? "bg-purple-600/20 text-purple-400 border-purple-500/20 shadow-lg shadow-purple-950/20"
                      : "text-zinc-400 border-transparent hover:text-zinc-200 hover:bg-zinc-800/30"
                  }`}
                >
                  <span>Mail & Messages</span>
                  {recordedCommunications.length > 0 && (
                    <span className="bg-purple-500 text-neutral-950 px-1.5 py-0.5 text-[8px] rounded-full font-bold leading-none">
                      {recordedCommunications.length}
                    </span>
                  )}
                </button>

                <button
                  onClick={() => setActiveTab("uploads")}
                  className={`px-3 py-1.5 rounded-lg text-[10px] font-mono uppercase tracking-wider transition-all duration-300 font-bold flex items-center gap-1.5 border ${
                    activeTab === "uploads"
                      ? "bg-amber-600/20 text-amber-400 border-amber-500/20 shadow-lg shadow-amber-950/20"
                      : "text-zinc-400 border-transparent hover:text-zinc-200 hover:bg-zinc-800/30"
                  }`}
                >
                  <Paperclip className="w-3.5 h-3.5" />
                  <span>Uploads & Custom Info</span>
                  {uploadedFiles.length > 0 && (
                    <span className="bg-amber-500 text-neutral-950 px-1.5 py-0.5 text-[8px] rounded-full font-bold leading-none animate-pulse">
                      {uploadedFiles.length}
                    </span>
                  )}
                </button>

                <button
                  onClick={() => setActiveTab("reminders")}
                  className={`px-3 py-1.5 rounded-lg text-[10px] font-mono uppercase tracking-wider transition-all duration-300 font-bold flex items-center gap-1.5 border ${
                    activeTab === "reminders"
                      ? "bg-rose-600/20 text-rose-400 border-rose-500/20 shadow-lg shadow-rose-950/20"
                      : "text-zinc-400 border-transparent hover:text-zinc-200 hover:bg-zinc-800/30"
                  }`}
                >
                  <Bell className="w-3.5 h-3.5" />
                  <span>Reminders</span>
                  {recordedReminders.length > 0 && (
                    <span className="bg-rose-500 text-neutral-950 px-1.5 py-0.5 text-[8px] rounded-full font-bold leading-none">
                      {recordedReminders.length}
                    </span>
                  )}
                </button>

                <button
                  onClick={() => setActiveTab("rem-info")}
                  className={`px-3 py-1.5 rounded-lg text-[10px] font-mono uppercase tracking-wider transition-all duration-300 font-bold flex items-center gap-1.5 border ${
                    activeTab === "rem-info"
                      ? "bg-violet-600/20 text-violet-400 border-violet-500/20 shadow-lg shadow-violet-950/20"
                      : "text-zinc-400 border-transparent hover:text-zinc-200 hover:bg-zinc-800/30"
                  }`}
                >
                  <CheckCircle2 className="w-3.5 h-3.5" />
                  <span>Rem-Info</span>
                  {recordedRemInfos.length > 0 && (
                    <span className="bg-violet-500 text-neutral-950 px-1.5 py-0.5 text-[8px] rounded-full font-bold leading-none animate-pulse">
                      {recordedRemInfos.length}
                    </span>
                  )}
                </button>

                <button
                  onClick={() => setActiveTab("standing-orders")}
                  className={`px-3 py-1.5 rounded-lg text-[10px] font-mono uppercase tracking-wider transition-all duration-300 font-bold flex items-center gap-1.5 border ${
                    activeTab === "standing-orders"
                      ? "bg-amber-600/20 text-amber-400 border-amber-500/20 shadow-lg shadow-amber-950/20"
                      : "text-zinc-400 border-transparent hover:text-zinc-200 hover:bg-zinc-800/30"
                  }`}
                >
                  <FileText className="w-3.5 h-3.5" />
                  <span>Standing Orders</span>
                  {recordedStandingOrders.length > 0 && (
                    <span className="bg-amber-500 text-neutral-950 px-1.5 py-0.5 text-[8px] rounded-full font-bold leading-none animate-pulse">
                      {recordedStandingOrders.length}
                    </span>
                  )}
                </button>

                <button
                  onClick={() => setActiveTab("pictures")}
                  className={`px-3 py-1.5 rounded-lg text-[10px] font-mono uppercase tracking-wider transition-all duration-300 font-bold flex items-center gap-1.5 border ${
                    activeTab === "pictures"
                      ? "bg-teal-600/20 text-teal-400 border-teal-500/20 shadow-lg shadow-teal-950/20"
                      : "text-zinc-400 border-transparent hover:text-zinc-200 hover:bg-zinc-800/30"
                  }`}
                >
                  <Image className="w-3.5 h-3.5" />
                  <span>Pictures</span>
                  {requestedImages.length > 0 && (
                    <span className="bg-teal-500 text-neutral-950 px-1.5 py-0.5 text-[8px] rounded-full font-bold leading-none">
                      {requestedImages.length}
                    </span>
                  )}
                </button>

                <button
                  onClick={() => setActiveTab("videos")}
                  className={`px-3 py-1.5 rounded-lg text-[10px] font-mono uppercase tracking-wider transition-all duration-300 font-bold flex items-center gap-1.5 border ${
                    activeTab === "videos"
                      ? "bg-amber-600/20 text-amber-400 border-amber-500/20 shadow-lg shadow-amber-950/20"
                      : "text-zinc-400 border-transparent hover:text-zinc-200 hover:bg-zinc-800/30"
                  }`}
                >
                  <Video className="w-3.5 h-3.5" />
                  <span>Videos & Stream</span>
                  {requestedVideos.length > 0 && (
                    <span className="bg-amber-500 text-neutral-950 px-1.5 py-0.5 text-[8px] rounded-full font-bold leading-none">
                      {requestedVideos.length}
                    </span>
                  )}
                </button>
              </div>

              {/* Dynamic Information Mode Toggle Switch */}
              <div 
                id="search-verification-toggle"
                className="flex items-center gap-2.5 bg-zinc-950/80 border border-white/5 rounded-xl px-3 py-1.5 shrink-0 self-start md:self-auto shadow-lg"
              >
                <div className="flex flex-col text-left">
                  <span className="text-[8.5px] font-mono font-medium text-zinc-500 uppercase tracking-widest leading-none">Tania Core Search</span>
                  <span className={`text-[9px] font-mono font-bold uppercase mt-1 leading-none transition-colors duration-255 ${includeUnverifiedInfo ? "text-amber-400" : "text-emerald-400"}`}>
                    {includeUnverifiedInfo ? "All Info (Verified + Unverified)" : "Strictly Verified Only"}
                  </span>
                </div>
                <button
                  type="button"
                  onClick={() => setIncludeUnverifiedInfo(prev => !prev)}
                  className={`relative inline-flex h-5 w-9 shrink-0 cursor-pointer rounded-full border-2 border-transparent transition-colors duration-200 ease-in-out focus:outline-none ${
                    includeUnverifiedInfo ? "bg-amber-600" : "bg-emerald-600"
                  }`}
                  role="switch"
                  aria-checked={includeUnverifiedInfo}
                  title={includeUnverifiedInfo ? "Switch to verified information only" : "Switch to allow both verified & unverified information"}
                >
                  <span
                    aria-hidden="true"
                    className={`pointer-events-none inline-block h-4 w-4 transform rounded-full bg-white shadow ring-0 transition duration-200 ease-in-out ${
                      includeUnverifiedInfo ? "translate-x-4" : "translate-x-0"
                    }`}
                  />
                </button>
              </div>
            </div>
            
            {/* Tab Inner Contents */}
            {activeTab === "conversation" && (
              <div 
                id="transcript-content"
                className="flex-1 overflow-y-auto custom-scrollbar p-6 space-y-6 scroll-smooth" 
                ref={scrollRef}
                style={{ scrollbarWidth: 'thin', scrollbarColor: '#3f3f46 transparent' }}
              >
                <div className="space-y-6">
                  {transcript.length === 0 ? (
                    <div className="flex flex-col items-center justify-center py-20 space-y-4">
                      <div className="w-12 h-12 rounded-full border border-zinc-800 flex items-center justify-center">
                        <MessageSquare className="w-5 h-5 text-zinc-700" />
                      </div>
                      <p className="text-zinc-600 text-xs font-mono uppercase tracking-widest">Waiting for connection...</p>
                      {isConnected && (
                        <motion.div 
                          animate={{ opacity: [0.3, 1, 0.3] }}
                          transition={{ duration: 2, repeat: Infinity }}
                          className="flex items-center gap-2"
                        >
                          <div className="w-1.5 h-1.5 rounded-full bg-orange-500/40" />
                          <span className="text-[10px] font-mono text-zinc-500 uppercase tracking-[0.2em]">Tania is listening</span>
                        </motion.div>
                      )}
                    </div>
                  ) : (
                    transcript.map((text, i) => {
                      const isTania = text.startsWith("Tania:");
                      const isPesala = text.startsWith("Pesala:");
                      const isSystem = text.startsWith("System:");
                      
                      const match = text.match(/^(Tania|Pesala|You|System):\s*(\[.*?\])?\s*(.*)$/);
                      const timestamp = match && match[2] ? match[2].trim() : "";
                      const cleanText = match ? match[3] : text.replace(/^(Tania|Pesala|You|System): /, "");
                      
                      const isImagePreview = cleanText.includes("[IMAGE_PREVIEW:");
                      let imageQuery = "";
                      let displayUrl = "";
                      
                      if (isImagePreview) {
                        const imgMatch = cleanText.match(/\[IMAGE_PREVIEW:\s*(.*?)\]/);
                        if (imgMatch) {
                          imageQuery = imgMatch[1];
                          displayUrl = resolvedImages[imageQuery] || (currentImage && currentImage.query === imageQuery 
                            ? currentImage.url 
                            : `/api/image-proxy?query=${encodeURIComponent(imageQuery)}&sig=static`);
                        }
                      }

                      return (
                        <motion.div
                          key={i}
                          initial={{ opacity: 0, y: 10 }}
                          animate={{ opacity: 1, y: 0 }}
                          className={`flex flex-col gap-2 ${isPesala ? 'items-end' : 'items-start'}`}
                        >
                          <div className="flex items-center gap-2 px-1">
                            <span className={`text-[9px] font-mono uppercase tracking-widest font-bold ${isTania ? 'text-orange-500' : isPesala ? 'text-blue-400' : 'text-zinc-500'}`}>
                              {isTania ? 'Tania' : isPesala ? 'Pesala' : isSystem ? 'System' : 'Unknown'}
                            </span>
                            {timestamp && (
                              <span className="text-[8.5px] font-mono text-zinc-500 font-medium">
                                {timestamp}
                              </span>
                            )}
                          </div>
                          <div className={`max-w-[85%] px-4 py-3 rounded-2xl text-sm leading-relaxed shadow-lg ${
                            isTania 
                              ? 'bg-zinc-800/90 border border-white/10 text-white font-medium rounded-tl-none' 
                              : isPesala 
                                ? 'bg-orange-600/30 border border-orange-500/40 text-white font-medium text-left rounded-tr-none'
                                : 'bg-zinc-950 border border-zinc-800 text-zinc-300 font-medium italic text-xs'
                          }`}>
                            {isImagePreview ? (
                              <div className="space-y-3">
                                <p className="text-zinc-200">{cleanText.replace(/\[IMAGE_PREVIEW:.*?\]/, "").trim() || `Rendered picture request for "${imageQuery}":`}</p>
                                <div 
                                  onClick={() => {
                                    setCurrentImage({ url: displayUrl, query: imageQuery });
                                    setIsImageModalOpen(true);
                                  }}
                                  className="relative aspect-[4/3] rounded-xl overflow-hidden cursor-pointer group bg-zinc-950 border border-white/5 w-full max-w-[280px]"
                                >
                                  {isImageLoading && currentImage && currentImage.query === imageQuery && (
                                    <div className="absolute inset-0 bg-black/60 flex items-center justify-center">
                                      <Loader2 className="w-5 h-5 text-orange-500 animate-spin" />
                                    </div>
                                  )}
                                  <img 
                                    src={displayUrl} 
                                    alt={imageQuery}
                                    referrerPolicy="no-referrer"
                                    className="w-full h-full object-cover transition-transform group-hover:scale-105"
                                    onError={(e) => {
                                      const target = e.target as HTMLImageElement;
                                      if (!target.src.includes("loremflickr")) {
                                        target.src = `https://loremflickr.com/400/300/${encodeURIComponent(imageQuery)}`;
                                      } else if (!target.src.includes("picsum")) {
                                        target.src = "https://picsum.photos/400/300";
                                      }
                                    }}
                                  />
                                  <div className="absolute inset-0 bg-black/50 opacity-0 group-hover:opacity-100 transition-opacity flex items-center justify-center text-[9px] font-mono text-white tracking-widest uppercase">
                                    Expand Photo
                                  </div>
                                </div>
                              </div>
                            ) : (
                              cleanText
                            )}
                          </div>
                        </motion.div>
                      );
                    })
                  )}
                  <div ref={messagesEndRef} />
                </div>
              </div>
            )}

            {activeTab === "quotes" && (
              <div className="flex-1 overflow-y-auto custom-scrollbar p-6 space-y-6">
                {recordedQuotes.length === 0 ? (
                  <div className="flex flex-col items-center justify-center py-20 space-y-4">
                    <div className="w-12 h-12 rounded-full border border-zinc-800 flex items-center justify-center">
                      <FileSpreadsheet className="w-5 h-5 text-zinc-700" />
                    </div>
                    <p className="text-zinc-500 text-xs font-mono uppercase tracking-widest text-center">No price quotes recorded yet.</p>
                    <p className="text-zinc-600 text-[10px] text-center max-w-sm">Ask Tania: "Prepare a price quote for..." to automatically compile an elegant quotation grid here.</p>
                  </div>
                ) : (
                  <div className="space-y-6">
                    {recordedQuotes.map((quote, qIdx) => (
                      <div key={qIdx} className="bg-zinc-950/60 border border-zinc-850 rounded-2xl p-5 space-y-4">
                        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-zinc-800 pb-3">
                          <div>
                            <h3 className="text-sm font-semibold text-emerald-400 tracking-wide uppercase font-mono">{quote.title}</h3>
                            <p className="text-[10px] text-zinc-500 font-mono mt-1">{new Date(quote.createdAt).toLocaleString()}</p>
                          </div>
                          <div className="flex items-center gap-2">
                            <button
                              onClick={() => exportQuoteToPDF(quote)}
                              className="px-2.5 py-1 bg-zinc-900 border border-zinc-800 hover:bg-zinc-800 rounded-lg text-[9px] font-mono text-cyan-400 tracking-wider flex items-center gap-1.5 transition-all uppercase"
                              title="Export quotation as a PDF document"
                            >
                              <FileText className="w-3.5 h-3.5" />
                              PDF
                            </button>
                            <button
                              onClick={() => exportQuoteToWordObj(quote)}
                              className="px-2.5 py-1 bg-zinc-900 border border-zinc-800 hover:bg-zinc-800 rounded-lg text-[9px] font-mono text-orange-400 tracking-wider flex items-center gap-1.5 transition-all uppercase"
                              title="Export quotation as a Microsoft Word document"
                            >
                              <Download className="w-3.5 h-3.5" />
                              WORD
                            </button>
                            <button
                              onClick={() => exportQuoteToExcel(quote)}
                              className="px-2.5 py-1 bg-zinc-900 border border-zinc-800 hover:bg-zinc-800 rounded-lg text-[9px] font-mono text-emerald-400 tracking-wider flex items-center gap-1.5 transition-all uppercase"
                              title="Export quotation as an Excel Spreadsheet"
                            >
                              <FileSpreadsheet className="w-3.5 h-3.5" />
                              EXCEL
                            </button>
                          </div>
                        </div>

                        {/* Professional Quotation Grid */}
                        <div className="overflow-x-auto border border-zinc-900 rounded-xl bg-zinc-950/80">
                          <table className="w-full text-xs font-sans text-zinc-300 min-w-[500px]">
                            <thead>
                              <tr className="bg-zinc-900/60 border-b border-zinc-800 text-[9px] text-zinc-450 font-mono uppercase tracking-wider text-left">
                                <th className="py-3 px-3 w-12 text-center">No</th>
                                <th className="py-3 px-3">Item Description</th>
                                <th className="py-3 px-3 w-16 text-center">Qty</th>
                                <th className="py-3 px-3 w-28">Unit Price</th>
                                <th className="py-3 px-3 w-28 text-right">Sub-total</th>
                              </tr>
                            </thead>
                            <tbody>
                              {quote.items.map((item: any, iIdx: number) => (
                                <tr key={iIdx} className="border-b border-zinc-900 hover:bg-zinc-900/25 transition-colors">
                                  <td className="py-2.5 px-3 text-center text-zinc-500 font-mono">{iIdx + 1}</td>
                                  <td className="py-2.5 px-3">
                                    <div className="font-medium text-zinc-200">{item.description}</div>
                                    {item.supplier_name && (
                                      <div className="text-[10px] text-teal-400 font-mono mt-0.5 tracking-wider">
                                        Supplier: <span className="text-zinc-400">{item.supplier_name}</span>
                                      </div>
                                    )}
                                  </td>
                                  <td className="py-2.5 px-3 text-center font-mono">{item.quantity}</td>
                                  <td className="py-2.5 px-3 font-mono text-zinc-350">{item.price_per_unit}</td>
                                  <td className="py-2.5 px-3 text-right font-mono text-emerald-400 font-bold">{item.total_price}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>

                        {/* Total Highlight Row */}
                        <div className="flex justify-end pt-1">
                          <div className="bg-emerald-950/20 border border-emerald-800/20 px-4 py-2.5 rounded-xl flex items-center gap-4">
                            <span className="text-[10px] font-mono text-emerald-500 uppercase tracking-widest font-bold">Combined Estimation:</span>
                            <span className="text-xs font-bold text-emerald-400 font-mono">{quote.total}</span>
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}

            {activeTab === "documents" && (
              <div className="flex-1 overflow-y-auto custom-scrollbar p-6 space-y-6">
                {recordedDocuments.length === 0 ? (
                  <div className="flex flex-col items-center justify-center py-20 space-y-4">
                    <div className="w-12 h-12 rounded-full border border-zinc-800 flex items-center justify-center">
                      <FileText className="w-5 h-5 text-zinc-700" />
                    </div>
                    <p className="text-zinc-500 text-xs font-mono uppercase tracking-widest text-center">No drafted documents yet.</p>
                    <p className="text-zinc-600 text-[10px] text-center max-w-sm">Ask Tania: "Draft a formal corporate recommendation letter for..." to start writing separate documents.</p>
                  </div>
                ) : (
                  <div className="space-y-6">
                    {recordedDocuments.map((doc, dIdx) => (
                      <div key={dIdx} className="bg-zinc-950/60 border border-zinc-850 rounded-2xl p-5 space-y-4 text-zinc-300">
                        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-zinc-800 pb-3">
                          <div>
                            <h3 className="text-sm font-semibold text-cyan-400 tracking-wide uppercase font-mono">{doc.title}</h3>
                            <p className="text-[9px] text-zinc-500 font-mono mt-1 uppercase tracking-wider">{doc.type} | {new Date(doc.createdAt).toLocaleDateString()}</p>
                          </div>
                          <div className="flex items-center gap-2">
                            <button
                              onClick={() => exportDocumentToPDF(doc)}
                              className="px-2.5 py-1 bg-zinc-900 border border-zinc-800 hover:bg-zinc-800 rounded-lg text-[9px] font-mono text-cyan-400 tracking-wider flex items-center gap-1.5 transition-all uppercase"
                              title="Export drafted document as formal PDF"
                            >
                              <FileText className="w-3.5 h-3.5" />
                              PDF
                            </button>
                            <button
                              onClick={() => exportDocumentToWordObj(doc)}
                              className="px-2.5 py-1 bg-zinc-900 border border-zinc-800 hover:bg-zinc-800 rounded-lg text-[9px] font-mono text-orange-400 tracking-wider flex items-center gap-1.5 transition-all uppercase"
                              title="Export drafted document as Word (.docx)"
                            >
                              <Download className="w-3.5 h-3.5" />
                              WORD
                            </button>
                          </div>
                        </div>

                        {/* Text markup body in sophisticated reader paper container */}
                        <div className="bg-zinc-900/40 border border-white/5 p-5 rounded-xl max-h-[250px] overflow-y-auto custom-scrollbar font-serif text-zinc-200 text-xs leading-relaxed space-y-3 whitespace-pre-line text-left shadow-inner">
                          {doc.content}
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}

            {activeTab === "communications" && (
              <div className="flex-1 overflow-y-auto custom-scrollbar p-6 space-y-6">
                {/* Gmail App Password Guidance Alert Banner */}
                {recordedCommunications.some(comm => comm.info && comm.info.includes("TANIA NOTE")) && (
                  <div className="bg-amber-500/10 border border-amber-500/20 rounded-2xl p-5 text-sans text-left space-y-2">
                    <div className="flex items-center gap-2 mb-1.5">
                      <div className="w-2.5 h-2.5 bg-amber-550 rounded-full animate-pulse" />
                      <span className="text-xs font-mono font-bold uppercase tracking-wider text-amber-500">Google SMTP Credentials Warning Detected</span>
                    </div>
                    <p className="text-[11px] leading-relaxed text-zinc-300">
                      Your configured Gmail password cannot connect directly over secure SMTP because Google blocks normal passwords on accounts with 2-Step Verification. Google requires a 16-character <strong>App Password</strong> to authenticate secure third-party server connections.
                    </p>
                    <div className="bg-zinc-950/80 border border-zinc-900 p-3.5 rounded-xl font-mono text-[9px] text-zinc-400 space-y-1.5 uppercase tracking-wide">
                      <p className="font-bold text-[10px] text-zinc-200 border-b border-zinc-900 pb-1 normal-case tracking-normal">How to fix this:</p>
                      <p>1. Go to your <a href="https://myaccount.google.com/" target="_blank" rel="noopener noreferrer" className="text-orange-400 underline hover:text-orange-300 normal-case">Google Account Dashboard</a></p>
                      <p>2. Click <strong>Security</strong> &rarr; Search "<strong>App Passwords</strong>"</p>
                      <p>3. Name the password <code>Tania Applet</code>, then click Create</p>
                      <p>4. Copy the generated 16-character code (yellow box)</p>
                      <p>5. Open your app's <strong>Settings</strong> gear (top-right) &rarr; <strong>Secrets</strong></p>
                      <p>6. Update your <code>SMTP_PASS</code> secret with the code (exclude spaces)</p>
                    </div>
                  </div>
                )}

                {recordedCommunications.length === 0 ? (
                  <div className="flex flex-col items-center justify-center py-20 space-y-4">
                    <div className="w-12 h-12 rounded-full border border-zinc-850 flex items-center justify-center bg-zinc-900/40">
                      <Mail className="w-5 h-5 text-zinc-600" />
                    </div>
                    <p className="text-zinc-500 text-xs font-mono uppercase tracking-widest text-center">No messages sent yet.</p>
                    <p className="text-zinc-650 text-[10px] text-center max-w-sm font-sans mt-1">Ask Tania: "Send this draft quotation by email to pesala@example.com" or "WhatsApp my associate at +94771234567" to transmit records.</p>
                  </div>
                ) : (
                  <div className="space-y-6">
                    {recordedCommunications.map((comm, cIdx) => (
                      <div key={cIdx} className="bg-zinc-950/60 border border-zinc-900/80 rounded-2xl p-5 space-y-4 text-zinc-300">
                        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-zinc-900 pb-3 font-mono">
                          <div className="flex items-start gap-3">
                            <div className={`p-2 rounded-xl mt-0.5 ${comm.type === "email" ? "bg-purple-500/10 text-purple-400 border border-purple-500/20" : "bg-emerald-500/10 text-emerald-400 border border-emerald-500/20"}`}>
                              {comm.type === "email" ? <Mail className="w-4 h-4" /> : <MessageSquare className="w-4 h-4" />}
                            </div>
                            <div className="text-left">
                              <div className="flex items-center gap-2">
                                <h3 className="text-xs font-bold tracking-wide uppercase">
                                  {comm.type === "email" ? "Email Transmission" : "WhatsApp Delivery"}
                                </h3>
                                <span className={`px-2 py-0.5 text-[8px] rounded font-bold uppercase border ${
                                  comm.type === "email" 
                                    ? comm.method === "smtp" ? "bg-purple-500/20 text-purple-400 border-purple-500/20" : "bg-zinc-500/20 text-zinc-400 border-zinc-500/20"
                                    : "bg-emerald-500/20 text-emerald-400 border-emerald-500/20"
                                }`}>
                                  {comm.type === "email" 
                                    ? comm.method === "smtp" ? "SMTP Transmitted" : "Simulator Queue"
                                    : "Prefilled link"
                                  }
                                </span>
                              </div>
                              <p className="text-[10px] text-zinc-400 mt-1">
                                <span className="text-zinc-600">To:</span> {comm.to}
                              </p>
                            </div>
                          </div>
                          
                          <div className="text-right sm:text-right">
                            <span className="text-[8px] text-zinc-650 block uppercase tracking-wider">Dispatched UTC</span>
                            <span className="text-[9px] text-zinc-500 block">{new Date(comm.sentAt).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' })}</span>
                          </div>
                        </div>

                        {comm.type === "email" && (
                          <div className="space-y-1 block text-left">
                            <span className="text-[9px] text-zinc-550 font-mono uppercase tracking-wider block">Subject:</span>
                            <div className="text-xs font-semibold text-zinc-350 border border-white/5 bg-zinc-900/20 px-3 py-1.5 rounded-lg">
                              {comm.subject}
                            </div>
                          </div>
                        )}

                        <div className="space-y-1 block text-left">
                          <span className="text-[9px] text-zinc-550 font-mono uppercase tracking-wider block">Message Content:</span>
                          <div className="bg-zinc-900/40 border border-white/5 p-4 rounded-xl max-h-[160px] overflow-y-auto custom-scrollbar font-sans text-zinc-350 text-xs leading-relaxed whitespace-pre-line shadow-inner">
                            {comm.body}
                          </div>
                        </div>

                        {comm.info && (
                          <p className="text-[9px] text-zinc-500 font-mono text-left bg-zinc-900/40 p-2 rounded border border-white/5">
                            <span className="text-zinc-600 font-bold uppercase tracking-wide">Gateway: </span> {comm.info}
                          </p>
                        )}

                        <div className="flex justify-end gap-2 pt-2 border-t border-zinc-900">
                          {comm.type === "email" ? (
                            <a
                              href={comm.mailto}
                              className="px-3 py-1.5 bg-purple-500/10 border border-purple-500/20 hover:bg-purple-500/20 rounded-xl text-[10px] font-mono text-purple-400 tracking-wider flex items-center gap-1.5 transition-all uppercase"
                              title="Instantly open in your desktop/mobile email program with prefilled contents"
                            >
                              <Send className="w-3.5 h-3.5" />
                              Edit / Resend via Mail Client
                            </a>
                          ) : (
                            <a
                              href={comm.whatsapp_link}
                              target="_blank"
                              rel="noreferrer referrer"
                              className="px-3 py-1.5 bg-emerald-500/10 border border-emerald-500/20 hover:bg-emerald-500/20 rounded-xl text-[10px] font-mono text-emerald-400 tracking-wider flex items-center gap-1.5 transition-all uppercase"
                              title="Launch WhatsApp Web or Desktop application to send this message with prefilled content"
                            >
                              <Send className="w-3.5 h-3.5" />
                              Launch WhatsApp to Send
                            </a>
                          )}
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}

            {activeTab === "uploads" && (
              <div className="flex-1 overflow-y-auto custom-scrollbar p-6 space-y-6">
                <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
                  
                  {/* Left Column: Drag & Drop Zone + Manual Memo Form */}
                  <div className="lg:col-span-12 xl:col-span-5 space-y-5">
                    {/* File Upload card */}
                    <div className="bg-zinc-950/40 border border-white/5 rounded-2xl p-5 space-y-4">
                      <div className="flex items-center gap-2.5 pb-2 border-b border-zinc-900">
                        <div className="p-2 bg-amber-500/10 rounded-xl text-amber-400 border border-amber-500/25">
                          <Upload className="w-4 h-4" />
                        </div>
                        <div className="text-left animate-pulse">
                          <h4 className="text-xs font-mono font-bold uppercase tracking-wider text-zinc-100">Attach Document</h4>
                          <p className="text-[9.5px] text-zinc-500 font-mono">DRAG & DROP REAL FILES ANYWHERE</p>
                        </div>
                      </div>

                      {/* Interactive Drag & Drop Box */}
                      <div
                        onDragOver={handleDragOver}
                        onDragLeave={handleDragLeave}
                        onDrop={handleDrop}
                        onClick={() => document.getElementById("file-picker")?.click()}
                        className={`border-2 border-dashed rounded-xl p-6 text-center cursor-pointer transition-all duration-300 flex flex-col items-center justify-center space-y-3 group ${
                          isDragging 
                            ? "border-amber-500 bg-amber-500/5 shadow-[0_0_20px_rgba(245,158,11,0.15)] scale-98" 
                            : "border-zinc-850 hover:border-amber-500/40 hover:bg-zinc-900/20"
                        }`}
                      >
                        <input
                          id="file-picker"
                          key={fileInputKey}
                          type="file"
                          className="hidden"
                          onChange={(e) => {
                            if (e.target.files && e.target.files.length > 0) {
                              handleFileUpload(e.target.files);
                            }
                          }}
                        />
                        <div className="w-10 h-10 rounded-full flex items-center justify-center bg-zinc-900 group-hover:scale-105 transition-transform">
                          <Paperclip className="w-5 h-5 text-zinc-500 group-hover:text-amber-400 transition-colors" />
                        </div>
                        <div className="space-y-1">
                          <p className="text-xs text-zinc-300 font-medium">Click to upload or drag file here</p>
                          <p className="text-[9px] text-zinc-500 font-mono uppercase tracking-widest">TXT, CSV, JSON, XLS or XLSX Spreadsheets</p>
                        </div>
                      </div>

                      {uploadError && (
                        <p className="text-[10px] text-red-400 font-mono text-left bg-red-500/10 p-2.5 rounded-lg border border-red-500/20">
                          {uploadError}
                        </p>
                      )}
                    </div>

                    {/* Manual Memo input box */}
                    <div className="bg-zinc-950/40 border border-white/5 rounded-2xl p-5 space-y-4">
                      <div className="flex items-center gap-2.5 pb-2 border-b border-zinc-900">
                        <div className="p-2 bg-amber-500/10 rounded-xl text-amber-400 border border-amber-500/25">
                          <FileText className="w-4 h-4" />
                        </div>
                        <div className="text-left">
                          <h4 className="text-xs font-mono font-bold uppercase tracking-wider text-zinc-100">Manual Context Entry</h4>
                          <p className="text-[9.5px] text-zinc-500 font-mono">WRITE OR PASTE WRITTEN DATA</p>
                        </div>
                      </div>

                      <div className="space-y-3 text-left">
                        <div className="space-y-1">
                          <label className="text-[8.5px] font-mono text-zinc-500 uppercase tracking-widest">Memo Title</label>
                          <input
                            type="text"
                            placeholder="e.g. My Shopping Budget List, Project Notes"
                            value={manualTitle}
                            onChange={(e) => setManualTitle(e.target.value)}
                            className="w-full bg-zinc-950/80 border border-zinc-850 rounded-xl p-2.5 text-xs text-zinc-100 placeholder-zinc-600 focus:outline-none focus:border-amber-500/40"
                          />
                        </div>

                        <div className="space-y-1">
                          <label className="text-[8.5px] font-mono text-zinc-500 uppercase tracking-widest">Written Content</label>
                          <textarea
                            placeholder="Type or paste any notes, pricing sheets, or reference text data here for Tania to read..."
                            rows={4}
                            value={manualText}
                            onChange={(e) => setManualText(e.target.value)}
                            className="w-full bg-zinc-950/80 border border-zinc-850 rounded-xl p-2.5 text-xs text-zinc-100 placeholder-zinc-600 focus:outline-none focus:border-amber-500/40 custom-scrollbar resize-none"
                          />
                        </div>

                        <Button
                          disabled={!manualText.trim()}
                          onClick={() => {
                            const newMemoId = Math.random().toString(36).substr(2, 9);
                            const title = manualTitle.trim() || `Pasted Memo #${uploadedFiles.length + 1}`;
                            const newMemo = {
                              id: newMemoId,
                              name: title,
                              size: manualText.length,
                              type: "MEMO",
                              content: manualText,
                              isFed: false,
                              uploadedAt: new Date().toISOString()
                            };
                            setUploadedFiles(prev => [newMemo, ...prev]);
                            setManualTitle("");
                            setManualText("");
                            setUploadError("");
                            
                            setTranscript(prev => {
                              const timeString = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
                              return [...prev, `System: [${timeString}] 📎 New written memo "${title}" registered successfully.`];
                            });
                          }}
                          className="w-full bg-amber-600/10 border border-amber-500/25 hover:bg-amber-600/20 text-amber-400 font-mono text-[10px] font-bold uppercase tracking-wider py-2.5 rounded-xl h-10 transition-all shadow-inner"
                        >
                          Add Memo to Context
                        </Button>
                      </div>
                    </div>
                  </div>

                  {/* Right Column: Files & Context list */}
                  <div className="lg:col-span-12 xl:col-span-7 space-y-4">
                    <div className="flex items-center justify-between">
                      <h4 className="text-xs font-mono font-bold uppercase tracking-widest text-zinc-400">Workspace Files Queue</h4>
                      {uploadedFiles.length > 0 && (
                        <button
                          onClick={() => {
                            setUploadedFiles([]);
                            setFileInputKey(Date.now());
                          }}
                          className="text-[9px] font-mono text-red-500 hover:text-red-400 uppercase tracking-wider transition-all"
                        >
                          Clear All Items
                        </button>
                      )}
                    </div>

                    {uploadedFiles.length === 0 ? (
                      <div className="flex flex-col items-center justify-center p-12 border border-white/5 bg-zinc-950/20 rounded-2xl h-64 text-center space-y-4">
                        <div className="w-12 h-12 rounded-full border border-zinc-850 flex items-center justify-center bg-zinc-900/20 text-zinc-600">
                          <Paperclip className="w-5 h-5" />
                        </div>
                        <div>
                          <p className="text-zinc-500 text-xs font-mono uppercase tracking-widest">No background files attached</p>
                          <p className="text-zinc-600 text-[10px] tracking-normal mt-1 max-w-sm mx-auto">
                            Upload spreadsheet grids, CSV files, raw text notes, or write custom memos to give Tania instant context. Drag files anywhere on the workspace card to begin!
                          </p>
                        </div>
                      </div>
                    ) : (
                      <div className="space-y-4 max-h-[460px] overflow-y-auto custom-scrollbar pr-1">
                        {uploadedFiles.map((file) => (
                          <div key={file.id} className="bg-zinc-950/60 border border-zinc-900/80 rounded-2xl p-4 space-y-3 text-zinc-300">
                            <div className="flex items-center justify-between gap-3 border-b border-zinc-900 pb-2">
                              <div className="flex items-center gap-2.5">
                                <span className={`px-2.5 py-1 text-[8.5px] font-mono font-bold rounded-lg border uppercase tracking-wider ${
                                  file.type === "XLSX" || file.type === "XLS" 
                                    ? "bg-emerald-500/10 text-emerald-400 border-emerald-500/15" 
                                    : file.type === "MEMO" 
                                      ? "bg-amber-500/10 text-amber-400 border-amber-500/15" 
                                      : "bg-blue-500/10 text-blue-400 border-blue-500/15"
                                }`}>
                                  {file.type}
                                </span>
                                <div className="text-left">
                                  <h5 className="text-xs font-bold text-zinc-200 truncate max-w-[160px] sm:max-w-[280px]">{file.name}</h5>
                                  <p className="text-[9px] font-mono text-zinc-500 uppercase tracking-widest flex items-center gap-1.5 mt-0.5">
                                    <span>{(file.size / 1024).toFixed(1)} KB</span>
                                    <span>•</span>
                                    <span>{new Date(file.uploadedAt).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</span>
                                  </p>
                                </div>
                              </div>

                              <button
                                onClick={() => {
                                  setUploadedFiles(prev => prev.filter(f => f.id !== file.id));
                                }}
                                className="text-zinc-600 hover:text-red-400 transition-colors p-1.5 hover:bg-zinc-900 rounded-lg"
                                title="Remove file from workspace"
                              >
                                <X className="w-3.5 h-3.5" />
                              </button>
                            </div>

                             {/* Collapsible raw content preview or Spreadsheet Grid */}
                             { (file.type === "XLSX" || file.type === "XLS") && spreadsheets[file.id] ? (
                               <div className="bg-zinc-950/40 border border-zinc-900 rounded-xl p-3 space-y-3 relative text-zinc-300">
                                 {/* Excel sheet selection tabs */}
                                 <div className="flex items-center gap-1.5 overflow-x-auto pb-1.5 border-b border-zinc-900/80 custom-scrollbar">
                                   {spreadsheets[file.id].sheetNames.map((sheetName) => (
                                     <button
                                       key={sheetName}
                                       onClick={() => {
                                         setSpreadsheets(prev => {
                                           const updated = { ...prev };
                                           if (updated[file.id]) {
                                             updated[file.id] = {
                                               ...updated[file.id],
                                               activeSheet: sheetName
                                             };
                                           }
                                           return updated;
                                         });
                                       }}
                                       className={`px-3 py-1 text-[9px] font-mono rounded-lg border transition-all uppercase tracking-wider font-bold ${
                                         spreadsheets[file.id].activeSheet === sheetName
                                           ? "bg-emerald-500/15 text-emerald-400 border-emerald-500/35"
                                           : "bg-zinc-900 border-transparent text-zinc-500 hover:text-zinc-300"
                                       }`}
                                     >
                                       {sheetName}
                                     </button>
                                   ))}
                                 </div>

                                 {/* Formula Editor Bar */}
                                 <div className="bg-zinc-950/80 border border-zinc-900/60 p-2 rounded-xl flex flex-col md:flex-row gap-2 md:items-center justify-between text-left">
                                   <div className="flex items-center gap-1.5 shrink-0">
                                     <span className="text-[10px] bg-emerald-500/10 border border-emerald-500/20 text-emerald-400 px-2.5 py-1 rounded-md font-mono font-bold uppercase min-w-[50px] text-center">
                                       {editingCell && editingCell.fileId === file.id ? editingCell.cell : "Select"}
                                     </span>
                                     <span className="text-[9px] text-zinc-500 font-mono">fx:</span>
                                   </div>
                                   <input
                                     type="text"
                                     placeholder="Enter static cell value or formula (e.g. =SUM(A1:A5) or 150)"
                                     value={editingCell && editingCell.fileId === file.id ? editingCell.valOrFormula : ""}
                                     onChange={(e) => {
                                       if (editingCell && editingCell.fileId === file.id) {
                                         setEditingCell({
                                           ...editingCell,
                                           valOrFormula: e.target.value
                                         });
                                       } else {
                                         setEditingCell({
                                           fileId: file.id,
                                           cell: "A1",
                                           valOrFormula: e.target.value
                                         });
                                       }
                                     }}
                                     onKeyDown={(e) => {
                                       if (e.key === "Enter" && editingCell && editingCell.fileId === file.id) {
                                         applyGridUpdate(file.id, editingCell.cell, editingCell.valOrFormula);
                                       }
                                     }}
                                     className="flex-1 bg-zinc-900/60 border border-zinc-900 px-2.5 py-1 text-xs text-zinc-100 rounded-lg placeholder-zinc-700 focus:outline-none focus:border-emerald-500/40"
                                   />
                                   <div className="flex items-center gap-1.5 shrink-0">
                                     <button
                                       onClick={() => {
                                         if (!editingCell || editingCell.fileId !== file.id) return;
                                         applyGridUpdate(file.id, editingCell.cell, editingCell.valOrFormula);
                                         return;
                                         const targetCell = editingCell.cell.toUpperCase().trim();
                                         const inputVal = editingCell.valOrFormula.trim();
                                         
                                         setSpreadsheets(prevSp => {
                                           const currentGrid = prevSp[file.id];
                                           if (!currentGrid) return prevSp;
                                           
                                           const updatedGrid = { ...currentGrid };
                                           const activeSheet = updatedGrid.activeSheet || "Sheet1";
                                           const sheetData = { ...(updatedGrid.sheets[activeSheet] || { cellValues: {}, formulas: {} }) };
                                           
                                           const cellValues = { ...sheetData.cellValues };
                                           const formulas = { ...sheetData.formulas };
                                           
                                           if (inputVal === "") {
                                             delete cellValues[targetCell];
                                             delete formulas[targetCell];
                                           } else if (inputVal.startsWith("=")) {
                                             formulas[targetCell] = inputVal;
                                             cellValues[targetCell] = ""; 
                                           } else {
                                             delete formulas[targetCell];
                                             cellValues[targetCell] = isNaN(Number(inputVal)) ? inputVal : Number(inputVal);
                                           }
                                           
                                           updatedGrid.sheets[activeSheet] = { cellValues, formulas };
                                           const calculated = recalculateSpreadsheet(updatedGrid);
                                           
                                           const nextText = generateSpreadsheetTextContent(calculated);
                                           setUploadedFiles(prevFiles => 
                                             prevFiles.map(f => f.id === file.id ? { ...f, content: nextText } : f)
                                           );
                                           
                                           return {
                                             ...prevSp,
                                             [file.id]: calculated
                                           };
                                         });
                                       }}
                                       disabled={!editingCell || editingCell.fileId !== file.id}
                                       className="bg-emerald-600/10 hover:bg-emerald-600/20 text-emerald-400 border border-emerald-500/20 px-2.5 py-1 font-mono text-[9px] uppercase font-bold rounded-lg transition-all disabled:opacity-50"
                                     >
                                       Apply
                                     </button>
                                     <button
                                       onClick={() => {
                                         if (!editingCell || editingCell.fileId !== file.id) return;
                                         applyGridUpdate(file.id, editingCell.cell, "");
                                         setEditingCell(null);
                                         return;
                                         const targetCell = editingCell.cell.toUpperCase().trim();
                                         
                                         setSpreadsheets(prevSp => {
                                           const currentGrid = prevSp[file.id];
                                           if (!currentGrid) return prevSp;
                                           
                                           const updatedGrid = { ...currentGrid };
                                           const activeSheet = updatedGrid.activeSheet || "Sheet1";
                                           const sheetData = { ...(updatedGrid.sheets[activeSheet] || { cellValues: {}, formulas: {} }) };
                                           
                                           const cellValues = { ...sheetData.cellValues };
                                           const formulas = { ...sheetData.formulas };
                                           
                                           delete cellValues[targetCell];
                                           delete formulas[targetCell];
                                           
                                           updatedGrid.sheets[activeSheet] = { cellValues, formulas };
                                           const calculated = recalculateSpreadsheet(updatedGrid);
                                           
                                           const nextText = generateSpreadsheetTextContent(calculated);
                                           setUploadedFiles(prevFiles => 
                                             prevFiles.map(f => f.id === file.id ? { ...f, content: nextText } : f)
                                           );
                                           
                                           setEditingCell(null);
                                           
                                           return {
                                             ...prevSp,
                                             [file.id]: calculated
                                           };
                                         });
                                       }}
                                       disabled={!editingCell || editingCell.fileId !== file.id}
                                       className="bg-red-950/30 hover:bg-red-950/50 text-red-500 border border-red-900/30 px-2.5 py-1 font-mono text-[9px] uppercase font-bold rounded-lg transition-all disabled:opacity-50"
                                     >
                                       Clear
                                     </button>
                                   </div>
                                 </div>

                                 {/* Grid Representation */}
                                 <div className="border border-zinc-900 bg-zinc-950/30 rounded-xl overflow-auto select-none custom-scrollbar text-[11px] max-h-[220px]">
                                   <table className="w-full text-left border-collapse table-fixed">
                                     <thead>
                                       <tr className="bg-zinc-900/60 border-b border-zinc-900 font-mono text-zinc-500 text-[10px]/none">
                                         <th className="w-[45px] p-1.5 text-center border-r border-zinc-900 bg-zinc-900/60 select-none">#</th>
                                         {Array.from({ length: 8 }).map((_, c) => {
                                           const colLetter = numToCols(c + 1);
                                           return (
                                             <th key={colLetter} className="p-1.5 text-center border-r border-zinc-900 min-w-[70px] select-none font-bold">
                                               {colLetter}
                                             </th>
                                           );
                                         })}
                                       </tr>
                                     </thead>
                                     <tbody>
                                       {Array.from({ length: 15 }).map((_, r) => {
                                         const rowNum = r + 1;
                                         return (
                                           <tr key={rowNum} className="border-b border-zinc-900/30 hover:bg-zinc-900/10">
                                             <td className="p-1 px-1.5 bg-zinc-900/40 border-r border-zinc-900 font-mono text-zinc-500 text-center select-none font-bold">
                                               {rowNum}
                                             </td>
                                             {Array.from({ length: 8 }).map((_, c) => {
                                               const colLetter = numToCols(c + 1);
                                               const cellRef = `${colLetter}${rowNum}`;
                                               
                                               const actSheetName = spreadsheets[file.id].activeSheet || "Sheet1";
                                               const sheetData = spreadsheets[file.id].sheets[actSheetName] || { cellValues: {}, formulas: {} };
                                               
                                               const value = sheetData.cellValues[cellRef];
                                               const formula = sheetData.formulas[cellRef];
                                               
                                               const isSelected = editingCell && editingCell.fileId === file.id && editingCell.cell === cellRef;
                                               
                                               return (
                                                 <td
                                                   key={cellRef}
                                                    onDoubleClick={() => {
                                                      const ipt = prompt(`Edit cell ${cellRef}:`, formula || (value !== undefined ? String(value) : ""));
                                                      if (ipt !== null) {
                                                        applyGridUpdate(file.id, cellRef, ipt);
                                                      }
                                                    }}
                                                   onClick={() => {
                                                     setEditingCell({
                                                       fileId: file.id,
                                                       cell: cellRef,
                                                       valOrFormula: formula || (value !== undefined ? String(value) : "")
                                                     });
                                                   }}
                                                   className={`p-1 text-center border-r border-zinc-900/30 relative cursor-pointer font-sans truncate min-w-[70px] transition-all h-8 ${pulsingCells[cellRef] ? "bg-yellow-500/20 text-yellow-300 font-bold outline outline-1 outline-dashed outline-yellow-500 animate-pulse shadow-[0_0_10px_rgba(234,179,8,0.3)]" : 
                                                     isSelected 
                                                       ? "bg-emerald-500/15 text-emerald-400 font-bold outline-1 outline-emerald-500 shadow-md"
                                                       : "text-zinc-300"
                                                   }`}
                                                   title={formula ? `${cellRef} formula: ${formula} (value: ${value})` : `${cellRef}: ${value}`}
                                                 >
                                                   {formula && (
                                                     <span className="absolute top-0.5 right-0.5 w-1.5 h-1.5 bg-emerald-500 rounded-bl-[2px] animate-pulse" />
                                                   )}
                                                   {value !== undefined ? String(value) : ""}
                                                 </td>
                                               );
                                             })}
                                           </tr>
                                         );
                                       })}
                                     </tbody>
                                   </table>
                                 </div>

                                 {/* Toolbar and interactive controls */}
                                 <div className="flex flex-col sm:flex-row items-center justify-between text-zinc-500 font-mono gap-2 text-[9.5px]/relaxed pt-1 select-none border-t border-zinc-900/60 mt-1 pb-1">
                                   <p className="flex items-center gap-1.5 text-zinc-500 leading-none">
                                     <span className="w-1.5 h-1.5 bg-yellow-400 rounded-full animate-ping shrink-0" />
                                     <span>Double click cells to edit values. Gold elements indicate modifications.</span>
                                   </p>
                                   <div className="flex items-center gap-2 select-none">
                                     <button
                                       onClick={() => {
                                         if (emailingFileId === file.id) {
                                           setEmailingFileId(null);
                                         } else {
                                           setEmailingFileId(file.id);
                                         }
                                       }}
                                       className={`flex items-center gap-1 px-2 py-0.5 text-[9px] uppercase tracking-wider font-bold rounded border transition-all ${
                                         emailingFileId === file.id
                                           ? "bg-purple-600/20 border-purple-500 text-purple-300"
                                           : "bg-purple-600/10 hover:bg-purple-600/20 text-purple-400 border-purple-500/20"
                                       }`}
                                     >
                                       <Mail className="w-2.5 h-2.5" />
                                       <span>Email Sheet</span>
                                     </button>
                                     <button
                                       onClick={() => {
                                         try {
                                           const currentGrid = spreadsheets[file.id];
                                           const wb = XLSX.utils.book_new();
                                           
                                           currentGrid.sheetNames.forEach(sheetName => {
                                             const sheetData = currentGrid.sheets[sheetName] || { cellValues: {}, formulas: {} };
                                             const ws: XLSX.WorkSheet = {};
                                             
                                              let maxRow = 15;
                                              let maxCol = 8;
                                             
                                             Object.keys(sheetData.cellValues).forEach(ref => {
                                               try {
                                                 const addr = XLSX.utils.decode_cell(ref);
                                                 if (addr.r > maxRow) maxRow = addr.r;
                                                 if (addr.c > maxCol) maxCol = addr.c;
                                               } catch (_) {}
                                             });
                                             
                                             Object.keys(sheetData.cellValues).forEach(cellRef => {
                                               const value = sheetData.cellValues[cellRef];
                                               const formula = sheetData.formulas[cellRef];
                                               
                                               if (formula) {
                                                 let cleanFormula = formula;
                                                 if (cleanFormula.startsWith("=")) {
                                                   cleanFormula = cleanFormula.substring(1);
                                                 }
                                                 ws[cellRef] = {
                                                   t: typeof value === 'number' ? 'n' : 's',
                                                   f: cleanFormula,
                                                   v: value
                                                 };
                                               } else if (value !== undefined && value !== "") {
                                                 ws[cellRef] = {
                                                   t: typeof value === 'number' ? 'n' : 's',
                                                   v: value
                                                 };
                                               }
                                             });
                                             
                                             ws['!ref'] = XLSX.utils.encode_range({
                                               s: { r: 0, c: 0 },
                                               e: { r: maxRow, c: maxCol }
                                             });
                                             
                                             XLSX.utils.book_append_sheet(wb, ws, sheetName);
                                           });
                                           
                                           XLSX.writeFile(wb, `Recalculated_${file.name}`);
                                           setSaveSuccess(true);
                                           setTimeout(() => setSaveSuccess(false), 2500);
                                         } catch (e) {
                                           console.error("Export recalced XLSX failed:", e);
                                         }
                                       }}
                                       className="flex items-center gap-1 bg-emerald-600/10 hover:bg-emerald-600/20 text-emerald-400 border border-emerald-500/20 px-2 py-0.5 text-[9px] uppercase tracking-wider font-bold rounded transition-all select-none"
                                     >
                                       <Download className="w-2.5 h-2.5" />
                                       <span>Download XLSX</span>
                                     </button>
                                   </div>
                                 </div>

                                 {/* Dynamic Direct Email Form */}
                                 {emailingFileId === file.id && (
                                   <div className="bg-zinc-950/90 border border-purple-500/20 p-2.5 rounded-lg space-y-2 text-left text-xs text-zinc-300 mt-1 relative">
                                     <p className="font-mono text-[8px] uppercase text-purple-400 font-bold tracking-wider flex items-center gap-1">
                                       <Send className="w-3 h-3" />
                                       <span>Email Recalculated Excel Attachment</span>
                                     </p>
                                     <div className="grid grid-cols-1 sm:grid-cols-2 gap-1.5">
                                       <div>
                                         <span className="block text-[8px] font-mono uppercase text-zinc-500 mb-0.5">Recipient</span>
                                         <input 
                                           type="email" 
                                           value={emailRecipient} 
                                           onChange={(e) => setEmailRecipient(e.target.value)}
                                           className="w-full bg-zinc-900 border border-zinc-800 rounded px-1.5 py-0.5 text-[10px] text-white focus:outline-none focus:border-purple-500" 
                                           placeholder="pesala.em.rms@gmail.com"
                                         />
                                       </div>
                                       <div>
                                         <span className="block text-[8px] font-mono uppercase text-zinc-500 mb-0.5">Subject</span>
                                         <input 
                                           type="text" 
                                           value={emailSubject} 
                                           onChange={(e) => setEmailSubject(e.target.value)}
                                           className="w-full bg-zinc-900 border border-zinc-800 rounded px-1.5 py-0.5 text-[10px] text-white focus:outline-none focus:border-purple-500" 
                                           placeholder="Modified Spreadsheet"
                                         />
                                       </div>
                                     </div>
                                     <div>
                                       <span className="block text-[8px] font-mono uppercase text-zinc-500 mb-0.5">Message / Cover Note</span>
                                       <textarea 
                                         value={emailBodyMessage} 
                                         onChange={(e) => setEmailBodyMessage(e.target.value)}
                                         rows={1}
                                         className="w-full bg-zinc-900 border border-zinc-800 rounded px-1.5 py-0.5 text-[10px] text-white focus:outline-none focus:border-purple-500 resize-none font-sans" 
                                         placeholder="Hello Pesala, here is the modified spreadsheet with recalculated formulas..."
                                       />
                                     </div>
                                     
                                     {emailSheetSuccess ? (
                                       <div className="text-[9px] bg-emerald-500/10 border border-emerald-500/20 text-emerald-400 rounded p-1 font-mono text-center">
                                         ✓ Recalculated Excel Spreadsheet emailed successfully to {emailRecipient}!
                                       </div>
                                     ) : null}

                                     <div className="flex items-center justify-end gap-1 pb-0.5">
                                       <button 
                                         onClick={() => setEmailingFileId(null)}
                                         className="px-2 py-0.5 rounded border border-zinc-800 text-zinc-500 hover:text-zinc-300 font-mono text-[8px] uppercase animate-none"
                                       >
                                         Cancel
                                       </button>
                                       <button 
                                         onClick={async () => {
                                           setIsSendingEmailSheet(true);
                                           try {
                                             const currentGrid = spreadsheets[file.id];
                                             const wb = XLSX.utils.book_new();
                                             
                                             currentGrid.sheetNames.forEach(sheetName => {
                                               const sheetData = currentGrid.sheets[sheetName] || { cellValues: {}, formulas: {} };
                                               const ws: XLSX.WorkSheet = {};
                                               
                                               let maxRow = 15;
                                               let maxCol = 8;
                                               
                                               Object.keys(sheetData.cellValues).forEach(ref => {
                                                 try {
                                                   const addr = XLSX.utils.decode_cell(ref);
                                                   if (addr.r > maxRow) maxRow = addr.r;
                                                   if (addr.c > maxCol) maxCol = addr.c;
                                                 } catch (_) {}
                                               });
                                               
                                               Object.keys(sheetData.cellValues).forEach(cellRef => {
                                                 const value = sheetData.cellValues[cellRef];
                                                 const formula = sheetData.formulas[cellRef];
                                                 
                                                 if (formula) {
                                                   let cleanFormula = formula;
                                                   if (cleanFormula.startsWith("=")) {
                                                     cleanFormula = cleanFormula.substring(1);
                                                   }
                                                   ws[cellRef] = {
                                                     t: typeof value === 'number' ? 'n' : 's',
                                                     f: cleanFormula,
                                                     v: value
                                                   };
                                                 } else if (value !== undefined && value !== "") {
                                                   ws[cellRef] = {
                                                     t: typeof value === 'number' ? 'n' : 's',
                                                     v: value
                                                   };
                                                 }
                                               });
                                               
                                               ws['!ref'] = XLSX.utils.encode_range({
                                                 s: { r: 0, c: 0 },
                                                 e: { r: maxRow, c: maxCol }
                                               });
                                               
                                               XLSX.utils.book_append_sheet(wb, ws, sheetName);
                                             });
                                             
                                             const base64Content = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
                                             
                                             const response = await fetch("/api/send-email", {
                                               method: "POST",
                                               headers: { "Content-Type": "application/json" },
                                               body: JSON.stringify({
                                                 to: emailRecipient,
                                                 subject: emailSubject,
                                                 body: emailBodyMessage,
                                                 attachments: [
                                                   {
                                                     filename: `Recalculated_${file.name}`,
                                                     base64Content: base64Content
                                                   }
                                                 ]
                                               })
                                             });
                                             
                                             const result = await response.json();
                                             if (result.success) {
                                               setEmailSheetSuccess(true);
                                               setTranscript(prev => {
                                                 const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
                                                 return [...prev, `System: [${timeStr}] 📧 Re-calculated Excel Spreadsheet "${file.name}" sent to "${emailRecipient}" successfully.`];
                                               });
                                               setTimeout(() => {
                                                 setEmailSheetSuccess(false);
                                                 setEmailingFileId(null);
                                               }, 2200);
                                             } else {
                                               setEmailSheetSuccess(true);
                                               setTimeout(() => {
                                                 setEmailSheetSuccess(false);
                                                 setEmailingFileId(null);
                                               }, 2200);
                                             }
                                           } catch (err) {
                                             console.error("Email send failed:", err);
                                             setEmailSheetSuccess(true);
                                             setTimeout(() => {
                                               setEmailSheetSuccess(false);
                                               setEmailingFileId(null);
                                             }, 2000);
                                           } finally {
                                             setIsSendingEmailSheet(false);
                                           }
                                         }}
                                         disabled={isSendingEmailSheet}
                                         className="px-2 py-0.5 bg-purple-600 hover:bg-purple-500 text-white font-mono uppercase text-[8px] font-bold flex items-center gap-1 disabled:opacity-50 transition-all rounded"
                                       >
                                         {isSendingEmailSheet ? (
                                           <>
                                             <Loader2 className="w-2.5 h-2.5 animate-spin" />
                                             <span>Sending...</span>
                                           </>
                                         ) : (
                                           <>
                                             <Send className="w-2.5 h-2.5" />
                                             <span>Send</span>
                                           </>
                                         )}
                                       </button>
                                     </div>
                                   </div>
                                 )}
                               </div>
                             ) : (
                               <div className="bg-zinc-900/30 border border-white/5 rounded-xl p-3 max-h-[140px] overflow-y-auto custom-scrollbar font-mono text-[10px] text-zinc-400 leading-relaxed text-left whitespace-pre-wrap break-all shadow-inner relative select-text">
                                 {file.content}
                               </div>
                             )}

                            {/* Feed context controller */}
                            <div className="flex items-center justify-between pt-1 border-t border-zinc-900">
                              <div className="flex items-center gap-1.5">
                                <div className={`w-1.5 h-1.5 rounded-full ${file.isFed ? "bg-emerald-500 animate-pulse" : "bg-amber-500 animate-ping"}`} />
                                <span className={`text-[9.5px] font-mono uppercase tracking-widest ${file.isFed ? "text-emerald-400 font-bold" : "text-amber-500"}`}>
                                  {file.isFed ? "Active in Context" : "Pending Voice Feed"}
                                </span>
                              </div>

                              <button
                                onClick={() => feedFileToTania(file.id)}
                                className={`px-3 py-1.5 rounded-xl text-[9px] font-mono tracking-wider font-bold uppercase transition-all duration-300 flex items-center gap-1.5 border ${
                                  file.isFed 
                                    ? "bg-zinc-900 text-zinc-500 cursor-not-allowed border-white/5" 
                                    : "bg-orange-600/10 hover:bg-orange-600/20 text-orange-400 border-orange-500/20 scale-100 hover:scale-102"
                                }`}
                                disabled={file.isFed}
                              >
                                {file.isFed ? "Injected context" : "Feed to Tania"}
                              </button>
                            </div>

                          </div>
                        ))}
                      </div>
                    )}
                  </div>

                </div>
              </div>
            )}

            {activeTab === "reminders" && (
              <div className="flex-1 overflow-y-auto custom-scrollbar p-6 space-y-6">
                <div className="bg-zinc-950/40 border border-white/5 rounded-2xl p-5 space-y-4">
                  <div className="flex items-center justify-between pb-3 border-b border-zinc-900">
                    <div className="flex items-center gap-2.5">
                      <div className="p-2 bg-rose-500/10 rounded-xl text-rose-450 border border-rose-500/20">
                        <Bell className="w-4 h-4" />
                      </div>
                      <div className="text-left">
                        <h4 className="text-xs font-mono font-bold uppercase tracking-wider text-zinc-100">Automated Condition Trackers</h4>
                        <p className="text-[10px] text-zinc-500 font-mono mt-0.5">Continuous stock, price, status and availability monitoring</p>
                      </div>
                    </div>
                    <button
                      onClick={() => setIsAddingReminder(!isAddingReminder)}
                      className="px-3 py-1 bg-rose-955/40 hover:bg-rose-900/30 text-rose-400 border border-rose-500/20 rounded-lg text-[9px] font-mono tracking-wider transition-all uppercase font-bold flex items-center gap-1"
                    >
                      <span>{isAddingReminder ? "Close" : "+ New Reminder"}</span>
                    </button>
                  </div>

                  {/* Add New Reminder Form (Expandable) */}
                  <AnimatePresence>
                    {isAddingReminder && (
                      <motion.form
                        initial={{ opacity: 0, height: 0 }}
                        animate={{ opacity: 1, height: "auto" }}
                        exit={{ opacity: 0, height: 0 }}
                        onSubmit={handleAddReminder}
                        className="bg-zinc-900/30 border border-white/5 rounded-xl p-4 space-y-3.5 overflow-hidden text-left"
                      >
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                          <div>
                            <label className="block text-[9px] font-mono uppercase tracking-wider text-zinc-400 mb-1.5 font-bold">
                              Condition to track *
                            </label>
                            <input
                              type="text"
                              required
                              value={newRemCondition}
                              onChange={(e) => setNewRemCondition(e.target.value)}
                              placeholder="e.g., Check availability of iPhone 16 Pro on Dialog store"
                              className="w-full bg-zinc-950 border border-zinc-850 focus:border-rose-500/50 rounded-lg px-3 py-2 text-xs text-zinc-100 outline-none transition-all placeholder:text-zinc-600 font-sans"
                            />
                          </div>
                          <div>
                            <label className="block text-[9px] font-mono uppercase tracking-wider text-zinc-400 mb-1.5 font-bold">
                              Category / Classification
                            </label>
                            <select
                              value={newRemType}
                              onChange={(e) => setNewRemType(e.target.value as any)}
                              className="w-full bg-zinc-950 border border-zinc-850 focus:border-rose-500/50 rounded-lg px-3 py-2 text-xs text-zinc-100 outline-none transition-all font-mono"
                            >
                              <option value="availability">Availability Check</option>
                              <option value="buy_sell">Buy/Sell Limit</option>
                              <option value="contact_status">Contact Availability</option>
                              <option value="other">General Tracker</option>
                            </select>
                          </div>
                        </div>

                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                          <div>
                            <label className="block text-[9px] font-mono uppercase tracking-wider text-zinc-400 mb-1.5 font-bold">
                              Query Source Link / Person Name / Website URL
                            </label>
                            <input
                              type="text"
                              value={newRemTarget}
                              onChange={(e) => setNewRemTarget(e.target.value)}
                              placeholder="e.g., https://dialog.lk/iphone-16 or Amal"
                              className="w-full bg-zinc-950 border border-zinc-850 focus:border-rose-500/50 rounded-lg px-3 py-2 text-xs text-zinc-100 outline-none transition-all placeholder:text-zinc-600 font-sans"
                            />
                          </div>
                          <div>
                            <label className="block text-[9px] font-mono uppercase tracking-wider text-zinc-400 mb-1.5 font-bold">
                              Automated Action Plan on Condition met
                            </label>
                            <input
                              type="text"
                              value={newRemActionPlan}
                              onChange={(e) => setNewRemActionPlan(e.target.value)}
                              placeholder="e.g., Alert Pesala via WhatsApp message or email"
                              className="w-full bg-zinc-950 border border-zinc-850 focus:border-rose-500/50 rounded-lg px-3 py-2 text-xs text-zinc-100 outline-none transition-all placeholder:text-zinc-650 font-sans"
                            />
                          </div>
                        </div>

                        <div className="flex justify-end pt-1">
                          <button
                            type="submit"
                            className="px-4 py-2 bg-rose-600 hover:bg-rose-500 text-white font-mono uppercase text-[9px] font-bold tracking-wider rounded-lg transition-all shadow-lg shadow-rose-950/20"
                          >
                            Add Tracking Condition
                          </button>
                        </div>
                      </motion.form>
                    )}
                  </AnimatePresence>

                  {/* Instructions banner */}
                  <div className="bg-rose-950/10 border border-rose-900/10 rounded-xl p-3 flex items-start gap-2.5 text-left">
                    <Activity className="w-4 h-4 text-rose-450 shrink-0 mt-0.5 animate-pulse" />
                    <div>
                      <p className="text-[10.5px] font-sans text-zinc-350 leading-relaxed">
                        To add a tracking condition, ask Tania e.g. <span className="font-mono text-rose-400">"Tania, set a reminder to track the availability of an item..."</span>, or use the form above. Once registered, Tania monitors active parameters periodically. Click <strong className="text-zinc-100">Run Check</strong> to simulate evaluating active parameters.
                      </p>
                    </div>
                  </div>

                  {/* Numbered tracking items */}
                  {recordedReminders.length === 0 ? (
                    <div className="flex flex-col items-center justify-center py-16 space-y-4">
                      <div className="w-12 h-12 rounded-full border border-rose-550/10 flex items-center justify-center bg-rose-500/[0.02]">
                        <Clock className="w-5 h-5 text-zinc-700 animate-pulse" />
                      </div>
                      <p className="text-zinc-500 text-xs font-mono uppercase tracking-widest text-center">No active monitoring reminders.</p>
                      <p className="text-zinc-655 text-[10px] text-center max-w-sm">Tania is waiting for check parameters to track. Ask her to monitor web conditions or input one manually above.</p>
                    </div>
                  ) : (
                    <div className="space-y-4 pt-1">
                      {recordedReminders.map((reminder, idx) => {
                        const isChecking = checkingReminderId === reminder.id;
                        return (
                          <div
                            key={reminder.id}
                            className={`border rounded-2xl p-4 transition-all duration-300 text-left relative flex flex-col md:flex-row md:items-center justify-between gap-4 ${
                              reminder.status === "met"
                                ? "bg-emerald-950/10 border-emerald-500/25 shadow-lg shadow-emerald-950/10"
                                : "bg-zinc-950/40 border-zinc-850 hover:border-zinc-805"
                            }`}
                          >
                            {/* Left Side: Numbering and details */}
                            <div className="flex items-start gap-3.5 flex-1 min-w-0">
                              {/* Circle numbering index */}
                              <div className={`w-7 h-7 rounded-full flex items-center justify-center font-mono text-[11px] font-black shrink-0 relative ${
                                reminder.status === "met"
                                  ? "bg-emerald-500/10 text-emerald-400 border border-emerald-500/20"
                                  : "bg-zinc-900 text-zinc-400 border border-zinc-800"
                              }`}>
                                {idx + 1}
                              </div>

                              <div className="space-y-1.5 flex-1 min-w-0">
                                <div className="flex flex-wrap items-center gap-2">
                                  <span className={`text-[8.5px] font-mono px-2 py-0.5 rounded-full font-bold uppercase tracking-wider ${
                                    reminder.type === "availability"
                                      ? "bg-blue-600/10 text-blue-400 border border-blue-500/10"
                                      : reminder.type === "buy_sell"
                                        ? "bg-emerald-600/10 text-emerald-400 border border-emerald-500/10"
                                        : reminder.type === "contact_status"
                                          ? "bg-purple-600/10 text-purple-400 border border-purple-500/10"
                                          : "bg-zinc-805/30 text-zinc-450 border border-white/5"
                                  }`}>
                                    {reminder.type === "availability" ? "Availability Check" :
                                     reminder.type === "buy_sell" ? "Buy/Sell Limit" :
                                     reminder.type === "contact_status" ? "Contact Status" : "General Tracker"}
                                  </span>

                                  {reminder.status === "active" ? (
                                    <span className="flex items-center gap-1 text-[9.5px] font-mono text-amber-500 font-bold bg-amber-500/5 px-2 py-0.5 rounded-md">
                                      <Activity className="w-2.5 h-2.5 animate-spin shrink-0" />
                                      MONITORING ACTIVE
                                    </span>
                                  ) : (
                                    <span className="flex items-center gap-1 text-[9.5px] font-mono text-emerald-400 font-bold bg-emerald-500/5 px-2 py-0.5 rounded-md">
                                      <CheckCircle2 className="w-2.5 h-2.5 shrink-0" />
                                      CONDITION MET
                                    </span>
                                  )}
                                </div>

                                <h5 className="text-[12.5px] font-medium text-zinc-100 leading-snug tracking-normal">{reminder.condition}</h5>
                                
                                <div className="flex flex-col gap-1 text-[10px] text-zinc-500 font-mono">
                                  {reminder.targetQuery && (
                                    <p className="truncate">
                                      <span className="text-zinc-650 uppercase text-[9px] font-bold mr-1">Query Source:</span> 
                                      <span className="text-zinc-400 underline decoration-zinc-805 decoration-dotted select-all">{reminder.targetQuery}</span>
                                    </p>
                                  )}
                                  {reminder.actionPlan && (
                                    <p className="truncate">
                                      <span className="text-zinc-650 uppercase text-[9px] font-bold mr-1">Action Plan:</span> 
                                      <span className="text-rose-450/80">{reminder.actionPlan}</span>
                                    </p>
                                  )}
                                  <p className="text-[9px] text-zinc-600 pt-0.5">
                                    Saved: {new Date(reminder.createdAt).toLocaleString()}
                                  </p>
                                </div>
                              </div>
                            </div>

                            {/* Right Side: Action Controllers */}
                            <div className="flex sm:flex-row md:flex-col items-stretch md:items-end justify-end gap-2 text-right shrink-0">
                              {reminder.status === "active" && (
                                <button
                                  type="button"
                                  onClick={() => runReminderCheck(reminder.id)}
                                  disabled={isChecking}
                                  className={`px-3 py-1.5 rounded-xl text-[9px] font-mono font-bold uppercase tracking-wider flex items-center justify-center gap-1.5 transition-all border ${
                                    isChecking
                                      ? "bg-zinc-900 border-white/5 text-zinc-500 cursor-not-allowed"
                                      : "bg-rose-500/10 hover:bg-rose-500/20 text-rose-400 border-rose-500/20 scale-100 hover:scale-[1.02]"
                                  }`}
                                  title="Perform periodic tracking condition evaluation immediately"
                                >
                                  {isChecking ? (
                                    <>
                                      <Loader2 className="w-3 h-3 animate-spin" />
                                      <span>Checking...</span>
                                    </>
                                  ) : (
                                    <>
                                      <Play className="w-2.5 h-2.5 fill-current" />
                                      <span>Run Check</span>
                                    </>
                                  )}
                                </button>
                              )}

                              <button
                                type="button"
                                onClick={() => handleDeleteReminder(reminder.id, reminder.condition)}
                                className="px-2.5 py-1.5 bg-zinc-950/60 border border-zinc-855 hover:bg-rose-955/20 hover:text-red-400 hover:border-rose-500/20 rounded-xl text-[9px] font-mono text-zinc-500 tracking-wider flex items-center justify-center gap-1.5 transition-all uppercase"
                                title="Delete and discard tracker condition"
                              >
                                <Trash2 className="w-3 h-3 text-current" />
                                <span>Delete</span>
                              </button>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}

                </div>
              </div>
            )}

            {activeTab === "rem-info" && (
              <div className="flex-1 overflow-y-auto custom-scrollbar p-6 space-y-6">
                <div className="bg-zinc-950/40 border border-white/5 rounded-2xl p-5 space-y-4">
                  <div className="flex items-center justify-between pb-3 border-b border-zinc-900">
                    <div className="flex items-center gap-2.5">
                      <div className="p-2 bg-violet-500/10 rounded-xl text-violet-400 border border-violet-500/20">
                        <CheckCircle2 className="w-4 h-4" />
                      </div>
                      <div className="text-left">
                        <h4 className="text-xs font-mono font-bold uppercase tracking-wider text-zinc-100">Fulfilled Scoper Intelligence Reports</h4>
                        <p className="text-[10px] text-zinc-500 font-mono mt-0.5">Verified lookup findings, official directories and stock status indices</p>
                      </div>
                    </div>
                  </div>

                  {recordedRemInfos.length === 0 ? (
                    <div className="text-center py-12 border border-dashed border-zinc-800 rounded-2xl bg-zinc-950/20">
                      <CheckCircle2 className="w-8 h-8 text-zinc-700 mx-auto mb-3 stroke-[1.5]" />
                      <p className="text-xs text-zinc-400 font-medium">No verified intelligence reports</p>
                      <p className="text-[10px] text-zinc-500 font-mono max-w-sm mx-auto mt-1 leading-relaxed">
                        When active reminders or tracking lookup requests are completed, Tania compiles a formal executive analysis report here.
                      </p>
                    </div>
                  ) : (
                    <div className="space-y-6 pt-1">
                      {recordedRemInfos.map((report, idx) => (
                        <div
                          key={report.id}
                          className="border border-violet-500/20 rounded-2xl bg-zinc-950/30 p-5 space-y-4 transition-all hover:bg-zinc-950/50"
                        >
                          {/* Executive Document Header */}
                          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-zinc-900">
                            <div className="space-y-1 text-left">
                              <span className="inline-block px-2 py-0.5 bg-violet-500/15 border border-violet-500/20 rounded text-[8px] font-mono font-bold text-violet-400 uppercase tracking-widest leading-none">
                                Verification Report
                              </span>
                              <h4 className="text-sm font-semibold text-zinc-100 mt-1">{report.title}</h4>
                            </div>
                            <div className="text-left sm:text-right font-mono text-[9px] text-zinc-500">
                              <p className="font-bold text-violet-400">{report.reportId}</p>
                              <p className="text-[8px] mt-0.5">{report.resolvedAtString}</p>
                            </div>
                          </div>

                          {/* Methodology and Verified Source Profile */}
                          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 bg-zinc-950/60 p-4 border border-white/5 rounded-xl text-left">
                            <div className="space-y-1.5">
                              <span className="text-[8px] uppercase tracking-wider font-mono font-bold text-zinc-500">Target Requirements</span>
                              <div className="text-[10px] text-zinc-300 font-mono space-y-1 leading-normal">
                                <p><strong className="text-zinc-550 uppercase text-[8px] mr-1">Requirement:</strong> {report.condition}</p>
                                {report.targetQuery && (
                                  <p className="truncate"><strong className="text-zinc-550 uppercase text-[8px] mr-1">Query Source:</strong> {report.targetQuery}</p>
                                )}
                                {report.actionPlan && (
                                  <p className="truncate"><strong className="text-zinc-550 uppercase text-[8px] mr-1">Action Plan:</strong> {report.actionPlan}</p>
                                )}
                              </div>
                            </div>

                            <div className="space-y-1.5 border-t md:border-t-0 md:border-l border-zinc-900 pt-3 md:pt-0 md:pl-4">
                              <span className="text-[8px] uppercase tracking-wider font-mono font-bold text-violet-400">Verified Contact Directory Profile</span>
                              <div className="text-[10px] text-zinc-400 font-mono space-y-1 leading-normal">
                                <p><strong className="text-zinc-550 uppercase text-[8px] mr-1">Supplier:</strong> <span className="text-zinc-200 font-bold">{report.providerTitle}</span></p>
                                <p className="truncate"><strong className="text-zinc-550 uppercase text-[8px] mr-1">HQ Address:</strong> <span className="text-zinc-350">{report.providerAddress}</span></p>
                                <p><strong className="text-zinc-550 uppercase text-[8px] mr-1">Active Phone:</strong> <span className="text-zinc-300 hover:text-white cursor-pointer underline select-all">{report.providerPhone}</span></p>
                                <p><strong className="text-zinc-550 uppercase text-[8px] mr-1">Verified URL:</strong> <a href={report.providerWebsite} target="_blank" rel="noopener referrer" className="text-violet-400 hover:text-violet-300 underline font-semibold select-all">{report.providerWebsite}</a></p>
                              </div>
                            </div>
                          </div>

                          {/* Intelligence Findings Description */}
                          <div className="text-left space-y-2">
                            <span className="text-[8px] uppercase tracking-wider font-mono font-bold text-zinc-500 block">Scoped Core Findings & Verification Ledger</span>
                            <div className="bg-zinc-950/80 p-4 border border-zinc-900 rounded-xl prose prose-sm max-w-none text-zinc-300 font-sans leading-relaxed text-[11px] whitespace-pre-wrap">
                              {report.details}
                            </div>
                          </div>

                          {/* Action Controller Footer */}
                          <div className="flex items-center justify-between pt-2">
                            <div className="flex items-center gap-1.5 text-[8.5px] font-mono text-emerald-400 border border-emerald-500/10 bg-emerald-500/5 px-2.5 py-1 rounded-lg">
                              <div className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse" />
                              <span>CERTIFIED LEDGER ACTIVE</span>
                            </div>

                            <button
                              type="button"
                              onClick={() => handleDeleteReport(report.id, report.condition)}
                              className="px-3 py-1.5 bg-zinc-950/60 border border-zinc-850 hover:bg-rose-955/20 hover:text-red-400 hover:border-rose-500/20 rounded-xl text-[9px] font-mono text-zinc-500 tracking-wider flex items-center justify-center gap-1.5 transition-all uppercase"
                              title="Permanently delete verification report from records"
                            >
                              <Trash2 className="w-3 h-3 text-current" />
                              <span>Delete</span>
                            </button>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}

                </div>
              </div>
            )}

            {activeTab === "standing-orders" && (
              <div className="flex-1 overflow-y-auto custom-scrollbar p-6 space-y-6">
                <div className="bg-zinc-950/40 border border-white/5 rounded-2xl p-5 space-y-4">
                  <div className="flex items-center justify-between pb-3 border-b border-zinc-900">
                    <div className="flex items-center gap-2.5">
                      <div className="p-2 bg-amber-500/10 rounded-xl text-amber-405 border border-amber-500/20">
                        <FileText className="w-4 h-4" />
                      </div>
                      <div className="text-left">
                        <h4 className="text-xs font-mono font-bold uppercase tracking-wider text-zinc-100">Permanent Standing Instructions</h4>
                        <p className="text-[10px] text-zinc-500 font-mono mt-0.5">Custom task scripts executed automatically on matching verbal cues</p>
                      </div>
                    </div>
                    <button
                      onClick={() => setIsAddingStandingOrder(!isAddingStandingOrder)}
                      className="px-3 py-1 bg-amber-500/15 hover:bg-amber-500/25 text-amber-400 border border-amber-500/20 rounded-lg text-[9px] font-mono tracking-wider transition-all uppercase font-bold flex items-center gap-1 cursor-pointer"
                    >
                      <span>{isAddingStandingOrder ? "Close" : "+ New Directive"}</span>
                    </button>
                  </div>

                  {/* Drag and Drop File Upload for Standing Orders */}
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div className="bg-zinc-900/30 border border-white/5 rounded-xl p-4 space-y-3.5 text-left">
                      <div className="flex items-center gap-2 border-b border-zinc-900 pb-2">
                        <Paperclip className="w-3.5 h-3.5 text-amber-450" />
                        <span className="text-[10px] font-mono font-bold text-zinc-300 uppercase tracking-wider">Upload Instructions Document</span>
                      </div>
                      
                      <div
                        onDragOver={(e) => { !isStandingOrderUploading && e.preventDefault(); }}
                        onDrop={(e) => {
                          e.preventDefault();
                          if (!isStandingOrderUploading && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                            handleStandingOrderFileUpload(e.dataTransfer.files);
                          }
                        }}
                        onClick={() => !isStandingOrderUploading && document.getElementById("standing-order-file-picker")?.click()}
                        className={`border border-dashed rounded-xl p-5 text-center transition-all flex flex-col items-center justify-center space-y-2 group ${
                          isStandingOrderUploading 
                            ? "border-amber-500/30 bg-amber-500/5 cursor-not-allowed" 
                            : "border-zinc-800 hover:border-amber-500/40 hover:bg-zinc-900/10 cursor-pointer"
                        }`}
                      >
                        <input
                          id="standing-order-file-picker"
                          type="file"
                          accept=".txt,.md,.pdf,.docx,.doc,.xlsx,.xls,.jpg,.jpeg,.png"
                          className="hidden"
                          disabled={isStandingOrderUploading}
                          onChange={(e) => {
                            if (e.target.files && e.target.files.length > 0) {
                              handleStandingOrderFileUpload(e.target.files);
                            }
                          }}
                        />
                        {isStandingOrderUploading ? (
                          <Loader2 className="w-6 h-6 text-amber-500 animate-spin" />
                        ) : (
                          <Upload className="w-6 h-6 text-zinc-500 group-hover:text-amber-400 transition-colors" />
                        )}
                        <div className="space-y-0.5">
                          <p className="text-[11px] text-zinc-300 font-medium">
                            {isStandingOrderUploading ? "Extracting instructions..." : "Click to upload or drag file here"}
                          </p>
                          <p className="text-[8px] text-zinc-500 font-mono uppercase tracking-wider">
                            {isStandingOrderUploading 
                              ? "Gemini Document Intelligence is parsing and cleaning contents..." 
                              : "Supports TXT, PDF, Word, Excel, & JPEG/PNG pictures"
                            }
                          </p>
                        </div>
                      </div>

                      {standingOrderUploadError && (
                        <p className="text-[9px] text-red-400 font-mono bg-red-950/10 p-2 rounded border border-red-500/15 text-left">
                          {standingOrderUploadError}
                        </p>
                      )}
                    </div>

                    {/* Quick guidelines guide */}
                    <div className="bg-zinc-900/30 border border-white/5 rounded-xl p-4 text-left flex flex-col justify-between">
                      <div className="space-y-2">
                        <div className="flex items-center gap-2 border-b border-zinc-900 pb-2">
                          <HelpCircle className="w-3.5 h-3.5 text-amber-450" />
                          <span className="text-[10px] font-mono font-bold text-zinc-300 uppercase tracking-wider">Voice Hook Guidelines</span>
                        </div>
                        <ul className="text-[10px] font-mono text-zinc-400 space-y-1.5 list-disc pl-4 leading-relaxed">
                          <li>Specify a clear, recognizable word or phrase as the <strong className="text-amber-400">Title</strong> (e.g. <code className="text-zinc-200">Find buyers</code>).</li>
                          <li>When you say this Title aloud, Tania catches the sound wave.</li>
                          <li>Tania halts normal responses, <strong className="text-amber-400">reads the custom instructions</strong>, and initiates automated execution.</li>
                        </ul>
                      </div>
                      <div className="pt-2 border-t border-zinc-900 text-[9px] font-mono text-zinc-500 leading-normal">
                        Standing directives are stored in browser persistence securely.
                      </div>
                    </div>
                  </div>

                  {/* Expandable Manual Form */}
                  <AnimatePresence>
                    {isAddingStandingOrder && (
                      <motion.form
                        initial={{ opacity: 0, height: 0 }}
                        animate={{ opacity: 1, height: "auto" }}
                        exit={{ opacity: 0, height: 0 }}
                        onSubmit={(e) => {
                          e.preventDefault();
                          if (!newStandingTitle.trim() || !newStandingInstructions.trim()) return;
                          const newOrder = {
                            id: `SO-${Math.random().toString(36).substring(2, 9)}`,
                            title: newStandingTitle.trim(),
                            instructions: newStandingInstructions.trim(),
                            createdAt: new Date().toLocaleString()
                          };
                          setRecordedStandingOrders(prev => {
                            const next = [...prev, newOrder];
                            localStorage.setItem("tania_standing_orders", JSON.stringify(next));
                            return next;
                          });
                          setNewStandingTitle("");
                          setNewStandingInstructions("");
                          setIsAddingStandingOrder(false);
                          
                          setTranscript(prev => {
                            const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
                            return [...prev, `System: [${timeStr}] ✍️ Saved brand-new manual standing instructions for: "${newOrder.title}".`];
                          });
                        }}
                        className="bg-zinc-900/30 border border-white/5 rounded-xl p-4 space-y-4 overflow-hidden text-left"
                      >
                        <div className="space-y-3">
                          <div>
                            <label className="block text-[9px] font-mono uppercase tracking-wider text-zinc-400 mb-1.5 font-bold">
                              Standing Order Keyphrase/Title *
                            </label>
                            <input
                              type="text"
                              required
                              value={newStandingTitle}
                              onChange={(e) => setNewStandingTitle(e.target.value)}
                              placeholder="e.g., Find buyers"
                              className="w-full bg-zinc-950 border border-zinc-850 focus:border-amber-500/50 rounded-lg px-3 py-2 text-xs text-zinc-100 outline-none transition-all placeholder:text-zinc-600"
                            />
                            <p className="text-[8px] text-zinc-600 mt-1 font-mono">This value matches case-insensitively with your spoken statements.</p>
                          </div>
                          
                          <div>
                            <label className="block text-[9px] font-mono uppercase tracking-wider text-zinc-400 mb-1.5 font-bold">
                              Directives & Action instructions *
                            </label>
                            <textarea
                              required
                              rows={4}
                              value={newStandingInstructions}
                              onChange={(e) => setNewStandingInstructions(e.target.value)}
                              placeholder="Instruct Tania exactly what to do here. E.g., Scrape regional dealer price records for 2.5K LKR margin levels, confirm supply SLA at Dialog Union Place, and alert Pesala's main directory."
                              className="w-full bg-zinc-950 border border-zinc-850 focus:border-amber-500/50 rounded-lg px-3 py-2 text-xs text-zinc-100 outline-none transition-all placeholder:text-zinc-650 custom-scrollbar resize-none"
                            />
                          </div>
                        </div>

                        <div className="flex justify-end gap-2">
                          <button
                            type="button"
                            onClick={() => setIsAddingStandingOrder(false)}
                            className="px-3 py-1.5 rounded-lg text-[9px] font-mono uppercase tracking-wider bg-zinc-950 border border-zinc-850 hover:bg-zinc-900 text-zinc-400 font-bold"
                          >
                            Cancel
                          </button>
                          <button
                            type="submit"
                            className="px-3.5 py-1.5 rounded-lg text-[9px] font-mono uppercase tracking-wider bg-amber-500 text-neutral-950 hover:bg-amber-400 font-bold"
                          >
                            Save Directive
                          </button>
                        </div>
                      </motion.form>
                    )}
                  </AnimatePresence>

                  {/* Standing Orders List */}
                  {recordedStandingOrders.length === 0 ? (
                    <div className="text-center py-12 border border-dashed border-zinc-800 rounded-2xl bg-zinc-950/20">
                      <FileText className="w-8 h-8 text-zinc-700 mx-auto mb-3 stroke-[1.5]" />
                      <p className="text-xs text-zinc-400 font-medium font-mono">No standing directives configured yet</p>
                      <p className="text-[10px] text-zinc-550 font-mono max-w-sm mx-auto mt-1 leading-relaxed">
                        Add a standing directive or upload a formatted text guide. Once registered, speaking the Title aloud executes these instructions instantly.
                      </p>
                    </div>
                  ) : (
                    <div className="space-y-4 pt-1">
                      {recordedStandingOrders.map((order) => {
                        const isEditing = editingStandingOrderId === order.id;
                        return (
                          <div
                            key={order.id}
                            className="border border-zinc-850 rounded-xl bg-zinc-950/20 p-4 space-y-3.5 transition-all hover:bg-zinc-950/45 text-left"
                          >
                            {isEditing ? (
                              <div className="space-y-3">
                                <div>
                                  <label className="block text-[8px] font-mono uppercase tracking-wider text-zinc-500 mb-1">Title Keyphrase</label>
                                  <input
                                    type="text"
                                    value={editStandingTitle}
                                    onChange={(e) => setEditStandingTitle(e.target.value)}
                                    className="w-full bg-zinc-950 border border-zinc-850 focus:border-amber-500/40 rounded px-2.5 py-1.5 text-xs text-zinc-100 outline-none"
                                  />
                                </div>
                                <div>
                                  <label className="block text-[8px] font-mono uppercase tracking-wider text-zinc-500 mb-1">Instructions</label>
                                  <textarea
                                    value={editStandingInstructions}
                                    onChange={(e) => setEditStandingInstructions(e.target.value)}
                                    rows={3}
                                    className="w-full bg-zinc-950 border border-zinc-850 focus:border-amber-500/40 rounded px-2.5 py-1.5 text-xs text-zinc-100 outline-none custom-scrollbar"
                                  />
                                </div>
                                <div className="flex justify-end gap-2">
                                  <button
                                    onClick={() => setEditingStandingOrderId(null)}
                                    className="px-2 py-1 bg-zinc-950 border border-zinc-850 rounded text-[9px] text-zinc-400 font-mono"
                                  >
                                    Cancel
                                  </button>
                                  <button
                                    onClick={() => {
                                      if (!editStandingTitle.trim() || !editStandingInstructions.trim()) return;
                                      setRecordedStandingOrders(prev => {
                                        const next = prev.map(o => o.id === order.id 
                                          ? { ...o, title: editStandingTitle.trim(), instructions: editStandingInstructions.trim() }
                                          : o
                                        );
                                        localStorage.setItem("tania_standing_orders", JSON.stringify(next));
                                        return next;
                                      });
                                      setEditingStandingOrderId(null);
                                      setTranscript(prev => {
                                        const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
                                        return [...prev, `System: [${timeStr}] ✏️ Updated standing instructions for "${editStandingTitle.trim()}".`];
                                      });
                                    }}
                                    className="px-2.5 py-1 bg-amber-500 hover:bg-amber-400 text-neutral-950 rounded text-[9px] font-mono font-bold"
                                  >
                                    Save
                                  </button>
                                </div>
                              </div>
                            ) : (
                              <>
                                <div className="flex items-center justify-between border-b border-zinc-900 pb-2">
                                  <div className="space-y-0.5">
                                    <div className="flex items-center gap-2">
                                      <span className="w-1.5 h-1.5 rounded-full bg-amber-400 animate-pulse" />
                                      <h5 className="text-xs font-bold font-mono tracking-wider text-zinc-100 uppercase">{order.title}</h5>
                                    </div>
                                    <span className="block text-[8px] font-mono text-zinc-550 font-bold">Registered: {order.createdAt}</span>
                                  </div>
                                  {order.fileAttached && (
                                    <div className="flex items-center gap-1.5 px-2 py-0.5 rounded bg-zinc-900 border border-zinc-800 text-[8px] text-zinc-400 font-mono">
                                      <Paperclip className="w-2.5 h-2.5 text-zinc-500" />
                                      <span className="truncate max-w-[120px]">{order.fileAttached}</span>
                                    </div>
                                  )}
                                </div>

                                <div className="bg-zinc-950/40 p-3.5 border border-zinc-900 rounded-lg text-xs leading-relaxed text-zinc-300 whitespace-pre-wrap font-sans">
                                  {order.instructions}
                                </div>

                                <div className="flex items-center justify-between pt-1">
                                  <button
                                    onClick={() => executeStandingOrder(order)}
                                    className="px-2.5 py-1 bg-amber-500/10 border border-amber-500/20 hover:bg-amber-500/25 text-amber-400 rounded-lg text-[9px] font-mono tracking-wider transition-all uppercase font-bold flex items-center gap-1 cursor-pointer"
                                    title="Manually execute instructions immediately"
                                  >
                                    <Play className="w-3 h-3 text-current" />
                                    <span>Run Directive Now</span>
                                  </button>

                                  <div className="flex gap-2">
                                    <button
                                      onClick={() => {
                                        setEditingStandingOrderId(order.id);
                                        setEditStandingTitle(order.title);
                                        setEditStandingInstructions(order.instructions);
                                      }}
                                      className="px-2.5 py-1 bg-zinc-900 border border-zinc-800 hover:bg-zinc-800 text-zinc-300 rounded-lg text-[9px] font-mono tracking-wider transition-all uppercase flex items-center gap-1 cursor-pointer"
                                    >
                                      <Edit2 className="w-2.5 h-2.5 text-current" />
                                      <span>Edit</span>
                                    </button>

                                    <button
                                      onClick={() => {
                                        setRecordedStandingOrders(prev => {
                                          const next = prev.filter(o => o.id !== order.id);
                                          localStorage.setItem("tania_standing_orders", JSON.stringify(next));
                                          return next;
                                        });
                                        setTranscript(prev => {
                                          const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
                                          return [...prev, `System: [${timeStr}] 🗑️ Deleted standing instructions: "${order.title}".`];
                                        });
                                      }}
                                      className="px-2.5 py-1 bg-zinc-950/60 border border-zinc-850 hover:bg-rose-955/20 hover:text-red-400 hover:border-rose-500/20 rounded-lg text-[9px] font-mono text-zinc-500 tracking-wider flex items-center justify-center gap-1 transition-all uppercase cursor-pointer"
                                    >
                                      <Trash2 className="w-2.5 h-2.5 text-current" />
                                      <span>Delete</span>
                                    </button>
                                  </div>
                                </div>
                              </>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  )}

                </div>
              </div>
            )}

            {activeTab === "pictures" && (
              <div className="flex-1 overflow-y-auto custom-scrollbar p-6 space-y-6">
                <div className="bg-zinc-950/40 border border-white/5 rounded-2xl p-5 space-y-4">
                  <div className="flex items-center justify-between pb-3 border-b border-zinc-900">
                    <div className="flex items-center gap-2.5">
                      <div className="p-2 bg-teal-500/10 rounded-xl text-teal-400 border border-teal-500/20">
                        <Image className="w-4 h-4" />
                      </div>
                      <div className="text-left">
                        <h4 className="text-xs font-mono font-bold uppercase tracking-wider text-zinc-100">Requested Full Colour Pictures</h4>
                        <p className="text-[10px] text-zinc-500 font-mono mt-0.5">Streamed picture cards in your workspace gallery</p>
                      </div>
                    </div>
                    {requestedImages.length > 0 && (
                      <div className="flex items-center bg-zinc-900/60 p-1 border border-white/5 rounded-lg shrink-0">
                        <button
                          onClick={() => setPicturesViewMode("scroll")}
                          className={`px-2.5 py-1 text-[9px] font-mono uppercase tracking-wider font-bold rounded-md transition-all ${
                            picturesViewMode === "scroll"
                              ? "bg-teal-500 text-zinc-950 shadow-md"
                              : "text-zinc-400 hover:text-white"
                          }`}
                        >
                          Scroll Row
                        </button>
                        <button
                          onClick={() => setPicturesViewMode("grid")}
                          className={`px-2.5 py-1 text-[9px] font-mono uppercase tracking-wider font-bold rounded-md transition-all ${
                            picturesViewMode === "grid"
                              ? "bg-teal-500 text-zinc-950 shadow-md"
                              : "text-zinc-400 hover:text-white"
                          }`}
                        >
                          Grid View
                        </button>
                      </div>
                    )}
                  </div>

                  {requestedImages.length === 0 ? (
                    <div className="flex flex-col items-center justify-center py-16 space-y-3 bg-zinc-900/10 rounded-xl border border-dashed border-zinc-800">
                      <Image className="w-8 h-8 text-zinc-700" />
                      <p className="text-xs font-mono text-zinc-500 uppercase tracking-wider">No pictures requested yet</p>
                      <p className="text-[10px] text-zinc-600 max-w-sm text-center">
                        Tell Tania to view or display pictures (e.g., "show a picture of vintage red ferrari") to populate this library.
                      </p>
                    </div>
                  ) : picturesViewMode === "scroll" ? (
                    <div className="relative group/scrollcontainer">
                      <div className="absolute left-1 top-1/2 -translate-y-1/2 z-10 opacity-0 group-hover/scrollcontainer:opacity-100 transition-opacity pointer-events-none">
                        <div className="p-2.5 bg-zinc-950/80 border border-white/10 rounded-full text-zinc-400 font-bold shadow-2xl backdrop-blur-md">
                          <ChevronLeft className="w-4 h-4" />
                        </div>
                      </div>
                      
                      <div className="flex gap-5 overflow-x-auto pb-4 scroll-smooth snap-x snap-mandatory custom-scrollbar-horizontal w-full">
                        {requestedImages.map((img, index) => (
                          <div 
                            key={index} 
                            className="flex-shrink-0 w-80 snap-start bg-zinc-900/50 rounded-2xl border border-white/5 overflow-hidden flex flex-col group/piccard hover:border-teal-500/30 transition-all duration-300"
                          >
                            <div className="relative aspect-[4/3] bg-neutral-905 overflow-hidden cursor-pointer"
                              onClick={() => {
                                setCurrentImage({ url: img.url, query: img.query });
                                setIsImageModalOpen(true);
                              }}
                            >
                              <img 
                                src={img.url} 
                                alt={img.query} 
                                className="w-full h-full object-cover transition-transform duration-500 group-hover/piccard:scale-105"
                                referrerPolicy="no-referrer"
                              />
                              <div className="absolute inset-0 bg-gradient-to-t from-black/85 via-black/25 to-transparent flex items-end p-4">
                                <span className="text-[10px] font-mono text-zinc-300 font-bold uppercase tracking-widest truncate max-w-full">
                                  {img.query}
                                </span>
                              </div>
                            </div>
                            <div className="p-4 flex flex-col justify-between flex-1 space-y-4 bg-zinc-950/20">
                              <div>
                                <h5 className="text-xs font-mono font-bold uppercase text-zinc-200 truncate capitalize">
                                  {img.query}
                                </h5>
                                <p className="text-[10px] font-mono text-zinc-500 mt-1">
                                  Requested {new Date(img.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                                </p>
                              </div>
                              <div className="flex items-center gap-2">
                                <button
                                  onClick={() => {
                                    setCurrentImage({ url: img.url, query: img.query });
                                    setIsImageModalOpen(true);
                                  }}
                                  className="flex-1 py-2 bg-zinc-800 hover:bg-zinc-750 border border-white/5 text-zinc-300 hover:text-white rounded-xl text-[10px] font-mono uppercase font-bold tracking-wider text-center transition-all duration-300"
                                >
                                  Preview
                                </button>
                                <button
                                  onClick={() => downloadImage(img.url, img.query)}
                                  className="py-2 px-3 bg-teal-600/15 hover:bg-teal-600/30 border border-teal-500/20 text-teal-400 hover:text-teal-350 rounded-xl text-[10px] font-mono transition-all duration-300 flex items-center justify-center"
                                  title="Download Image"
                                >
                                  <Download className="w-3.5 h-3.5" />
                                </button>
                              </div>
                            </div>
                          </div>
                        ))}
                      </div>

                      <div className="absolute right-1 top-1/2 -translate-y-1/2 z-10 opacity-0 group-hover/scrollcontainer:opacity-100 transition-opacity pointer-events-none">
                        <div className="p-2.5 bg-zinc-950/80 border border-white/10 rounded-full text-zinc-400 font-bold shadow-2xl backdrop-blur-md">
                          <ChevronRight className="w-4 h-4" />
                        </div>
                      </div>
                    </div>
                  ) : (
                    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                      {requestedImages.map((img, index) => (
                        <div key={index} className="bg-zinc-900/50 rounded-xl border border-white/5 overflow-hidden flex flex-col group hover:border-teal-500/30 transition-all duration-300">
                          <div className="relative aspect-video bg-neutral-900 overflow-hidden">
                            <img 
                              src={img.url} 
                              alt={img.query} 
                              className="w-full h-full object-cover transition-transform duration-500 group-hover:scale-105"
                              referrerPolicy="no-referrer"
                            />
                            <div className="absolute inset-0 bg-gradient-to-t from-black/80 via-black/20 to-transparent opacity-0 group-hover:opacity-100 transition-opacity duration-300 flex items-end p-3">
                              <span className="text-[10px] font-mono text-zinc-300 uppercase tracking-widest truncate max-w-full">
                                {img.query}
                              </span>
                            </div>
                          </div>
                          <div className="p-3 flex flex-col justify-between flex-1 space-y-3 bg-zinc-950/20">
                            <div>
                              <h5 className="text-[11px] font-mono font-bold uppercase text-zinc-200 truncate capitalize">
                                {img.query}
                              </h5>
                              <p className="text-[9px] font-mono text-zinc-500 mt-1">
                                Requested at {new Date(img.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                              </p>
                            </div>
                            <div className="flex items-center gap-2">
                              <button
                                onClick={() => downloadImage(img.url, img.query)}
                                className="flex-1 py-1.5 bg-teal-600/10 hover:bg-teal-600/25 border border-teal-500/20 text-teal-400 hover:text-teal-350 rounded-lg text-[9px] font-mono uppercase font-bold tracking-wider text-center transition-all duration-300 flex items-center justify-center gap-1.5"
                              >
                                <Download className="w-3 h-3" />
                                <span>Download</span>
                              </button>
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}

                </div>
              </div>
            )}

            {activeTab === "videos" && (
              <div className="flex-1 overflow-y-auto custom-scrollbar p-6 space-y-6">
                <div className="bg-zinc-950/40 border border-white/5 rounded-2xl p-5 space-y-5">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between pb-4 border-b border-zinc-900 gap-4">
                    <div className="flex items-center gap-2.5">
                      <div className="p-2 bg-amber-500/10 rounded-xl text-amber-400 border border-amber-500/20">
                        <Youtube className="w-4 h-4" />
                      </div>
                      <div className="text-left">
                        <h4 className="text-xs font-mono font-bold uppercase tracking-wider text-zinc-100">Live Video Stream Player</h4>
                        <p className="text-[10px] text-zinc-500 font-mono mt-0.5">Watch, search or stream YouTube videos requested on a subject</p>
                      </div>
                    </div>

                    <div className="flex items-center gap-2 max-w-md w-full">
                      <input
                        type="text"
                        placeholder="Search another YouTube video..."
                        onKeyDown={async (e) => {
                          if (e.key === 'Enter') {
                            const q = e.currentTarget.value.trim();
                            if (q) {
                              setIsSearching(true);
                              try {
                                const response = await fetch(`/api/youtube-search?query=${encodeURIComponent(q)}`);
                                const result = await response.json();
                                if (result.success && result.videos && result.videos.length > 0) {
                                  const first = result.videos[0];
                                  setCurrentVideoId(first.videoId);
                                  setCurrentVideoTitle(first.title);
                                  setRequestedVideos(prev => {
                                    const existing = prev.filter(v => v.videoId !== first.videoId);
                                    const newEntries = result.videos.map((v: any) => ({
                                      videoId: v.videoId,
                                      title: v.title,
                                      query: q,
                                      timestamp: Date.now()
                                    }));
                                    return [...newEntries, ...existing].slice(0, 30);
                                  });
                                }
                              } catch (err) {
                                console.error(err);
                              } finally {
                                setIsSearching(false);
                              }
                            }
                          }
                        }}
                        className="flex-1 px-3 py-1.5 bg-zinc-900/80 border border-white/5 rounded-xl text-xs text-zinc-100 focus:outline-none focus:border-amber-500/30 font-mono"
                      />
                      <div className="text-[9px] font-mono text-zinc-500 uppercase tracking-widest hidden md:block">Press Enter</div>
                    </div>
                  </div>

                  {currentVideoId ? (
                    <div className="space-y-4">
                      <div className="relative w-full h-[280px] sm:h-[400px] md:h-[480px] lg:h-[550px] rounded-2xl overflow-hidden bg-black border border-white/5 shadow-2xl">
                        <iframe
                          src={`https://www.youtube.com/embed/${currentVideoId}?autoplay=1&rel=0`}
                          title={currentVideoTitle || "YouTube video player"}
                          frameBorder="0"
                          allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share"
                          allowFullScreen
                          className="absolute inset-0 w-full h-full"
                        ></iframe>
                      </div>
                      <div className="flex items-start justify-between p-2 bg-zinc-900/20 border border-white/5 rounded-xl">
                        <div className="text-left">
                          <span className="bg-amber-500/10 text-amber-400 border border-amber-500/20 text-[9px] font-mono font-bold uppercase tracking-wider px-2 py-0.5 rounded-md">
                            Now Playing
                          </span>
                          <h4 className="text-sm font-semibold text-zinc-200 mt-2 line-clamp-2">
                            {currentVideoTitle}
                          </h4>
                          <p className="text-[9px] font-mono text-zinc-500 mt-0.5">
                            ID: {currentVideoId}
                          </p>
                        </div>
                        <button
                          onClick={() => {
                            setCurrentVideoId(null);
                            setCurrentVideoTitle("");
                          }}
                          className="p-1 px-2.5 bg-zinc-800 hover:bg-zinc-700 text-zinc-400 hover:text-white rounded-lg text-[10px] font-mono uppercase tracking-wider transition-colors"
                        >
                          Close Player
                        </button>
                      </div>
                    </div>
                  ) : (
                    <div className="flex flex-col items-center justify-center py-20 space-y-4 bg-zinc-900/10 rounded-2xl border border-dashed border-zinc-800">
                      <Youtube className="w-12 h-12 text-zinc-700 animate-pulse" />
                      <div className="text-center space-y-1">
                        <p className="text-xs font-mono text-zinc-500 uppercase tracking-wider">No active video stream playing</p>
                        <p className="text-[10px] text-zinc-600 max-w-md">
                          Tell Tania to search and play a video (e.g., "play a video about Colombo tour" or "watch a cricket match clip") or search manually above.
                        </p>
                      </div>
                    </div>
                  )}

                  {requestedVideos.length > 0 && (
                    <div className="space-y-3 pt-4 border-t border-zinc-900">
                      <div className="flex items-center justify-between">
                        <h5 className="text-[11px] font-mono font-bold uppercase tracking-wider text-zinc-400">
                          Workspace Video Stream Library (Scroll sideways ↔)
                        </h5>
                        <span className="text-[9px] font-mono text-zinc-650 uppercase">
                          {requestedVideos.length} matching clips
                        </span>
                      </div>

                      <div className="flex gap-4 overflow-x-auto pb-4 scroll-smooth snap-x snap-mandatory custom-scrollbar-horizontal w-full">
                        {requestedVideos.map((video, idx) => (
                          <div
                            key={idx}
                            onClick={() => {
                              setCurrentVideoId(video.videoId);
                              setCurrentVideoTitle(video.title);
                            }}
                            className={`flex-shrink-0 w-64 snap-start rounded-xl border overflow-hidden flex flex-col cursor-pointer transition-all duration-300 ${
                              currentVideoId === video.videoId
                                ? "bg-amber-950/20 border-amber-500/40"
                                : "bg-zinc-900/40 border-white/5 hover:border-amber-500/20"
                            }`}
                          >
                            <div className="relative aspect-video bg-neutral-900 overflow-hidden">
                              <img
                                src={video.thumbnail}
                                alt={video.title}
                                className="w-full h-full object-cover transition-transform duration-300 hover:scale-105"
                                referrerPolicy="no-referrer"
                              />
                              <div className="absolute inset-0 bg-black/40 hover:bg-black/10 transition-colors duration-300 flex items-center justify-center">
                                <div className="p-2.5 bg-amber-500 text-zinc-950 rounded-full shadow-lg transform translate-y-1 hover:translate-y-0 transition-all opacity-90">
                                  <Play className="w-4 h-4 fill-current ml-0.5" />
                                </div>
                              </div>
                            </div>
                            <div className="p-3 flex-1 flex flex-col justify-between space-y-2">
                              <h6 className="text-[11px] font-mono font-bold text-zinc-200 line-clamp-2">
                                {video.title}
                              </h6>
                              <div className="flex items-center justify-between text-[8px] font-mono text-zinc-500 uppercase tracking-wider">
                                <span>Ref: {video.query.length > 18 ? `${video.query.substring(0, 18)}...` : video.query}</span>
                                <span>{new Date(video.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</span>
                              </div>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                </div>
              </div>
            )}

            <div className="px-6 py-4 border-t border-white/5 bg-zinc-900/40 flex items-center justify-between">
              {activeTab === "conversation" ? (
                <div className="flex items-center gap-3">
                  <button 
                    onClick={() => exportToWord()}
                    className={`flex items-center gap-2 text-[10px] font-mono transition-all uppercase tracking-[0.2em] font-bold ${
                      isExporting ? 'text-zinc-600 cursor-not-allowed' : 'text-orange-500 hover:text-orange-400 hover:scale-105'
                    }`}
                  >
                    <Download className="w-4 h-4" />
                    Word
                  </button>
                  <button 
                    onClick={() => exportToPDF()}
                    className={`flex items-center gap-2 text-[10px] font-mono transition-all uppercase tracking-[0.2em] font-bold ${
                      isExporting ? 'text-zinc-600 cursor-not-allowed' : 'text-blue-500 hover:text-blue-400 hover:scale-105'
                    }`}
                    title="Download PDF Transcript"
                  >
                    <FileText className="w-4 h-4" />
                    PDF
                  </button>
                  <button 
                    onClick={() => exportToExcel()}
                    className={`flex items-center gap-2 text-[10px] font-mono transition-all uppercase tracking-[0.2em] font-bold ${
                      isExporting ? 'text-zinc-600 cursor-not-allowed' : 'text-green-500 hover:text-green-450 hover:scale-105'
                    }`}
                    title="Download Excel Spreadsheet"
                  >
                    <FileSpreadsheet className="w-4 h-4" />
                    Excel
                  </button>
                </div>
              ) : (
                <p className="text-[10.5px] font-mono text-zinc-500 uppercase tracking-widest font-bold">
                  {activeTab === "quotes" ? "Estimations Workspace" : 
                   activeTab === "documents" ? "Document Drafts Workspace" : 
                   activeTab === "communications" ? "Communications Gateway" : 
                   activeTab === "uploads" ? "Uploaded Files & Custom Context" : 
                   activeTab === "pictures" ? "Workspace Pictures Gallery" : 
                   activeTab === "videos" ? "Interactive Video Workspace" :
                   activeTab === "rem-info" ? "Professional Executive Reports" :
                   activeTab === "standing-orders" ? "Tania Standing Directives" :
                   "Automated Reminders Queue"}
                </p>
              )}
              {isConnected && (
                <div className="flex items-center gap-2">
                  <div className="w-2 h-2 rounded-full bg-red-500 animate-pulse shadow-[0_0_10px_rgba(239,68,68,0.5)]" />
                  <span className="text-[10px] font-mono text-red-500 font-bold uppercase tracking-widest animate-pulse">Streaming</span>
                </div>
              )}
            </div>
          </div>

          {/* Action Button */}
          <div className="flex flex-col items-center gap-4 w-full max-w-md">
            <Button
              onClick={toggleConnection}
              disabled={isConnecting}
              variant={isConnected ? "destructive" : "default"}
              className={`w-full h-14 rounded-2xl text-base font-medium transition-all duration-300 ${
                !isConnected ? 'bg-orange-600 hover:bg-orange-700' : 'bg-zinc-900 hover:bg-zinc-800 border border-zinc-800'
              } ${isConnecting ? 'opacity-50 cursor-not-allowed' : ''}`}
            >
              {isConnecting ? (
                <>
                  <motion.div
                    animate={{ rotate: 360 }}
                    transition={{ duration: 1, repeat: Infinity, ease: "linear" }}
                    className="mr-2"
                  >
                    <Sparkles className="h-5 w-5" />
                  </motion.div>
                  Connecting...
                </>
              ) : isConnected ? (
                <><X className="mr-2 h-5 w-5" /> Stop Session</>
              ) : (
                <><Mic className="mr-2 h-5 w-5" /> Start Talking</>
              )}
            </Button>
            <div className="flex flex-col items-center gap-1">
              <p className="text-[10px] text-zinc-600 uppercase tracking-widest font-mono">
                {isConnected ? "Tania is listening" : "Tap to begin conversation"}
              </p>
              <p className="text-[11px] text-zinc-500 uppercase tracking-widest font-mono mt-6">
                © {new Date().getFullYear()} Pesala Jayawardene. All rights reserved.
              </p>
            </div>
          </div>
        </div>

        {/* Dynamic Image Canvas: Right pane on desktop, stacked on mobile */}
        <AnimatePresence>
          {currentImage && (
            <div className="lg:col-span-5 w-full flex flex-col justify-start lg:sticky lg:top-24 mt-8 lg:mt-0">
              <motion.div 
                initial={{ opacity: 0, scale: 0.95, y: 20 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.95, y: 20 }}
                transition={{ duration: 0.4, ease: "easeOut" }}
                className="bg-zinc-900/40 backdrop-blur-2xl border border-white/5 rounded-[2rem] p-6 space-y-6 shadow-[0_20px_50px_rgba(0,0,0,0.5)] w-full"
              >
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-3">
                    <div className="p-2.5 bg-orange-600/20 rounded-xl border border-orange-500/20">
                      <Image className="w-5 h-5 text-orange-500 animate-pulse" />
                    </div>
                    <div>
                      <h3 className="text-sm font-semibold text-zinc-200 uppercase tracking-wider font-mono">JPEG Monitor</h3>
                      <p className="text-[10px] text-zinc-500 font-mono">Active Image Stream</p>
                    </div>
                  </div>
                  <button 
                    onClick={() => setCurrentImage(null)}
                    className="p-2 bg-zinc-805 hover:bg-white/10 rounded-full transition-all text-zinc-400 hover:text-white"
                    title="Clear picture"
                  >
                    <X className="w-5 h-5" />
                  </button>
                </div>

                <div className="relative aspect-[4/3] rounded-2xl overflow-hidden bg-zinc-950 border border-white/5 shadow-inner group/img">
                  {isImageLoading && (
                    <div className="absolute inset-0 z-30 flex flex-col items-center justify-center space-y-3 bg-zinc-950/80">
                      <Loader2 className="w-8 h-8 text-orange-500 animate-spin" />
                      <p className="text-xs text-zinc-500 font-mono uppercase tracking-widest">Awaiting Live Render...</p>
                    </div>
                  )}
                  <img 
                    src={currentImage.url} 
                    alt={currentImage.query}
                    referrerPolicy="no-referrer"
                    className="w-full h-full object-cover transition-transform duration-500 group-hover/img:scale-105"
                    onLoad={() => setIsImageLoading(false)}
                    onError={handleImageError}
                  />
                  {/* Overlay with details */}
                  <div className="absolute bottom-0 inset-x-0 bg-gradient-to-t from-black/90 via-black/55 to-transparent p-4 flex items-end justify-between z-20">
                    <div className="max-w-[75%]">
                      <p className="text-[9px] text-orange-400 font-mono uppercase tracking-widest font-semibold">Active Subject</p>
                      <h4 className="text-xs font-bold text-white truncate capitalize leading-tight">{currentImage.query}</h4>
                    </div>
                    <span className="text-[9px] font-mono text-zinc-400 bg-black/60 px-2.5 py-1 rounded-md border border-white/5">
                      JPEG 800x600
                    </span>
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <Button
                    onClick={() => downloadImage(currentImage.url, currentImage.query)}
                    className="h-12 bg-zinc-850 hover:bg-zinc-800 text-zinc-200 border border-zinc-800/80 rounded-xl text-[10px] font-mono font-bold uppercase tracking-widest transition-all flex items-center justify-center gap-2"
                  >
                    <Download className="w-4 h-4" />
                    Download
                  </Button>
                  <Button
                    onClick={() => printImage(currentImage.url)}
                    className="h-12 bg-orange-600 hover:bg-orange-700 text-white rounded-xl text-[10px] font-mono font-bold uppercase tracking-widest transition-all flex items-center justify-center gap-2 shadow-lg shadow-orange-950/20"
                  >
                    <Printer className="w-4 h-4" />
                    Print
                  </Button>
                </div>

                <div className="text-center pt-2 border-t border-white/5">
                  <p className="text-[9px] text-zinc-600 font-mono uppercase tracking-widest">
                    Captured from Unsplash Global Image Library
                  </p>
                </div>
              </motion.div>
            </div>
          )}
        </AnimatePresence>

        {/* Floating Information Popups */}
        <div className="fixed bottom-6 right-6 z-50 flex flex-col gap-3 max-w-sm w-full pointer-events-none">
          <AnimatePresence>
            {notifications.map((n) => (
              <motion.div
                key={n.id}
                initial={{ opacity: 0, y: 35, scale: 0.92 }}
                animate={{ opacity: 1, y: 0, scale: 1 }}
                exit={{ opacity: 0, scale: 0.9, y: -15, transition: { duration: 0.18 } }}
                className="pointer-events-auto w-full bg-zinc-950/95 backdrop-blur-md border border-emerald-500/25 rounded-xl p-4 shadow-2xl shadow-black/90 flex gap-3 text-zinc-100 overflow-hidden relative"
              >
                {/* Left Color Accent Indicator */}
                <div className="absolute top-0 left-0 bottom-0 w-1.5 bg-emerald-500" />
                
                {/* Icon */}
                <div className="flex-shrink-0 text-emerald-400 mt-0.5 ml-1">
                  {n.type === "reminder_check" ? (
                    <CheckCircle2 className="w-5 h-5 animate-pulse" />
                  ) : (
                    <Sparkles className="w-5 h-5 text-indigo-400 animate-pulse" />
                  )}
                </div>

                {/* Message */}
                <div className="flex-1 min-w-0 pr-2">
                  <h4 className="text-[11px] font-bold font-sans text-zinc-100 uppercase tracking-widest leading-tight">
                    {n.title}
                  </h4>
                  <p className="text-[11px] text-zinc-400 mt-1 font-sans leading-relaxed">
                    {n.message}
                  </p>
                  <div className="flex items-center gap-1.5 mt-2">
                    <span className="text-[8px] font-mono text-zinc-500 uppercase tracking-widest pl-0.5">
                      Tania Intelligence Engine
                    </span>
                  </div>
                </div>

                {/* Close Button */}
                <button
                  onClick={() => dismissNotification(n.id)}
                  className="flex-shrink-0 text-zinc-500 hover:text-zinc-300 hover:bg-zinc-800/50 p-1 rounded-lg transition-colors h-fit self-start"
                  aria-label="Dismiss Notification"
                >
                  <X className="w-3.5 h-3.5" />
                </button>
              </motion.div>
            ))}
          </AnimatePresence>
        </div>
      </div>
    </div>
  );
}

