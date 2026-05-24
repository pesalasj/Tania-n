import React, { useState, useEffect, useRef } from "react";
import { LiveAPI } from "@/src/lib/live-api";
import { Button } from "@/components/ui/button";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Mic, MicOff, MessageSquare, Sparkles, X, Download, Lock, User, LogOut, History, Printer, Save, FileText } from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
import { Document, Packer, Paragraph, TextRun, Header, Footer, AlignmentType, PageNumber } from "docx";
import { saveAs } from "file-saver";
import { jsPDF } from "jspdf";
import html2canvas from "html2canvas";
import { db, auth } from "@/src/lib/firebase";
import { collection, addDoc, query, where, orderBy, onSnapshot, serverTimestamp, Timestamp, getDocFromServer, doc, deleteDoc, writeBatch, getDocs } from "firebase/firestore";
import { signInAnonymously, onAuthStateChanged } from "firebase/auth";

declare global {
  interface Window {
    aistudio?: {
      hasSelectedApiKey: () => Promise<boolean>;
      openSelectKey: () => Promise<void>;
    };
  }
}

const SYSTEM_INSTRUCTION = `
You are Tania, a warm, wise, and highly multilingual AI assistant. 
You are fluent in over 15 languages, including Sinhala, Tamil, English, French, German, Spanish, Italian, Japanese, Mandarin, Korean, Arabic, Russian, Hindi, Portuguese, and Dutch.
You speak with a gentle, hospitable tone and use traditional greetings like "Ayubowan" when appropriate.
Your personality is knowledgeable about Sri Lankan culture but also globally aware and versatile in communication.
You are here to talk with the user by voice and provide conversational text in the transcript.
CRITICAL: When the session starts, you MUST greet the user immediately without waiting for them to speak.
CRITICAL: You MUST provide a text transcription for EVERYTHING you say. Never speak without also providing the corresponding text in the model turn.
The user you are talking to is Pesala Jayawardene. Address him as "Pesala" frequently and warmly.
Your name is Tania. Never refer to yourself as anything else.
You are currently running on version v112.01 of the AI Assistant core.
You have access to Google Search to answer any questions accurately in any of your supported languages.
You are equipped with tools to record, save, and export conversations. 
Use 'export_transcript' to generate a Word document or 'export_pdf' for a PDF report.
Use 'save_to_cloud' to persist the conversation in the user's history database.
When Pesala Jayawardene asks to save, remember, or record the session, use the 'save_to_cloud' tool.
`;

const TOOLS = [
  { googleSearch: {} },
  {
    functionDeclarations: [
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
        name: "save_to_cloud",
        description: "Saves the current conversation to the secure cloud history database.",
        parameters: { type: "object", properties: {} }
      }
    ]
  }
];

export default function App() {
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
  
  const [lastSavedIndex, setLastSavedIndex] = useState(-1);
  const [isAutoSaving, setIsAutoSaving] = useState(false);
  
  const liveApiRef = useRef<LiveAPI | null>(null);
  const hasSavedSessionRef = useRef(false);
  const scrollRef = useRef<HTMLDivElement>(null);
  const messagesEndRef = useRef<HTMLDivElement>(null);

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

    setTranscript((prev) => {
      const finalLine = parts.length > 1 ? trimmedLine : `System: ${trimmedLine}`;
      console.log("Transcript Processing:", finalLine);
      
      const currentTranscript = [...prev];
      if (currentTranscript.length === 0) {
        transcriptRef.current = [finalLine];
        return [finalLine];
      }
      
      const lastLine = currentTranscript[currentTranscript.length - 1];
      
      // 1. Normalization Check (ignore minor punctuation/spacing differences)
      const normalize = (s: string) => s.toLowerCase().replace(/[.,!?;:]/g, "").trim();
      if (normalize(lastLine) === normalize(finalLine)) {
        console.log("Skipping duplicate normalized line");
        return prev;
      }

      // Special handling for System messages - don't merge if they are different
      if (role === "System") {
        const newTranscript = [...currentTranscript, finalLine];
        transcriptRef.current = newTranscript;
        return newTranscript;
      }

      const lastParts = lastLine.split(': ');
      const lastRole = lastParts[0];
      const lastContent = lastLine.slice(lastRole.length + 2).trim();

      if (lastRole === role) {
        // Merge the speaker texts intelligently
        const normLast = lastContent.toLowerCase().trim();
        const normNew = content.toLowerCase().trim();
        
        let mergedContent = lastContent;
        let didMerge = false;
        
        if (normLast.includes(normNew)) {
          // If the last content already contains the new content, skip/do nothing (it's redundant or old)
          return prev;
        } else if (normNew.startsWith(normLast)) {
          // Cumulative update (the new string is a complete superset)
          mergedContent = content;
          didMerge = true;
        } else {
          // Check for word-level overlapping suffix/prefix to stitch speech fragments together
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
            // No direct overlap but same speaker, join them with a space
            const needsSpace = !lastContent.endsWith(" ") && !content.startsWith(" ");
            mergedContent = lastContent + (needsSpace ? " " : "") + content;
            didMerge = true;
          }
        }
        
        if (didMerge) {
          const mergedLine = `${role}: ${mergedContent}`;
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
      } catch (error) {
        if(error instanceof Error && error.message.includes('the client is offline')) {
          console.error("Please check your Firebase configuration.");
          setError("Firebase connection failed. Please check your configuration.");
        }
      }
    };
    testConnection();

    const unsubscribe = onAuthStateChanged(auth, (u) => {
      if (u) {
        setUser(u);
      } else {
        signInAnonymously(auth).catch(console.error);
      }
    });
    return () => unsubscribe();
  }, []);

  // Fetch History
  useEffect(() => {
    if (!user || !isLoggedIn) return;

    const q = query(
      collection(db, "conversations"),
      where("userId", "==", user.uid),
      orderBy("createdAt", "desc")
    );

    const unsubscribe = onSnapshot(q, (snapshot) => {
      const docs = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      }));
      setHistory(docs);
    }, (err) => {
      console.error("Firestore error:", err);
    });

    return () => unsubscribe();
  }, [user, isLoggedIn]);

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
    try {
      // Generate a topic from the first few messages
      const firstUserMessage = currentTranscript.find(t => t.startsWith("Pesala: "))?.replace("Pesala: ", "") || "";
      const firstTaniaMessage = currentTranscript.find(t => t.startsWith("Tania: "))?.replace("Tania: ", "") || "";
      const topic = (firstUserMessage || firstTaniaMessage || "New Conversation").slice(0, 40) + ( (firstUserMessage || firstTaniaMessage).length > 40 ? "..." : "");

      await addDoc(collection(db, "conversations"), {
        userId: user.uid,
        createdAt: serverTimestamp(),
        transcript: currentTranscript,
        topic: topic
      });
      setSaveSuccess(true);
      setTimeout(() => setSaveSuccess(false), 3000);
    } catch (err) {
      console.error("Error saving conversation:", err);
      setError("Failed to save conversation to history.");
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
    if (isLoggedIn && !isConnected) {
      toggleConnection();
    }
  }, [isLoggedIn]);

  const clearTranscript = () => {
    setTranscript([]);
    transcriptRef.current = [];
  };

  const deleteHistoryItem = async (id: string) => {
    try {
      await deleteDoc(doc(db, "conversations", id));
      // History will update automatically via onSnapshot
    } catch (err) {
      console.error("Error deleting history item:", err);
      setError("Failed to delete conversation.");
    }
  };

  const deleteAllHistory = async () => {
    if (!user) return;
    if (!window.confirm("Are you sure you want to delete all conversation history? This cannot be undone.")) return;

    try {
      const q = query(collection(db, "conversations"), where("userId", "==", user.uid));
      const querySnapshot = await getDocs(q);
      const batch = writeBatch(db);
      querySnapshot.forEach((doc) => {
        batch.delete(doc.ref);
      });
      await batch.commit();
    } catch (err) {
      console.error("Error deleting all history:", err);
      setError("Failed to clear history.");
    }
  };

  const exportToWord = async (content?: string[] | any) => {
    // If called as an event handler, content will be the event object.
    // We only want to use content if it's explicitly an array of strings.
    const currentTranscript = Array.isArray(content) ? content : transcriptRef.current;
    
    if (!currentTranscript || currentTranscript.length === 0) {
      setError("No conversation to export yet. Start talking to Tania first!");
      return;
    }

    setIsExporting(true);
    try {
      const doc = new Document({
        sections: [
          {
            properties: {
              page: {
                margin: {
                  top: 720,
                  right: 720,
                  bottom: 720,
                  left: 720,
                },
              },
            },
            headers: {
              default: new Header({
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({
                        text: "TANIA AI ASSISTANT - OFFICIAL TRANSCRIPT",
                        bold: true,
                        color: "EA580C",
                        size: 20,
                      }),
                    ],
                    alignment: AlignmentType.CENTER,
                  }),
                ],
              }),
            },
            footers: {
              default: new Footer({
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({
                        text: `© ${new Date().getFullYear()} Pesala Jayawardene. Confidential. Page `,
                        size: 18,
                      }),
                      new TextRun({
                        children: [PageNumber.CURRENT],
                        size: 18,
                      }),
                    ],
                    alignment: AlignmentType.CENTER,
                  }),
                ],
              }),
            },
            children: [
              new Paragraph({
                children: [
                  new TextRun({
                    text: "CONVERSATION TRANSCRIPT",
                    bold: true,
                    size: 36,
                    color: "18181B",
                  }),
                ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 200 },
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: `Date: ${new Date().toLocaleString()}`,
                    size: 20,
                    color: "71717A",
                  }),
                ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 400 },
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "User: Pesala Jayawardene",
                    bold: true,
                    size: 20,
                  }),
                ],
                spacing: { after: 400 },
              }),
              ...currentTranscript.map((line) => {
                const isTania = line.startsWith("Tania:");
                const speaker = isTania ? "Tania" : "Pesala";
                const text = line.replace(/^(Tania|Pesala|You): /, "");

                return new Paragraph({
                  children: [
                    new TextRun({
                      text: `${speaker}: `,
                      bold: true,
                      color: isTania ? "EA580C" : "18181B",
                    }),
                    new TextRun({
                      text: text,
                    }),
                  ],
                  spacing: { before: 120, after: 120 },
                  indent: { left: 720, hanging: 720 },
                });
              }),
            ],
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
      const pageWidth = 210;
      const margin = 10;
      const contentWidth = pageWidth - (margin * 2);

      // Create a temporary hidden container to render the transcript for high-quality capture
      const container = document.createElement('div');
      container.id = "pdf-export-container";
      container.style.position = 'absolute';
      container.style.left = '-9999px';
      container.style.top = '0';
      container.style.width = '800px'; 
      container.style.backgroundColor = '#050505'; 
      container.style.color = '#f4f4f5';
      container.style.fontFamily = 'Arial, sans-serif';
      container.style.padding = '40px';
      container.style.boxSizing = 'border-box';
      
      // Build the HTML structure with INLINE EXPLICIT HEX COLORS ONLY
      let html = `
        <div style="margin-bottom: 40px; text-align: center; border-bottom: 2px solid #ea580c; padding-bottom: 20px;">
          <h1 style="color: #ea580c; margin: 0; font-size: 28px;">TANIA ASSISTANT FULL REPORT</h1>
          <p style="color: #71717a; margin: 10px 0 0 0; font-size: 14px;">Session Type: Multilingual Live Interaction</p>
          <p style="color: #71717a; margin: 5px 0 0 0; font-size: 14px;">Date: ${new Date().toLocaleString()}</p>
          <p style="color: #71717a; margin: 5px 0 0 0; font-size: 14px;">User: Pesala Jayawardene</p>
          <p style="color: #a1a1aa; margin: 15px 0 0 0; font-size: 12px; font-style: italic;">Status: Complete Dialogue & Research Export</p>
        </div>
        <div style="display: flex; flex-direction: column; gap: 24px;">
      `;
      
      transcriptToExport.forEach(line => {
        const isTania = line.startsWith("Tania:");
        const isPesala = line.startsWith("Pesala:");
        const isSystem = line.startsWith("System:");
        const cleanText = line.replace(/^(Tania|Pesala|You|System): /, "");
        const speaker = isTania ? "Tania" : isPesala ? "Pesala" : isSystem ? "System" : "";
        
        const speakerColor = isTania ? '#ea580c' : isPesala ? '#60a5fa' : '#71717a';
        const bgColor = isTania ? '#1c1c1e' : isPesala ? '#2a1a10' : '#050505';
        const borderColor = isTania ? '#3a3a3c' : isPesala ? '#4a2a1a' : '#1c1c1e';
        
        html += `
          <div style="margin-bottom: 20px; text-align: ${isPesala ? 'right' : 'left'}">
            <div style="font-size: 11px; font-weight: bold; color: ${speakerColor}; text-transform: uppercase; margin-bottom: 4px;">
              ${speaker}
            </div>
            <div style="display: inline-block; max-width: 85%; padding: 12px 16px; border-radius: 12px; font-size: 14px; line-height: 1.6; background-color: ${bgColor}; border: 1px solid ${borderColor}; color: #ffffff; text-align: left; ${isSystem ? 'font-style: italic;' : ''}">
              ${cleanText}
            </div>
          </div>
        `;
      });
      
      html += `
        </div>
        <div style="margin-top: 60px; text-align: center; color: #52525b; font-size: 10px; border-top: 1px solid #1c1c1e; padding-top: 20px;">
          © ${new Date().getFullYear()} Pesala Jayawardene. Tania AI Software v112.01. Generated via AI Studio Build.
        </div>
      `;
      
      container.innerHTML = html;
      document.body.appendChild(container);

      // Capture the container
      // Adding a small delay to ensure rendering is complete
      await new Promise(resolve => setTimeout(resolve, 100));

      const canvas = await html2canvas(container, {
        scale: 2,
        useCORS: true,
        backgroundColor: "#050505",
        logging: false,
        onclone: (clonedDoc) => {
          // html2canvas fails on CSS functions it doesn't understand, like oklch (Tailwind 4)
          // Since our export container uses inline HEX styles, we can safely remove
          // all style/link tags from the cloned document to avoid these errors.
          const styleTags = Array.from(clonedDoc.getElementsByTagName('style'));
          styleTags.forEach(tag => tag.remove());
          
          const linkTags = Array.from(clonedDoc.getElementsByTagName('link'));
          linkTags.forEach(tag => {
            if (tag.rel === 'stylesheet') tag.remove();
          });
        }
      });
      
      document.body.removeChild(container);

      const imgData = canvas.toDataURL('image/png');
      const imgHeight = (canvas.height * contentWidth) / canvas.width;
      
      let heightLeft = imgHeight;
      let position = margin;

      // Add the first page
      pdf.addImage(imgData, 'PNG', margin, position, contentWidth, imgHeight);
      heightLeft -= pageHeight;

      // Add subsequent pages if needed
      while (heightLeft >= 0) {
        position = heightLeft - imgHeight;
        pdf.addPage();
        pdf.addImage(imgData, 'PNG', margin, position, contentWidth, imgHeight);
        heightLeft -= pageHeight;
      }

      pdf.save(`Tania_Transcript_${new Date().toISOString().split('T')[0]}.pdf`);
      setSaveSuccess(true);
      setTimeout(() => setSaveSuccess(false), 3000);
    } catch (error) {
      console.error("PDF Export error:", error);
      setError("Failed to export PDF. Try using the 'Print' button or copying the text as a backup.");
    } finally {
      setIsExporting(false);
    }
  };

  const toggleConnection = async () => {
    setError(null);
    if (isConnected) {
      // Auto-save and export on manual disconnect if there is content
      if (transcriptRef.current.length > 2) {
        saveConversation();
        exportToPDF();
      }
      liveApiRef.current?.disconnect();
      setIsConnected(false);
      setIsTalking(false);
      setVolume(0);
    } else {
      setIsConnecting(true);
      try {
        // Retrieve API key from various possible sources
        let apiKey = "";
        
        // 1. Try defined process.env (Vite define)
        if (typeof process !== "undefined" && process.env) {
          apiKey = process.env.GEMINI_API_KEY || process.env.API_KEY || "";
        }
        
        // 2. Try import.meta.env (Vite standard)
        if (!apiKey && (import.meta as any).env) {
          apiKey = (import.meta as any).env.VITE_GEMINI_API_KEY || "";
        }
        
        // Sanitize and validate
        apiKey = apiKey.trim();
        if (apiKey === "undefined" || apiKey === "null") apiKey = "";
        
        if (!apiKey) {
          setError("Tania's voice engine needs an API Key. Please add 'GEMINI_API_KEY' in the 'Settings > Secrets' panel and click 'Save & Restart'.");
          setIsConnecting(false);
          return;
        }

        console.log("Connecting with model: gemini-3.1-flash-live-preview");
        const api = new LiveAPI({
          apiKey,
          model: "gemini-3.1-flash-live-preview",
          systemInstruction: SYSTEM_INSTRUCTION,
          tools: TOOLS,
        });

        hasSavedSessionRef.current = false; // Reset saved state for the new session

        const handleSessionEnded = () => {
          setIsConnected(false);
          setIsTalking(false);
          setVolume(0);
          const currentTranscript = transcriptRef.current;
          if (currentTranscript.length > 2 && !hasSavedSessionRef.current) {
            console.log("Auto-saving live transcript upon session termination...");
            saveConversation();
          }
        };

        await api.connect({
          onOpen: () => {
            console.log("Connection opened");
            setIsConnecting(false);
            addTranscriptLine("System: Connection established. Tania is waking up...");
            // Small delay to ensure session is fully ready for input
            setTimeout(() => {
              addTranscriptLine("System: Sending greeting request...");
              api.sendText("Tania, please greet Pesala Jayawardene immediately with a warm 'Ayubowan' and tell him you are ready to help. Speak now.");
            }, 1000);
          },
          onVolumeChange: (v) => {
            setVolume(v);
            setIsTalking(v > 0.005);
          },
          onTranscript: (text) => {
            console.log("Tania Transcript Received:", text);
            addTranscriptLine(text);
          },
          onInterrupted: () => {
            setIsTalking(false);
            setVolume(0);
          },
          onClose: () => {
            console.log("Session onClose callback received in UI");
            handleSessionEnded();
          },
          onError: (err) => {
            console.error("Live API Error:", err);
            setIsConnecting(false);
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
          },
          onToolCall: async (fc) => {
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
            if (fc.name === "save_to_cloud") {
              await saveConversation();
              return { success: true, message: "Conversation saved to cloud secure history." };
            }
            return { error: "Unknown tool call" };
          }
        });

        liveApiRef.current = api;
        setIsConnected(true);
      } catch (err: any) {
        console.error("Connection error:", err);
        setIsConnecting(false);
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
            <ScrollArea className="flex-1 p-6">
              <div className="space-y-6">
                {history.length === 0 ? (
                  <p className="text-zinc-500 text-center py-20 font-mono text-xs uppercase tracking-widest">No saved conversations</p>
                ) : (
                  history.map((item) => (
                    <div key={item.id} className="bg-zinc-950/50 border border-zinc-800 rounded-2xl p-4 space-y-3 relative group">
                      <div className="flex flex-col gap-1">
                        <div className="flex items-center justify-between">
                          <span className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest">
                            {item.createdAt?.toDate ? item.createdAt.toDate().toLocaleString() : 'Recent'}
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
            </ScrollArea>
          </motion.div>
        )}
      </AnimatePresence>

      <div className="w-full max-w-4xl space-y-8 flex flex-col items-center">
        {/* Header */}
        <div className="flex flex-col items-center text-center space-y-2 max-w-md w-full">
          <div className="w-12 h-12 bg-orange-600 rounded-2xl flex items-center justify-center shadow-xl shadow-orange-900/20 mb-2">
            <Sparkles className="text-white w-7 h-7" />
          </div>
          <h1 className="text-2xl font-medium tracking-tight">Tania</h1>
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

        {/* Transcript Area */}
        <div className="bg-zinc-900/40 backdrop-blur-2xl border border-white/5 rounded-[2rem] overflow-hidden flex flex-col h-[320px] sm:h-[420px] md:h-[500px] w-full shadow-[0_20px_50px_rgba(0,0,0,0.5)] relative group/transcript transition-all">
          <div className="sticky top-0 z-20 px-6 py-4 border-b border-white/5 bg-zinc-900/80 backdrop-blur-md flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="flex items-center gap-2">
                <div className={`w-2 h-2 rounded-full ${isConnected ? 'bg-emerald-500 animate-pulse' : isConnecting ? 'bg-orange-500 animate-bounce' : 'bg-red-500'} shadow-[0_0_8px_rgba(234,88,12,0.5)]`} />
                <span className="text-[10px] font-mono uppercase tracking-[0.2em] text-zinc-100 font-bold">
                  {isConnected ? 'Live Stream' : isConnecting ? 'Connecting...' : 'Ready'}
                </span>
              </div>
              {transcript.length > 0 && (
                <button 
                  onClick={clearTranscript}
                  className="ml-2 text-[8px] text-zinc-500 hover:text-orange-400 uppercase tracking-widest font-mono transition-all hover:scale-110"
                >
                  [Clear]
                </button>
              )}
            </div>
            <div className="flex items-center gap-4">
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
              <div className="h-4 w-[1px] bg-white/10" />
              <button 
                onClick={saveConversation}
                className={`p-1.5 transition-all hover:scale-110 ${isSaving ? 'text-zinc-600' : 'text-emerald-500 hover:text-emerald-400'}`}
                title="Save to Cloud"
                disabled={isSaving}
              >
                <Save className="w-4 h-4" />
              </button>
              <button 
                onClick={() => exportToPDF()}
                className={`p-1.5 transition-all hover:scale-110 ${isExporting ? 'text-zinc-600' : 'text-blue-500 hover:text-blue-400'}`}
                title="Export as PDF"
                disabled={isExporting}
              >
                <FileText className="w-4 h-4" />
              </button>
              <button 
                onClick={() => printConversation(transcript)}
                className="p-1.5 text-zinc-400 hover:text-orange-400 transition-all hover:scale-110"
                title="Print Transcript"
              >
                <Printer className="w-4 h-4" />
              </button>
            </div>
          </div>
          
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
                  const cleanText = text.replace(/^(Tania|Pesala|You|System): /, "");
                  
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
                      </div>
                      <div className={`max-w-[85%] px-4 py-3 rounded-2xl text-sm leading-relaxed shadow-lg ${
                        isTania 
                          ? 'bg-zinc-800/50 border border-white/5 text-zinc-200 rounded-tl-none' 
                          : isPesala 
                            ? 'bg-orange-600/20 border border-orange-500/20 text-orange-50 text-left rounded-tr-none'
                            : 'bg-zinc-900/80 border border-zinc-800 text-zinc-400 italic text-xs'
                      }`}>
                        {cleanText}
                      </div>
                    </motion.div>
                  );
                })
              )}
              <div ref={messagesEndRef} />
            </div>
          </div>

          <div className="px-6 py-4 border-t border-white/5 bg-zinc-900/40 flex items-center justify-between">
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
            </div>
            {isConnected && (
              <div className="flex items-center gap-2">
                <div className="w-2 h-2 rounded-full bg-red-500 animate-pulse shadow-[0_0_10px_rgba(239,68,68,0.5)]" />
                <span className="text-[10px] font-mono text-red-500 font-bold uppercase tracking-widest">Live</span>
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
    </div>
  );
}

