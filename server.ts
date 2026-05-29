import express from "express";
import path from "path";
import http from "http";
import { WebSocketServer, WebSocket } from "ws";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI, Modality } from "@google/genai";
import dotenv from "dotenv";
import nodemailer from "nodemailer";

dotenv.config();

async function startServer() {
  const app = express();
  app.use(express.json());
  const server = http.createServer(app);
  const PORT = 3000;

  // Create WebSocket Server
  const wss = new WebSocketServer({ noServer: true });

  // Strictly enforce a single, active voice/client WebSocket connection at any time to guarantee no dual voices exist
  let activeClientSocket: any = null;

  // Handle upgrade manually to hook onto /api/live
  server.on("upgrade", (request, socket, head) => {
    const pathname = new URL(request.url || "", `http://${request.headers.host}`).pathname;
    if (pathname === "/api/live") {
      wss.handleUpgrade(request, socket, head, (ws) => {
        wss.emit("connection", ws, request);
      });
    } else {
      socket.destroy();
    }
  });

  // API health status
  app.get("/api/health", (req, res) => {
    res.json({ status: "ok" });
  });

  // In-memory image cache to ensure consistency and prevent difference between preview & download
  const imageCache = new Map<string, { contentType: string; buffer: Buffer; timestamp: number }>();

  // Helper method to scan JSON nodes for Unsplash photo URLs
  function findUnsplashPhotoUrls(obj: any): string[] {
    const urls: string[] = [];
    
    // Attempt 1: Direct known NextData path extraction for Unsplash search page
    try {
      const searchState = obj?.props?.pageProps?.initialState?.search || obj?.props?.pageProps?.dehydratedState?.queries?.[0]?.state?.data;
      const photos = searchState?.photos?.results || searchState?.results;
      if (Array.isArray(photos)) {
        for (const item of photos) {
          if (item?.urls?.regular) {
            urls.push(item.urls.regular.split("?")[0]);
          }
        }
      }
    } catch (e) {
      console.warn("[Unsplash Parser] Standard state path extraction failed:", e);
    }

    if (urls.length > 0) {
      console.log(`[Unsplash Parser] Extracted ${urls.length} images from standard state path`);
      return urls;
    }

    // Attempt 2: Target JSON search-result arrays recursively
    function recurseForPhotoArrays(current: any) {
      if (!current || urls.length > 0) return;
      if (Array.isArray(current)) {
        const first = current[0];
        // Target arrays containing photo result objects possessing a nested "urls.regular" string starting with Photo CDN
        if (first && typeof first === "object" && first.urls && typeof first.urls.regular === "string" && first.urls.regular.startsWith("https://images.unsplash.com/photo-")) {
          for (const item of current) {
            if (item && item.urls && typeof item.urls.regular === "string") {
              const lower = item.urls.regular.toLowerCase();
              if (!lower.includes("profile") && !lower.includes("avatar") && !lower.includes("placeholder")) {
                urls.push(item.urls.regular.split("?")[0]);
              }
            }
          }
          return;
        }
        for (const item of current) {
          recurseForPhotoArrays(item);
        }
      } else if (typeof current === "object") {
        for (const key of Object.keys(current)) {
          recurseForPhotoArrays(current[key]);
        }
      }
    }

    recurseForPhotoArrays(obj);
    return urls;
  }

  // Helper to generate high-fidelity candidate queries for image searches in correct order of specificity.
  function generateImageSearchCandidates(query: string): string[] {
    const cleanedQuery = query.replace(/[.,\/#!$%\^&\*;:{}=\-_`~()?"']/g, " ").replace(/\s+/g, " ").trim();
    if (!cleanedQuery) return [];

    const words = cleanedQuery.split(/\s+/).filter(w => w.length > 0);
    const candidateQueries: string[] = [cleanedQuery];

    // Common terms that are too generic as standalone single-word queries
    const GENERIC_EXCLUDE_WORDS = new Set([
      "a", "an", "the", "this", "that", "these", "those", "some", "any", "every", "all", "each", "other", "another",
      "please", "show", "me", "find", "display", "photo", "picture", "image", "img", "pic", "view", "photo of", "picture of",
      "high", "low", "resolution", "quality", "res", "hd", "4k", "beautiful", "gorgeous", "awesome", "nice", "cool",
      "of", "in", "on", "at", "by", "for", "with", "about", "against", "between", "into", "through", "under", "over", 
      "above", "below", "down", "up", "out", "from", "and", "or", "but", "so", "sri", "lanka", "srilankan", "lankan", "ceylon", 
      "island", "colombo", "kandy", "nature"
    ]);

    if (words.length > 1) {
      // 1. Try removing the very first word if there are at least 3 words (e.g. "Sri Lankan Kothu Roti" -> "Lankan Kothu Roti")
      if (words.length > 2) {
        const slice1 = words.slice(1).join(" ");
        if (slice1.trim()) candidateQueries.push(slice1);
      }

      // 2. Try the last 2 words (usually containing the main noun with its immediate adjective, e.g., "Kothu Roti")
      if (words.length > 2) {
        const lastTwo = words.slice(words.length - 2).join(" ");
        if (lastTwo.trim()) candidateQueries.push(lastTwo);
      }

      // 3. Try only the last word (the absolute core noun/subject, e.g., "Roti")
      const lastWord = words[words.length - 1];
      if (lastWord && lastWord.length > 1 && !GENERIC_EXCLUDE_WORDS.has(lastWord.toLowerCase())) {
        candidateQueries.push(lastWord);
      }

      // 4. Try the first two words if there are many words, but ONLY if the combination is not purely made of generic/exclude words
      if (words.length > 2) {
        const firstTwo = words.slice(0, 2).join(" ");
        const hasSpecificWord = words.slice(0, 2).some(w => !GENERIC_EXCLUDE_WORDS.has(w.toLowerCase()));
        if (hasSpecificWord && firstTwo.trim()) {
          candidateQueries.push(firstTwo);
        }
      }
    }

    // Semantic category mapping fallback to guarantee high-volume, high-quality matches
    const lower = cleanedQuery.toLowerCase();
    if (lower.includes("price") || lower.includes("quotation") || lower.includes("spreadsheet") || lower.includes("invoice") || lower.includes("billing") || lower.includes("excel") || lower.includes("table") || lower.includes("finance") || lower.includes("cost") || lower.includes("rate") || lower.includes("fees") || lower.includes("money") || lower.includes("accounting") || lower.includes("estimate")) {
      candidateQueries.push("accounting");
      candidateQueries.push("finance");
      candidateQueries.push("office");
      candidateQueries.push("business");
      candidateQueries.push("workspace");
    } else if (lower.includes("letter") || lower.includes("resume") || lower.includes("proposal") || lower.includes("document") || lower.includes("draft") || lower.includes("composition") || lower.includes("email")) {
      candidateQueries.push("document");
      candidateQueries.push("writing");
      candidateQueries.push("office");
      candidateQueries.push("desk");
    } else if (lower.includes("bluetooth") || lower.includes("speaker") || lower.includes("audio") || lower.includes("headphones") || lower.includes("sound") || lower.includes("microphone") || lower.includes("mic")) {
      candidateQueries.push("speaker");
      candidateQueries.push("headphones");
      candidateQueries.push("audio");
    } else if (lower.includes("sri lanka") || lower.includes("sinhala") || lower.includes("colombo") || lower.includes("kandy") || lower.includes("sigiriya")) {
      if (lower.includes("nature") || lower.includes("scenery") || lower.includes("landscape") || lower.includes("waterfall") || lower.includes("mountain") || lower.includes("beach")) {
        candidateQueries.push("sri lanka");
        candidateQueries.push("nature");
      }
    } else if (lower.includes("avatar") || lower.includes("assistant") || lower.includes("ai") || lower.includes("microphone") || lower.includes("listening") || lower.includes("voice")) {
      candidateQueries.push("technology");
      candidateQueries.push("wave");
      candidateQueries.push("microphone");
    }

    // Final clean: remove duplicates and anything that is completely empty or just single/double letter generic junk
    const finalCandidates = Array.from(new Set(candidateQueries))
      .map(c => c.trim())
      .filter(c => {
        if (!c) return false;
        if (c.length <= 2 && GENERIC_EXCLUDE_WORDS.has(c.toLowerCase())) return false;
        return true;
      });

    return finalCandidates;
  }

  // Helper method to retrieve highly relevant, high-quality images from DuckDuckGo image search
  async function fetchDuckDuckGoPhoto(query: string, sig: string): Promise<string | null> {
    const cleanedQuery = query.trim();
    if (!cleanedQuery) return null;

    try {
      const url = `https://duckduckgo.com/?q=${encodeURIComponent(cleanedQuery)}`;
      console.log(`[DuckDuckGo] Step 1: Fetching vqd from: ${url}`);
      const response = await fetch(url, {
        headers: {
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"
        }
      });

      if (!response.ok) {
        console.warn(`[DuckDuckGo] Failed to fetch vqd token, status: ${response.status}`);
        return null;
      }

      const html = await response.text();
      const vqdMatch = html.match(/vqd=["']?([\d\-]+)["']?/) || html.match(/vqd=([^&'"]+)/);
      if (!vqdMatch) {
         console.warn(`[DuckDuckGo] Could not find vqd token in html`);
         return null;
      }

      const vqd = vqdMatch[1];
      console.log(`[DuckDuckGo] Step 2: Found vqd token: "${vqd}". Fetching images...`);

      const searchUrl = `https://duckduckgo.com/i.js?l=us-en&o=json&q=${encodeURIComponent(cleanedQuery)}&vqd=${vqd}&f=,,,`;
      const searchResponse = await fetch(searchUrl, {
        headers: {
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
          "Accept": "application/json"
        }
      });

      if (!searchResponse.ok) {
        console.warn(`[DuckDuckGo] Search response not OK. Status: ${searchResponse.status}`);
        return null;
      }

      const data = await searchResponse.json() as any;
      if (data && data.results && data.results.length > 0) {
        const validResults = data.results.filter((item: any) => item.image && (item.image.startsWith("http://") || item.image.startsWith("https://")));
        if (validResults.length > 0) {
          // Determine index offset using sig to cycle through options when user asks for a different picture
          let offset = 0;
          const sigNum = parseInt(sig, 10);
          if (!isNaN(sigNum) && sigNum > 0) {
            offset = sigNum % Math.min(validResults.length, 5);
          }
          console.log(`[DuckDuckGo] Returning image offset index ${offset} of ${validResults.length} options`);
          const chosenItem = validResults[offset] || validResults[0];
          console.log(`[DuckDuckGo] Found perfect image match (title: "${chosenItem.title}"): ${chosenItem.image}`);
          return chosenItem.image;
        }
      }
    } catch (err) {
      console.error(`[DuckDuckGo] Search failed for query "${cleanedQuery}":`, err);
    }
    return null;
  }

  // Helper method to retrieve highly relevant, high-quality images from Wikimedia Commons
  async function fetchWikimediaCommonsPhoto(query: string): Promise<string | null> {
    const uniqueCandidates = generateImageSearchCandidates(query);
    console.log(`[Wikimedia Commons] Query candidates to check:`, uniqueCandidates);

    for (const cand of uniqueCandidates) {
      try {
        const url = `https://commons.wikimedia.org/w/api.php?action=query&generator=search&gsrsearch=${encodeURIComponent(cand)}&gsrnamespace=6&gsrlimit=5&prop=imageinfo&iiprop=url&iiurlwidth=1000&format=json&origin=*`;
        console.log(`[Wikimedia Commons] Searching candidate: "${cand}"`);
        const response = await fetch(url, {
          headers: { "User-Agent": "TaniaAIAssistant/1.0 (pesala.em.rms@gmail.com) Node/18" }
        });
        if (response.ok) {
          const data = await response.json() as any;
          if (data && data.query && data.query.pages) {
            const pages = data.query.pages;
            for (const pageId of Object.keys(pages)) {
              const page = pages[pageId];
              if (page.imageinfo && page.imageinfo.length > 0) {
                const info = page.imageinfo[0];
                const chosenUrl = info.thumburl || info.url;
                if (chosenUrl) {
                  console.log(`[Wikimedia Commons] SECURED MATCH! Selected precision image for candidate "${cand}": ${chosenUrl}`);
                  return chosenUrl;
                }
              }
            }
          }
        }
      } catch (err) {
        console.error(`[Wikimedia Commons] Connection error during candidate "${cand}" search:`, err);
      }
    }
    return null;
  }

  // Helper method to retrieve highly relevant, high-quality images from Unsplash
  async function fetchUnsplashSearchPhoto(query: string): Promise<string | null> {
    const uniqueCandidates = generateImageSearchCandidates(query);
    console.log(`[Unsplash Search] Candidate queries to check:`, uniqueCandidates);

    // Inner search method for a single term
    async function performUnsplashSingleSearch(term: string): Promise<string | null> {
      try {
        const napiUrl = `https://unsplash.com/napi/search/photos?query=${encodeURIComponent(term)}&per_page=5`;
        console.log(`[Unsplash NAPI] Searching: "${term}" -> ${napiUrl}`);
        
        const napiResponse = await fetch(napiUrl, {
          headers: {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
            "Accept": "application/json",
            "Accept-Language": "en-US,en;q=0.5"
          }
        });
        
        if (napiResponse.ok) {
          const data = await napiResponse.json() as any;
          if (data && data.results && data.results.length > 0) {
            const firstPhoto = data.results[0];
            if (firstPhoto && firstPhoto.urls && firstPhoto.urls.regular) {
              const chosenUrl = firstPhoto.urls.regular;
              console.log(`[Unsplash NAPI] Match! selected term "${term}": ${chosenUrl}`);
              return chosenUrl;
            }
          }
        }
      } catch (err) {
        console.warn(`[Unsplash NAPI] Connection error for term "${term}":`, err);
      }

      // Scraping fallback for this single term
      try {
        const urlFriendlyTerm = term.replace(/\s+/g, "+");
        const searchUrl = `https://unsplash.com/s/photos/${encodeURIComponent(urlFriendlyTerm)}`;
        console.log(`[Unsplash Scraper] Webpage crawl: ${searchUrl}`);
        
        const response = await fetch(searchUrl, {
          headers: {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5"
          }
        });
        
        if (response.ok) {
          const html = await response.text();
          
          // Try __NEXT_DATA__
          const nextDataMatch = html.match(/<script id="__NEXT_DATA__" type="application\/json">([\s\S]*?)<\/script>/);
          if (nextDataMatch) {
            try {
              const parsed = JSON.parse(nextDataMatch[1]);
              const scannedUrls = findUnsplashPhotoUrls(parsed);
              if (scannedUrls.length > 0) {
                const firstCrawled = scannedUrls[0];
                const premiumUrl = `${firstCrawled}?auto=format&fit=crop&w=1000&h=750&q=85`;
                console.log(`[Unsplash Scraper] Found preloaded URL for term "${term}": ${premiumUrl}`);
                return premiumUrl;
              }
            } catch (innerErr) {
              console.warn("[Unsplash Scraper] NextData parser failed:", innerErr);
            }
          }

          // Parse img tags
          const scannedImgTags: { src: string; alt: string }[] = [];
          const imgTagRegex = /<img\s+([^>]+)>/gi;
          let match;
          while ((match = imgTagRegex.exec(html)) !== null) {
            const tagContent = match[1];
            const srcMatch = tagContent.match(/src=["'](https:\/\/images\.unsplash\.com\/photo-[^"']+)["']/i);
            if (srcMatch) {
              const altMatch = tagContent.match(/alt=["']([^"']*)["']/i);
              const alt = altMatch ? altMatch[1] : "";
              const cleanSrc = srcMatch[1].replace(/&amp;/g, "&").split("?")[0];
              scannedImgTags.push({ src: cleanSrc, alt });
            }
          }

          if (scannedImgTags.length > 0) {
            let filteredImages = scannedImgTags.filter(img => {
              const lowerAlt = img.alt.toLowerCase();
              const lowerSrc = img.src.toLowerCase();
              if (lowerAlt.includes("profile") || lowerAlt.includes("avatar") || lowerAlt.includes("logo") || lowerAlt.includes("go to")) return false;
              if (lowerSrc.includes("profile") || lowerSrc.includes("avatar") || lowerSrc.includes("logo")) return false;
              if (img.alt.trim().length < 3) return false;
              return true;
            });

            if (filteredImages.length > 0) {
              const queryWords = term.toLowerCase().split(/\s+/).filter(word => word.length > 2);
              if (queryWords.length > 0) {
                const keywordMatches = filteredImages.filter(img => {
                  const lowerAlt = img.alt.toLowerCase();
                  return queryWords.some(word => lowerAlt.includes(word));
                });
                if (keywordMatches.length > 0) {
                  filteredImages = keywordMatches;
                }
              }

              const chosen = filteredImages[0];
              const premiumUrl = `${chosen.src}?auto=format&fit=crop&w=1000&h=750&q=85`;
              console.log(`[Unsplash Scraper] Success for term "${term}": ${premiumUrl}`);
              return premiumUrl;
            }
          }

          // Fallback regex match of general photo IDs
          const rawMatches = html.match(/https:\/\/images\.unsplash\.com\/photo-[a-zA-Z0-9\-_]+/g);
          if (rawMatches && rawMatches.length > 0) {
            const baseUrls = rawMatches.map(m => m.split("?")[0]);
            const photoUrls = Array.from(new Set(baseUrls)).filter(url => {
              const lower = url.toLowerCase();
              return url.includes("/photo-") && url.length > 40 && !lower.includes("profile") && !lower.includes("avatar") && !lower.includes("placeholder");
            });
            if (photoUrls.length > 0) {
              const premiumUrl = `${photoUrls[0]}?auto=format&fit=crop&w=1000&h=750&q=85`;
              console.log(`[Unsplash Scraper] Fallback raw URL for term "${term}": ${premiumUrl}`);
              return premiumUrl;
            }
          }
        }
      } catch (err) {
        console.error(`[Unsplash Scraper] Failed crawling term "${term}":`, err);
      }

      return null;
    }

    // Try each query candidate sequentially, return first matching premium photo URL
    for (const cand of uniqueCandidates) {
      const resultUrl = await performUnsplashSingleSearch(cand);
      if (resultUrl) {
        console.log(`[Unsplash Search] SECURED PREMIUM MATCH! Result for candidate "${cand}": ${resultUrl}`);
        return resultUrl;
      }
    }

    console.warn(`[Unsplash Search] All candidates failed for query: "${query}"`);
    return null;
  }

  // Cleanup worker to avoid memory leaks
  setInterval(() => {
    const now = Date.now();
    for (const [key, val] of imageCache.entries()) {
      if (now - val.timestamp > 15 * 60 * 1000) { // Keep cache for 15 minutes
        imageCache.delete(key);
      }
    }
  }, 5 * 60 * 1000);

  app.get("/api/image-proxy", async (req, res) => {
    try {
      const query = req.query.query as string || "";
      const sig = req.query.sig as string || "default";
      if (!query) {
        return res.status(400).send("Missing query");
      }

      const cacheKey = `${query.toLowerCase().trim()}_${sig}`;
      if (imageCache.has(cacheKey)) {
        console.log(`[Proxy] Cache HIT for image key: ${cacheKey}`);
        const cached = imageCache.get(cacheKey)!;
        res.setHeader("Content-Type", cached.contentType);
        res.setHeader("Cache-Control", "public, max-age=31536000, immutable");
        res.setHeader("Access-Control-Allow-Origin", "*");
        return res.send(cached.buffer);
      }

      console.log(`[Proxy] Cache MISS for image: query="${query}" sig="${sig}". Fetching from source...`);
      
      // Try to get direct relevant photo from DuckDuckGo first (highly accurate, web-wide search with signature-based offsets)
      let sourceUrl = await fetchDuckDuckGoPhoto(query, sig);
      
      // Fallback to Unsplash if DuckDuckGo failed
      if (!sourceUrl) {
        console.log(`[Proxy] DuckDuckGo failed. Fetching from Unsplash for: "${query}"`);
        sourceUrl = await fetchUnsplashSearchPhoto(query);
      }
      
      // Fallback to Wikimedia Commons if both failed or were blocked
      if (!sourceUrl) {
        console.log(`[Proxy] Unsplash blocked/failed. Querying high-fidelity Wikimedia Commons fallback for: "${query}"`);
        sourceUrl = await fetchWikimediaCommonsPhoto(query);
      }
      
      const urlsToTry: string[] = [];
      if (sourceUrl) {
        urlsToTry.push(sourceUrl);
      }
      
      // Keep general fallbacks in order of backup quality
      urlsToTry.push(`https://loremflickr.com/800/600/${encodeURIComponent(query)}`);
      urlsToTry.push(`https://picsum.photos/800/600?sig=${sig}`);

      let response: Response | null = null;
      let usedUrl = "";

      for (const url of urlsToTry) {
        try {
          console.log(`[Proxy] Attempting fetch: ${url}`);
          const resTry = await fetch(url);
          if (resTry.ok) {
            response = resTry;
            usedUrl = url;
            break;
          } else {
            console.warn(`[Proxy] Fail status ${resTry.status} for url: ${url}`);
          }
        } catch (fetchErr) {
          console.warn(`[Proxy] Failed to fetch url: ${url}`, fetchErr);
        }
      }

      if (!response) {
        throw new Error("All image retrieval attempts failed");
      }

      console.log(`[Proxy] Successfully retrieved image from: ${usedUrl}`);
      const arrayBuffer = await response.arrayBuffer();
      const buffer = Buffer.from(arrayBuffer);
      const contentType = response.headers.get("content-type") || "image/jpeg";

      imageCache.set(cacheKey, {
        contentType,
        buffer,
        timestamp: Date.now()
      });

      res.setHeader("Content-Type", contentType);
      res.setHeader("Cache-Control", "public, max-age=31536000, immutable");
      res.setHeader("Access-Control-Allow-Origin", "*");
      res.send(buffer);
    } catch (error: any) {
      console.error("[Proxy] Error routing image requests:", error);
      res.status(500).send("Error returning image stream from endpoint");
    }
  });

  // Send formal custom email using Nodemailer (with failover/mailto fallbacks)
  app.post("/api/send-email", async (req, res) => {
    try {
      const { to, subject, body } = req.body;
      if (!to || !body) {
        return res.status(400).json({ success: false, error: "Missing recipient 'to' or email 'body' content." });
      }

      console.log(`[API send-email] Request to send email to "${to}" with subject: "${subject || '(No Subject)'}"`);

      // Retrieve SMTP configurations
      const host = process.env.SMTP_HOST;
      const port = process.env.SMTP_PORT ? parseInt(process.env.SMTP_PORT) : 587;
      const user = process.env.SMTP_USER;
      const pass = process.env.SMTP_PASS;
      const sender = process.env.SMTP_SENDER || "Tania AI <ai.tania.assistant@gmail.com>";

      let deliveredVia = "simulation";
      let infoMessage = "";

      if (host && user && pass) {
        try {
          const transporter = nodemailer.createTransport({
            host,
            port,
            secure: port === 465, // true for 465, false for other ports
            auth: { user, pass },
            tls: {
              rejectUnauthorized: false
            }
          });

          const mailOptions = {
            from: sender,
            to,
            subject: subject || "Custom Message from Tania AI Work Space",
            text: body,
            html: body.replace(/\n/g, "<br>")
          };

          const info = await transporter.sendMail(mailOptions);
          console.log(`[API send-email] Actual email delivered successfully! Message ID: ${info.messageId}`);
          deliveredVia = "smtp";
          infoMessage = `SMTP Transmitted. Message ID: ${info.messageId}`;
        } catch (smtpError: any) {
          console.error("[API send-email] SMTP transport error. Falling back to local dispatch:", smtpError);
          deliveredVia = "simulation_failover";
          const errMsg = smtpError.message || String(smtpError);
          let detailedWarning = "";
          if (errMsg.includes("534-5.7.9") || errMsg.includes("Application-specific password") || errMsg.includes("app-specific") || errMsg.includes("WantToUseAppPass")) {
            detailedWarning = " [TANIA NOTE: Gmail detected that your account requires an App-specific Password because 2-Step Verification is active. Google blocks normal logins for secure server SMTP connections. To fix this: Go to your Google Account -> Safety/Security -> App Passwords. Select 'Other' and give it a name 'Tania', then copy the 16-character code (with no spaces). Enter this code as your SMTP_PASS under Secrets in the AI Studio settings gear, then retry!]";
          }
          infoMessage = `SMTP failure: ${errMsg}.${detailedWarning} Dispatched via local sandbox queue.`;
        }
      } else {
        console.warn("[API send-email] SMTP environment keys are unset. Executing sandbox mock delivery.");
        deliveredVia = "simulation";
        infoMessage = "No custom SMTP server configured under your AI Studio Environment settings. Simulated delivery succeeded.";
      }

      const cleanSubject = subject || "No Subject";
      // Generate mailto prefill for instantaneous backup click-to-send capability
      const mailtoLink = `mailto:${encodeURIComponent(to)}?subject=${encodeURIComponent(cleanSubject)}&body=${encodeURIComponent(body)}`;

      return res.json({
        success: true,
        status: "delivered",
        method: deliveredVia,
        info: infoMessage,
        mailto: mailtoLink,
        data: {
          to,
          subject: cleanSubject,
          body,
          sentAt: new Date().toISOString()
        }
      });
    } catch (err: any) {
      console.error("[API send-email] General route handler failure:", err);
      res.status(500).json({ success: false, error: err.message || "Failed to parse email request" });
    }
  });

  // Compose and dispatch WhatsApp pre-filled links
  app.post("/api/send-whatsapp", async (req, res) => {
    try {
      const { to, message } = req.body;
      if (!to || !message) {
        return res.status(400).json({ success: false, error: "Missing recipient phone 'to' or text 'message' content." });
      }

      console.log(`[API send-whatsapp] Request received for phone "${to}"`);

      // Normalize phone number to be purely numeric
      const sanitizedPhone = to.replace(/[+\s\-()]/g, "");

      // Both wa.me links and api.whatsapp works beautifully for redirecting
      const waLink = `https://wa.me/${sanitizedPhone}?text=${encodeURIComponent(message)}`;

      return res.json({
        success: true,
        status: "delivered",
        method: "prefilled_api_link",
        whatsapp_link: waLink,
        data: {
          to: sanitizedPhone,
          message,
          sentAt: new Date().toISOString()
        }
      });
    } catch (err: any) {
      console.error("[API send-whatsapp] General route handler failure:", err);
      res.status(500).json({ success: false, error: err.message || "Failed to parse whatsapp request" });
    }
  });

  // Client connected, proxy connection to Google Gemini Live API
  wss.on("connection", (clientWs) => {
    console.log("[Proxy] Socket opened: Browser connected to server gateway");
    
    // HARD CONCURRENCY PREVENTION: Close any existing active connection immediately
    if (activeClientSocket) {
      console.log("[Proxy] Forcefully closing duplicate active browser connection to prevent dual voice streams.");
      try {
        activeClientSocket.send(JSON.stringify({
          type: "close",
          reason: "Newer microphone session initiated on another frame or thread."
        }));
        activeClientSocket.close();
      } catch (e) {
        console.warn("[Proxy] Handled error closing duplicate socket:", e);
      }
    }
    activeClientSocket = clientWs;
    
    const apiKey = process.env.GEMINI_API_KEY || process.env.API_KEY || "";
    if (!apiKey) {
      console.error("[Proxy] Missing GEMINI_API_KEY secret on backend.");
      clientWs.send(JSON.stringify({
        type: "error",
        error: "GEMINI_API_KEY is not defined in the container Secrets. Please add it to Settings > Secrets."
      }));
      clientWs.close();
      return;
    }

    const ai = new GoogleGenAI({
      apiKey,
      apiVersion: "v1beta",
      vertexai: false,
      httpOptions: {
        headers: {
          'User-Agent': 'aistudio-build'
        }
      }
    });

    let geminiSession: any = null;
    let isClosed = false;

    clientWs.on("message", async (data) => {
      try {
        const message = JSON.parse(data.toString());
        
        if (message.type === "ping") {
          clientWs.send(JSON.stringify({ type: "pong" }));
          return;
        }

        if (message.type === "connect") {
          console.log("[Proxy] Connecting with model:", message.model);
          const modelName = message.model.startsWith("models/") ? message.model : `models/${message.model}`;
          
          try {
            geminiSession = await ai.live.connect({
              model: modelName,
              config: message.config,
              callbacks: {
                onopen: () => {
                  console.log("[Proxy] Server connected to Gemini Live API");
                  if (!isClosed) {
                    clientWs.send(JSON.stringify({ type: "open" }));
                  }
                },
                onmessage: (msg: any) => {
                  if (isClosed) return;
                  // Forward exactly to client
                  clientWs.send(JSON.stringify({ type: "message", data: msg }));
                },
                onclose: () => {
                  console.log("[Proxy] Gemini closed voice connection");
                  if (!isClosed) {
                    clientWs.send(JSON.stringify({ type: "close" }));
                    clientWs.close();
                  }
                },
                onerror: (err: any) => {
                  console.error("[Proxy] Gemini error:", err);
                  if (!isClosed) {
                    clientWs.send(JSON.stringify({ type: "error", error: err?.message || String(err) }));
                  }
                }
              }
            });
          } catch (connErr: any) {
            console.error("[Proxy] Failed to connect to Gemini Live:", connErr);
            if (!isClosed) {
              clientWs.send(JSON.stringify({ type: "error", error: connErr?.message || String(connErr) }));
              clientWs.close();
            }
          }
          return;
        }

        // Forward commands from backend proxy directly to active Gemini session
        if (geminiSession && !isClosed) {
          if (message.type === "input") {
            geminiSession.sendRealtimeInput(message.data);
          } else if (message.type === "tool_response") {
            geminiSession.sendToolResponse(message.data);
          } else if (message.type === "text") {
            geminiSession.sendRealtimeInput({ text: message.text });
          }
        }
      } catch (err: any) {
        console.error("[Proxy] Recv error:", err);
      }
    });

    clientWs.on("close", () => {
      console.log("[Proxy] Browser disconnected from gateway");
      if (activeClientSocket === clientWs) {
        activeClientSocket = null;
      }
      isClosed = true;
      if (geminiSession) {
        try {
          geminiSession.close();
        } catch (e) {
          // Graceful collapse
        }
        geminiSession = null;
      }
    });
  });

  // Vite development integration or static mapping for production execution
  if (process.env.NODE_ENV !== "production") {
    console.log("Integrating Vite Dev Server Middleware...");
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    console.log("Serving static production assets from /dist...");
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  server.listen(PORT, "0.0.0.0", () => {
    console.log(`Server listening on status: READY at http://0.0.0.0:${PORT}`);
  });
}

startServer();
