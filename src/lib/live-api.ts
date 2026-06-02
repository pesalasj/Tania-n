export enum Modality {
  MODALITY_UNSPECIFIED = "MODALITY_UNSPECIFIED",
  TEXT = "TEXT",
  IMAGE = "IMAGE",
  AUDIO = "AUDIO",
}

export interface LiveAPIConfig {
  apiKey?: string; // Optional: API Key is held securely server-side
  model: string;
  systemInstruction?: string;
  tools?: any[];
}

const tabSessionId = typeof window !== 'undefined' ? Math.random().toString(36).substring(2, 11) : '';

const safeLocalStorage = {
  getItem(key: string): string | null {
    try {
      const localVal = localStorage.getItem(key);
      if (localVal !== null) return localVal;
      return typeof window !== 'undefined' ? (window as any)[`__win_storage_${key}`] || null : null;
    } catch (e) {
      if (typeof window !== 'undefined') {
        return (window as any)[`__win_storage_${key}`] || null;
      }
      return null;
    }
  },
  setItem(key: string, value: string): void {
    try {
      if (typeof window !== 'undefined') {
        (window as any)[`__win_storage_${key}`] = value;
      }
      localStorage.setItem(key, value);
    } catch (e) {}
  }
};

export class LiveAPI {
  private ws: WebSocket | null = null;
  private audioContext: AudioContext | null = null;
  private stream: MediaStream | null = null;
  private processor: ScriptProcessorNode | null = null;
  private onVolumeChange: (volume: number) => void = () => {};
  private onTranscript: (text: string) => void = () => {};
  private onInterrupted: () => void = () => {};
  private onToolCall: (toolCall: any) => Promise<any> = async () => ({});
  private isConnected = false;
  private audioQueue: Float32Array[] = [];
  private isPlaying = false;
  private isUserSpeaking = false;
  private nextStreamTime = 0;
  private lastTaniaSpeechTime = 0;
  private silentFramesCount = 0;
  private isAudioActive = false;
  private sentSilenceFramesAfterSpeechCount = 0;
  private inputAudioAccumulator: number[] = [];
  private playingSources: AudioBufferSourceNode[] = [];
  private analyser: AnalyserNode | null = null;
  private volumeInterval: any = null;
  private pingInterval: any = null;
  private concurrencyInterval: any = null;
  private isDisposed = false;
  public outputDeviceId = "default";
  private instanceId = Math.random().toString(36).substring(2, 11);

  constructor(private config: LiveAPIConfig) {
    console.log("Initialized LiveAPI Browser Proxy Interface Client");
    if (!(window as any).__allLiveApis) {
      (window as any).__allLiveApis = [];
    }
    (window as any).__allLiveApis.push(this);
  }

  setAudioOutputDevice(deviceId: string) {
    this.outputDeviceId = deviceId || "default";
    if (this.audioContext && typeof (this.audioContext as any).setSinkId === 'function') {
      (this.audioContext as any).setSinkId(this.outputDeviceId)
        .then(() => console.log(`[LiveAPI] Routed context speakers successfully to: ${this.outputDeviceId}`))
        .catch((err: any) => console.error("[LiveAPI] Failed to route context speakers to sinkId:", err));
    }
  }

  async connect(callbacks: {
    onVolumeChange?: (volume: number) => void;
    onTranscript?: (text: string) => void;
    onInterrupted?: () => void;
    onError?: (error: any) => void;
    onOpen?: () => void;
    onClose?: () => void;
    onToolCall?: (toolCall: any) => Promise<any>;
  }) {
    if (this.isConnected) return;
    this.isDisposed = false;

    // Pre-initialize and resume AudioContext in user-gesture-safe thread for Android/iOS compatibility
    try {
      const AudioCtx = window.AudioContext || (window as any).webkitAudioContext;
      if (AudioCtx) {
        this.audioContext = new AudioCtx();
        if (this.outputDeviceId && this.outputDeviceId !== "default" && typeof (this.audioContext as any).setSinkId === 'function') {
          (this.audioContext as any).setSinkId(this.outputDeviceId).catch((e: any) => console.error("[Audio] Connect setSinkId failed:", e));
        }
        if (this.audioContext.state === 'suspended') {
          await this.audioContext.resume();
        }
        if (this.isDisposed) {
          console.log("[Audio] Connection aborted during AudioContext resumption.");
          this.disconnect();
          return;
        }
        console.log("AudioContext pre-initialized successfully in user-gesture-safe scope.");
      }
    } catch (e) {
      console.warn("Failed to pre-initialize AudioCtx in gesture scope:", e);
    }

    this.onVolumeChange = callbacks.onVolumeChange || this.onVolumeChange;
    this.onTranscript = callbacks.onTranscript || this.onTranscript;
    this.onInterrupted = callbacks.onInterrupted || this.onInterrupted;
    const onOpen = callbacks.onOpen || (() => {});
    const onClose = callbacks.onClose || (() => {});
    const onError = callbacks.onError || ((err) => console.error("Live API Error:", err));
    this.onToolCall = callbacks.onToolCall || this.onToolCall;

    if (this.isDisposed) {
      console.log("[Connect] Connection aborted before WebSocket handshake.");
      return;
    }

    // Capture exclusive active ownership for this specific LiveAPI instance to absolute-prevent duplication in the same or separate tabs
    safeLocalStorage.setItem("__active_tania_session_tab_id", tabSessionId);
    safeLocalStorage.setItem("__active_tania_live_api_instance_id", this.instanceId);
    if (this.concurrencyInterval) {
      clearInterval(this.concurrencyInterval);
    }
    this.concurrencyInterval = setInterval(() => {
      try {
        const activeId = safeLocalStorage.getItem("__active_tania_session_tab_id");
        const activeInstanceId = safeLocalStorage.getItem("__active_tania_live_api_instance_id");
        const globalActiveInstance = typeof window !== 'undefined' ? (window as any).__activeLiveApi : null;
        
        let shouldDisconnect = false;
        if (activeId && activeId !== tabSessionId) {
          shouldDisconnect = true;
          console.warn("[Concurrency Control] Stale browser tab context detected. Forcing clean disconnect...");
        } else if (activeInstanceId && activeInstanceId !== this.instanceId) {
          shouldDisconnect = true;
          console.warn("[Concurrency Control] New LiveAPI instance has overridden active role in this tab. Forcing clean disconnect to prevent dual voices...");
        } else if (globalActiveInstance && globalActiveInstance !== this) {
          shouldDisconnect = true;
          console.warn("[Concurrency Control] Stale memory instance overridden in window scope. Forcing clean disconnect...");
        }

        if (shouldDisconnect) {
          this.disconnect();
          onClose();
        }
      } catch (err) {}
    }, 400);

    try {
      console.log("Connecting client to server WebSocket proxy...");
      this.onTranscript("System: Negotiating connection with voice engine backend...");

      const protocol = window.location.protocol === "https:" ? "wss:" : "ws:";
      const wsUrl = `${protocol}//${window.location.host}/api/live`;
      this.ws = new WebSocket(wsUrl);

      this.ws.onopen = () => {
        if (this.isDisposed) {
          console.log("[Connect] WS opened but LiveAPI matches isDisposed. Hard collapsing.");
          try { this.ws?.close(); } catch (e) {}
          this.ws = null;
          return;
        }
        console.log("WebSocket connection to server proxy opened");
        // Handshake: initiate backend connection to Gemini Live API
        this.ws?.send(JSON.stringify({
          type: "connect",
          model: this.config.model,
          config: {
            systemInstruction: this.config.systemInstruction || "",
            tools: this.config.tools || [],
            responseModalities: [Modality.AUDIO],
            speechConfig: {
              voiceConfig: { prebuiltVoiceConfig: { voiceName: "Kore" } },
            },
            inputAudioTranscription: {},
            outputAudioTranscription: {},
          }
        }));

        // Keep connection alive with proxy/servers against 30-second idle rules
        this.pingInterval = setInterval(() => {
          if (this.ws && this.ws.readyState === WebSocket.OPEN) {
            this.ws.send(JSON.stringify({ type: "ping" }));
          }
        }, 15000); // Send heartbeat every 15 seconds
      };

      this.ws.onmessage = async (event) => {
        if (this.isDisposed) return;
        try {
          // Robust client-side concurrency safety: verify this instance is still active
          const activeId = safeLocalStorage.getItem("__active_tania_session_tab_id");
          const activeInstanceId = safeLocalStorage.getItem("__active_tania_live_api_instance_id");
          const globalActiveInstance = typeof window !== 'undefined' ? (window as any).__activeLiveApi : null;
          
          if (
            (activeId && activeId !== tabSessionId) ||
            (activeInstanceId && activeInstanceId !== this.instanceId) ||
            (globalActiveInstance && globalActiveInstance !== this)
          ) {
            console.warn("[Concurrency Control] Stale/superseded LiveAPI instance detected on legacy message packet. Hot-collapsing and muting immediately!");
            this.disconnect();
            onClose();
            return;
          }

          const message = JSON.parse(event.data);
          
          if (message.type === "open") {
            console.log("Live API voice gateway opened fully.");
            this.isConnected = true;
            await this.setupAudioSystem();
            onOpen();
          } else if (message.type === "close") {
            console.log("Server indicated voice connection closed");
            this.disconnect();
            onClose();
          } else if (message.type === "error") {
            console.error("Server-side setup error:", message.error);
            onError(new Error(message.error));
            this.disconnect();
          } else if (message.type === "message") {
            // Forward real-time content payload to parser
            await this.handleMessage(message.data);
          }
        } catch (parseErr) {
          console.error("Error parsing websocket packet from server:", parseErr);
        }
      };

      this.ws.onclose = () => {
        console.log("Proxy WebSocket closed");
        this.disconnect();
        onClose();
      };

      this.ws.onerror = (error) => {
        console.error("Proxy WebSocket error:", error);
        onError(new Error("Network error - failed to reach server gateway."));
        this.disconnect();
        onClose();
      };

    } catch (error) {
      console.error("Connect failure:", error);
      onError(error);
      throw error;
    }
  }

  private async handleMessage(message: any) {
    if (!message) return;
    console.log("Live API Message Processing:", Object.keys(message));
    
    // 1. Handle Model Turn (Audio + Text)
    const serverContent = message.serverContent;
    if (serverContent) {
      const modelTurn = serverContent.modelTurn;
      if (modelTurn?.parts) {
        for (const part of modelTurn.parts) {
          if (part.inlineData?.data) {
            this.handleAudioOutput(part.inlineData.data);
          }
          if (part.text && part.text.trim()) {
            console.log("Tania Text Part from modelTurn:", part.text.trim());
            this.onTranscript(`Tania: ${part.text.trim()}`);
          }
        }
      }

      // 2. Handle Audio Transcriptions
      const inputTranscription = serverContent.inputTranscription;
      if (inputTranscription?.text?.trim()) {
        console.log("Input Speech Transcription:", inputTranscription.text.trim());
        this.onTranscript(`Pesala: ${inputTranscription.text.trim()}`);
      }

      const outputTranscription = serverContent.outputTranscription;
      if (outputTranscription?.text?.trim()) {
        console.log("Output Speech Transcription:", outputTranscription.text.trim());
        this.onTranscript(`Tania: ${outputTranscription.text.trim()}`);
      }

      const audioTranscription = serverContent.audioTranscription;
      if (audioTranscription?.text?.trim()) {
        console.log("Legacy Audio Transcription:", audioTranscription.text.trim(), "Role:", audioTranscription.role);
        // Map roles correctly
        let role = audioTranscription.role;
        if (!role) {
          role = (this.isPlaying || this.audioQueue.length > 0) ? "model" : "user";
        }

        const prefix = (role === "user" || role === "USER") ? "Pesala: " : "Tania: ";
        this.onTranscript(`${prefix}${audioTranscription.text.trim()}`);
      }

      // 3. Handle User Content (Type-in chat results)
      const userContent = serverContent.userContent;
      if (userContent?.parts) {
        for (const part of userContent.parts) {
          if (part.text && part.text.trim()) {
            console.log("User Text Part from userContent:", part.text.trim());
            this.onTranscript(`Pesala: ${part.text.trim()}`);
          }
        }
      }

      if (serverContent.interrupted) {
        console.log("Tania Interrupted");
        this.stopPlayback();
        this.onInterrupted();
      }
    }

    // 4. Handle Tool Calls
    const toolCall = message.toolCall || serverContent?.toolCall;
    if (toolCall) {
      console.log("Tool Call detected:", toolCall);
      const functionResponses = [];
      for (const fc of toolCall.functionCalls || []) {
        try {
          const callId = fc.id || fc.functionCallId || fc.callId;
          console.log("Processing Tool Call:", fc.name, "ID:", callId);
          const response = await this.onToolCall(fc);
          functionResponses.push({
            id: callId,
            name: fc.name,
            response: { output: response }
          });
        } catch (err) {
          const callId = fc.id || fc.functionCallId || fc.callId;
          functionResponses.push({
            id: callId,
            name: fc.name,
            response: { output: { error: String(err) } }
          });
        }
      }
      if (functionResponses.length > 0 && this.ws && this.isConnected) {
        console.log("Sending Tool Responses:", functionResponses);
        this.ws.send(JSON.stringify({
          type: "tool_response",
          data: { functionResponses }
        }));
      }
    }
  }

  private async setupAudioSystem() {
    if (this.isDisposed) {
      console.log("[Audio] setupAudioSystem aborted because instance is already disposed.");
      return;
    }
    console.log("Setting up audio system...");
    this.inputAudioAccumulator = [];
    
    // Setup Microphone
    try {
      if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
        console.warn("Browser does not support getUserMedia. Proceeding without microphone.");
      } else {
        try {
          const devices = await navigator.mediaDevices.enumerateDevices();
          if (this.isDisposed) return;
          const hasMic = devices.some(d => d.kind === 'audioinput');
          console.log("Available audio input devices:", devices.filter(d => d.kind === 'audioinput'));
          
          if (!hasMic) {
            console.warn("No audio input devices found via enumerateDevices.");
          }
        } catch (e) {
          console.warn("Could not enumerate devices:", e);
        }

        if (this.isDisposed) return;

        try {
          this.stream = await navigator.mediaDevices.getUserMedia({ 
            audio: {
              echoCancellation: true,
              noiseSuppression: true,
              autoGainControl: true
            } 
          });
        } catch (err: any) {
          if (this.isDisposed) return;
          console.warn("Microphone with full constraints failed, retrying with basic audio:true...", err);
          try {
            this.stream = await navigator.mediaDevices.getUserMedia({ audio: true });
          } catch (basicErr: any) {
            console.warn("Microphone capture unavailable or not attached:", basicErr);
            this.onTranscript("System: Microphone not detected or permission denied. You can still stream Tania's voice responses and type to chat!");
          }
        }
      }
    } catch (err: any) {
      if (this.isDisposed) return;
      console.warn("Microphone Access Setup Warning:", err);
      this.onTranscript("System: Microphone not found or access denied. You can still type to chat.");
    }

    if (this.isDisposed || !this.ws || this.ws.readyState !== WebSocket.OPEN) {
      if (this.stream) {
        this.stream.getTracks().forEach(track => track.stop());
        this.stream = null;
      }
      return;
    }

    // Reuse or create AudioContext
    if (!this.audioContext) {
      const AudioCtx = window.AudioContext || (window as any).webkitAudioContext;
      if (AudioCtx) {
        // Create with native hardware default sampleRate for superior stability, responsiveness, and to avoid double-resampling glitches
        this.audioContext = new AudioCtx();
        if (this.outputDeviceId && this.outputDeviceId !== "default" && typeof (this.audioContext as any).setSinkId === 'function') {
          (this.audioContext as any).setSinkId(this.outputDeviceId).catch((e: any) => console.error("[Audio] Setup setSinkId failed:", e));
        }
      }
    }
    
    if (this.audioContext && this.audioContext.state === 'suspended') {
      try {
        await this.audioContext.resume();
      } catch (err) {
        console.error("Failed to resume AudioContext:", err);
      }
    }

    if (this.isDisposed || !this.ws || this.ws.readyState !== WebSocket.OPEN) {
      console.log("[Audio] Aborting audio context setup because WS is inactive or instance is disposed");
      if (this.stream) {
        this.stream.getTracks().forEach(track => track.stop());
        this.stream = null;
      }
      if (this.audioContext) {
        if (this.audioContext.state !== 'closed') {
          this.audioContext.close().catch(e => console.error("Error closing AudioContext on safe dispose:", e));
        }
        this.audioContext = null;
      }
      return;
    }

    // Setup Analyser for lip-sync / volume rendering
    this.analyser = this.audioContext!.createAnalyser();
    this.analyser.fftSize = 256;
    this.analyser.smoothingTimeConstant = 0.1; 
    this.analyser.minDecibels = -90;
    this.analyser.maxDecibels = -10;
    this.analyser.connect(this.audioContext!.destination);

    // Recording processor
    if (this.stream) {
      const source = this.audioContext!.createMediaStreamSource(this.stream);
      this.processor = this.audioContext!.createScriptProcessor(2048, 1, 1);
      
      this.processor.onaudioprocess = (e) => {
        if (!this.isConnected || !this.ws) return;
        
        try {
          const inputData = e.inputBuffer.getChannelData(0);
          const currentSampleRate = e.inputBuffer.sampleRate;
          
          const downsampledData = this.downsampleBuffer(inputData, currentSampleRate, 16000);
          
          // Accumulate downsampled samples into stable 100ms blocks (1600 samples at 16000Hz).
          // This keeps packet frequency constant and low (10 packets/second), completely avoiding network queuing and reducing latency to sub-second real-time speeds!
          for (let i = 0; i < downsampledData.length; i++) {
            this.inputAudioAccumulator.push(downsampledData[i]);
          }

          const CHUNK_SIZE = 1600; // 100ms of audio at 16k
          while (this.inputAudioAccumulator.length >= CHUNK_SIZE) {
            const chunk = new Float32Array(this.inputAudioAccumulator.slice(0, CHUNK_SIZE));
            this.inputAudioAccumulator = this.inputAudioAccumulator.slice(CHUNK_SIZE);

            let maxVal = 0;
            for (let i = 0; i < chunk.length; i++) {
              const abs = Math.abs(chunk[i]);
              if (abs > maxVal) maxVal = abs;
            }

            // Determine if Tania is currently speaking or just finished speaking (within a robust cushion to defeat acoustic echo completely).
            // When Tania is speaking, has recently spoken, or has scheduled audio in the Web Audio context queue,
            // we mute transmission to completely eliminate loudspeaker echo feedback loops and resolve dual voice issues!
            const isTaniaPlaying = (this.playingSources.length > 0) ||
                                   (!!this.audioContext && this.audioContext.currentTime < this.nextStreamTime) ||
                                   (Date.now() - this.lastTaniaSpeechTime < 1500);

            // Dynamic Echo Cancellation and Ambient Noise Gate setup
            // When Tania is playing, we enforce an exceptionally high threshold to ignore feedback bleeding.
            const threshold = isTaniaPlaying ? 0.95 : 0.015;

            if (isTaniaPlaying) {
              // Strictly mute and clear any user voice activity during playback
              this.isAudioActive = false;
              this.silentFramesCount = 0;
              this.sentSilenceFramesAfterSpeechCount = 0;
            } else {
              if (maxVal > threshold) {
                this.isAudioActive = true;
                this.silentFramesCount = 0;
              } else {
                this.silentFramesCount++;
                // With 100ms frames, 3 silent frames is exactly 300ms of quiet before closing the microphone speaking gate.
                if (this.silentFramesCount >= 3) {
                  this.isAudioActive = false;
                }
              }
            }

            // Speaking status depends on physical voice activity
            this.isUserSpeaking = this.isAudioActive;

            // Determine if we should send this block.
            // Only send active speech blocks or up to 3 padding silence blocks after active speech blocks to prevent word clipping.
            let shouldSend = false;
            let dataToSend = chunk;

            if (this.isAudioActive && !isTaniaPlaying) {
              shouldSend = true;
              this.sentSilenceFramesAfterSpeechCount = 3; // Reset hangover frames tickets
            } else {
              if (this.sentSilenceFramesAfterSpeechCount > 0 && !isTaniaPlaying) {
                shouldSend = true;
                dataToSend = new Float32Array(CHUNK_SIZE); // Send pristine digital silence for trailing frame
                this.sentSilenceFramesAfterSpeechCount--;
              }
            }

            if (shouldSend) {
              const pcmData = this.float32ToInt16(dataToSend);
              const base64Data = this.arrayBufferToBase64(pcmData.buffer);

              this.ws.send(JSON.stringify({
                type: "input",
                data: {
                  audio: { data: base64Data, mimeType: 'audio/pcm;rate=16000' }
                }
              }));
            }
          }
        } catch (err) {
          console.error("Error sending audio input:", err);
        }
      };

      source.connect(this.processor);
      // Route microphone capture through a silent gain node to force ScriptProcessorNode activity without bleeding feedback to speakers!
      const silenceGain = this.audioContext!.createGain();
      silenceGain.gain.value = 0;
      this.processor.connect(silenceGain);
      silenceGain.connect(this.audioContext!.destination);
    }

    this.startVolumeTracking();
  }

  private startVolumeTracking() {
    if (!this.analyser) return;
    const dataArray = new Float32Array(this.analyser.frequencyBinCount);
    const track = () => {
      if (!this.isConnected || !this.analyser) return;
      
      this.analyser.getFloatTimeDomainData(dataArray);
      let sum = 0;
      for (let i = 0; i < dataArray.length; i++) {
        sum += dataArray[i] * dataArray[i];
      }
      const rms = Math.sqrt(sum / dataArray.length);
      
      this.onVolumeChange(rms);
      this.volumeInterval = requestAnimationFrame(track);
    };
    this.volumeInterval = requestAnimationFrame(track);
  }

  private handleAudioOutput(base64: string) {
    if (this.isDisposed || !this.isConnected || !this.ws || this.ws.readyState !== WebSocket.OPEN) return;
    
    // Concurrency guard: verify we are still the active window session & instance
    const activeId = safeLocalStorage.getItem("__active_tania_session_tab_id");
    const activeInstanceId = safeLocalStorage.getItem("__active_tania_live_api_instance_id");
    if ((activeId && activeId !== tabSessionId) || (activeInstanceId && activeInstanceId !== this.instanceId)) {
      console.warn("[Audio Control] Newer/different LiveAPI instance detected during audio playback. Auto-disconnecting and muting voice to prevent overlapping audio.");
      this.disconnect();
      return;
    }

    const audioData = this.base64ToFloat32(base64);
    
    if (!this.audioContext || !this.analyser) return;

    // Resiliency: ensure the audio context is active and not suspended by browser idle policies
    if (this.audioContext.state === "suspended") {
      this.audioContext.resume().catch((err) => console.error("[Audio] Auto-resuming suspended Context failed:", err));
    }

    // Mutex stop: immediately query the global array for other instances' active buffer sources and clean them up
    if ((window as any).__allTaniaAudioSources) {
      (window as any).__allTaniaAudioSources = (window as any).__allTaniaAudioSources.filter((s: any) => {
        if (s._instanceId && s._instanceId !== this.instanceId) {
          try {
            s.stop();
          } catch (e) {}
          return false;
        }
        return true;
      });
    } else {
      (window as any).__allTaniaAudioSources = [];
    }

    const buffer = this.audioContext.createBuffer(1, audioData.length, 24000);
    buffer.getChannelData(0).set(audioData);

    const source = this.audioContext.createBufferSource();
    source.buffer = buffer;
    (source as any)._instanceId = this.instanceId;
    (window as any).__allTaniaAudioSources.push(source);

    // Connect audio source directly to frequency analyser & speakers for peak fidelity, maximum clarity, and zero cracking!
    source.connect(this.analyser);

    const currentTime = this.audioContext.currentTime;
    const isMobile = /Android|iPhone|iPad/i.test(navigator.userAgent);
    
    // Professional tight audio buffering configuration: snappier response, smoother transitions, zero latency breaks!
    const initialCushion = isMobile ? 0.16 : 0.06;
    const minCushion = isMobile ? 0.035 : 0.012; // Extremely fast stream recovery window

    let playTime = this.nextStreamTime;

    // Self-healing playhead timeline scheduler:
    // To prevent Web Audio API from cutting off or crackling, the playTime must be strictly in the future.
    // If the playTime falls below current time + minCushion, we absorb the jitter dynamically.
    if (playTime < currentTime + minCushion) {
      const gap = currentTime - playTime;
      // If it is minor timing delay, bridge the gaps with a safe future cushion to heal the stream instantly.
      if (playTime > 0 && gap <= 0.45) {
        playTime = currentTime + minCushion;
      } else {
        // Fresh start or major network delay: schedule with a robust start cushion
        playTime = currentTime + initialCushion;
      }
    } else if (playTime > currentTime + 15.0) {
      // Protection against custom buffer accumulation / extreme delay drift: resync queue timeline with a clean cut to prevent dual voices
      console.warn(`[Audio] Extreme queue delay drift/accumulation detected (${(playTime - currentTime).toFixed(2)}s). Resuming with stable cushion after a clean playback stop.`);
      this.stopPlayback();
      playTime = currentTime + initialCushion;
    }

    const duration = buffer.duration;

    // Schedule the source to start playing precisely at playTime (playTime is guaranteed > currentTime + minCushion)
    source.start(playTime);
    
    // Track active source to know if Tania is playing
    this.playingSources.push(source);
    this.lastTaniaSpeechTime = Date.now();
    
    source.onended = () => {
      this.playingSources = this.playingSources.filter(s => s !== source);
      if ((window as any).__allTaniaAudioSources) {
        (window as any).__allTaniaAudioSources = (window as any).__allTaniaAudioSources.filter((s: any) => s !== source);
      }
    };

    // Update nextStreamTime to point to the virtual end of this buffer
    this.nextStreamTime = playTime + duration;
  }

  private stopPlayback() {
    if ((window as any).__allTaniaAudioSources) {
      (window as any).__allTaniaAudioSources = (window as any).__allTaniaAudioSources.filter((s: any) => {
        if (s._instanceId === this.instanceId) {
          try {
            s.stop();
          } catch (e) {}
          return false;
        }
        return true;
      });
    }
    this.playingSources.forEach(source => {
      try {
        source.stop();
      } catch (err) {
        // State error safe
      }
    });
    this.playingSources = [];
    this.audioQueue = [];
    this.isPlaying = false;
    this.nextStreamTime = 0;
  }

  private downsampleBuffer(buffer: Float32Array, inputSampleRate: number, outputSampleRate: number = 16000): Float32Array {
    if (inputSampleRate === outputSampleRate) {
      return buffer;
    }
    const sampleRateRatio = inputSampleRate / outputSampleRate;
    const newLength = Math.round(buffer.length / sampleRateRatio);
    const result = new Float32Array(newLength);
    let offsetResult = 0;
    let offsetBuffer = 0;
    while (offsetResult < result.length) {
      const nextOffsetBuffer = Math.round((offsetResult + 1) * sampleRateRatio);
      let accum = 0;
      let count = 0;
      for (let i = offsetBuffer; i < nextOffsetBuffer && i < buffer.length; i++) {
        accum += buffer[i];
        count++;
      }
      result[offsetResult] = count > 0 ? accum / count : 0;
      offsetResult++;
      offsetBuffer = nextOffsetBuffer;
    }
    return result;
  }

  private float32ToInt16(float32: Float32Array): Int16Array {
    const int16 = new Int16Array(float32.length);
    for (let i = 0; i < float32.length; i++) {
      const s = Math.max(-1, Math.min(1, float32[i]));
      int16[i] = s < 0 ? s * 0x8000 : s * 0x7FFF;
    }
    return int16;
  }

  private base64ToFloat32(base64: string): Float32Array {
    const binary = atob(base64);
    const len = binary.length;
    const numSamples = Math.floor(len / 2);
    const float32 = new Float32Array(numSamples);
    
    for (let i = 0; i < numSamples; i++) {
      const low = binary.charCodeAt(i * 2);
      const high = binary.charCodeAt(i * 2 + 1);
      let val = low | (high << 8);
      if (val & 0x8000) {
        val |= ~0xffff;
      }
      float32[i] = val / 32768.0;
    }
    return float32;
  }

  private arrayBufferToBase64(buffer: ArrayBuffer): string {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    for (let i = 0; i < bytes.byteLength; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
  }

  sendText(text: string) {
    if (this.ws && this.isConnected) {
      this.ws.send(JSON.stringify({ type: "text", text }));
    }
  }

  disconnect() {
    this.isDisposed = true;

    // Immediately stop any actively scheduled Audio Buffer Sources in the browser to prevent overlapping audio
    this.stopPlayback();

    if (this.concurrencyInterval) {
      clearInterval(this.concurrencyInterval);
      this.concurrencyInterval = null;
    }

    if (this.pingInterval) {
      clearInterval(this.pingInterval);
      this.pingInterval = null;
    }

    this.isConnected = false;
    this.isPlaying = false;
    this.audioQueue = [];
    this.inputAudioAccumulator = [];

    if (this.ws) {
      try {
        this.ws.close();
      } catch (e) {
        console.error("Error closing client WebSocket:", e);
      }
      this.ws = null;
    }
    if (this.stream) {
      try {
        this.stream.getTracks().forEach(track => track.stop());
      } catch (e) {}
      this.stream = null;
    }
    if (this.processor) {
      try {
        this.processor.disconnect();
      } catch (e) {}
      this.processor = null;
    }
    if (this.volumeInterval) {
      cancelAnimationFrame(this.volumeInterval);
      this.volumeInterval = null;
    }
    if (this.analyser) {
      try {
        this.analyser.disconnect();
      } catch (e) {}
      this.analyser = null;
    }
    if (this.audioContext) {
      try {
        if (this.audioContext.state !== 'closed') {
          this.audioContext.close().catch(e => console.error("Error closing AudioContext:", e));
        }
      } catch (e) {}
      this.audioContext = null;
    }
  }

  static disconnectAll() {
    console.log("[Global Cleanup] Cleaning and stopping all global Tania audio playing sources...");
    if ((window as any).__allTaniaAudioSources) {
      (window as any).__allTaniaAudioSources.forEach((s: any) => {
        try {
          s.stop();
        } catch (e) {}
      });
      (window as any).__allTaniaAudioSources = [];
    }

    const apis = (window as any).__allLiveApis || [];
    console.log(`[Global Cleanup] Disconnecting all ${apis.length} registered LiveAPI instances...`);
    apis.forEach((api: any) => {
      try {
        api.disconnect();
      } catch (e) {
        console.warn("Error disconnecting registered LiveAPI instance:", e);
      }
    });
    (window as any).__allLiveApis = [];
  }
}
