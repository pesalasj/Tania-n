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
  private playingSources: AudioBufferSourceNode[] = [];
  private analyser: AnalyserNode | null = null;
  private volumeInterval: any = null;
  private pingInterval: any = null;
  private isDisposed = false;
  public outputDeviceId = "default";

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
          this.stream = await navigator.mediaDevices.getUserMedia({ audio: true });
        }
      }
    } catch (err: any) {
      if (this.isDisposed) return;
      console.error("Microphone Access Error:", err);
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
      this.processor = this.audioContext!.createScriptProcessor(4096, 1, 1);
      
      this.processor.onaudioprocess = (e) => {
        if (!this.isConnected || !this.ws) return;
        
        try {
          const inputData = e.inputBuffer.getChannelData(0);
          const currentSampleRate = e.inputBuffer.sampleRate;
          
          const downsampledData = this.downsampleBuffer(inputData, currentSampleRate, 16000);
          
          let maxVal = 0;
          for (let i = 0; i < downsampledData.length; i++) {
            const abs = Math.abs(downsampledData[i]);
            if (abs > maxVal) maxVal = abs;
          }
          this.isUserSpeaking = maxVal > 0.001;

          const pcmData = this.float32ToInt16(downsampledData);
          const base64Data = this.arrayBufferToBase64(pcmData.buffer);
          
          this.ws.send(JSON.stringify({
            type: "input",
            data: {
              audio: { data: base64Data, mimeType: 'audio/pcm;rate=16000' }
            }
          }));
        } catch (err) {
          console.error("Error sending audio input:", err);
        }
      };

      source.connect(this.processor);
      this.processor.connect(this.audioContext!.destination);
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
    const audioData = this.base64ToFloat32(base64);
    
    if (!this.audioContext || !this.analyser) return;

    const buffer = this.audioContext.createBuffer(1, audioData.length, 24000);
    buffer.getChannelData(0).set(audioData);

    const source = this.audioContext.createBufferSource();
    source.buffer = buffer;

    // Create a local GainNode per audio chunk to execute sub-millisecond edge fading/crossfading
    // This removes high-frequency click bursts at slice/packet boundaries
    const chunkGain = this.audioContext.createGain();
    
    // Connect track source -> gain -> primary frequency analyser
    source.connect(chunkGain);
    chunkGain.connect(this.analyser);

    const currentTime = this.audioContext.currentTime;
    const isMobile = /Android|iPhone|iPad/i.test(navigator.userAgent);
    
    // Maintain a safe/protective lookahead cushion to smooth over brief network arrival delays and avoid word-to-word choppiness
    const schedulingOffset = isMobile ? 0.12 : 0.05;

    // Dynamic latency adjustment:
    if (this.nextStreamTime < currentTime) {
      // Underflow: The queue finished or is empty, start fresh from the lookahead offset
      this.nextStreamTime = currentTime + schedulingOffset;
    } else if (this.nextStreamTime > currentTime + 3.0) {
      // Extreme overflow/drift: Clear all active scheduled sources to prevent overlapping/dual voices
      console.warn(`[Audio] Overflow/drift threshold exceeded (${(this.nextStreamTime - currentTime).toFixed(2)}s). Resetting audio scheduler.`);
      this.playingSources.forEach(s => {
        try {
          s.stop();
        } catch (e) {}
      });
      this.playingSources = [];
      this.nextStreamTime = currentTime + schedulingOffset;
    }

    const duration = buffer.duration;
    const playTime = this.nextStreamTime;

    // Apply 3ms input-and-output micro-fades to blend discontinuous PCM slices with perfect warmth
    const fadeDuration = Math.min(0.003, duration / 2.1);
    
    chunkGain.gain.setValueAtTime(0, playTime);
    chunkGain.gain.linearRampToValueAtTime(1, playTime + fadeDuration);
    chunkGain.gain.setValueAtTime(1, playTime + duration - fadeDuration);
    chunkGain.gain.linearRampToValueAtTime(0, playTime + duration);

    source.start(playTime);
    
    this.playingSources.push(source);
    source.onended = () => {
      this.playingSources = this.playingSources.filter(s => s !== source);
    };

    this.nextStreamTime += duration;
  }

  private stopPlayback() {
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
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    const int16 = new Int16Array(bytes.buffer);
    const float32 = new Float32Array(int16.length);
    for (let i = 0; i < int16.length; i++) {
      float32[i] = int16[i] / 32768.0;
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
    if (!this.isConnected && !this.ws && !this.audioContext) return;

    if (this.pingInterval) {
      clearInterval(this.pingInterval);
      this.pingInterval = null;
    }

    this.isConnected = false;
    this.isPlaying = false;
    this.audioQueue = [];

    if (this.ws) {
      try {
        this.ws.close();
      } catch (e) {
        console.error("Error closing client WebSocket:", e);
      }
      this.ws = null;
    }
    if (this.stream) {
      this.stream.getTracks().forEach(track => track.stop());
      this.stream = null;
    }
    if (this.processor) {
      this.processor.disconnect();
      this.processor = null;
    }
    if (this.volumeInterval) {
      cancelAnimationFrame(this.volumeInterval);
      this.volumeInterval = null;
    }
    if (this.analyser) {
      this.analyser.disconnect();
      this.analyser = null;
    }
    if (this.audioContext) {
      if (this.audioContext.state !== 'closed') {
        this.audioContext.close().catch(e => console.error("Error closing AudioContext:", e));
      }
      this.audioContext = null;
    }
  }

  static disconnectAll() {
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
