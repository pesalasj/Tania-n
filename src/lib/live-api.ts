import { GoogleGenAI, Modality } from "@google/genai";

export interface LiveAPIConfig {
  apiKey: string;
  model: string;
  systemInstruction?: string;
  tools?: any[];
}

export class LiveAPI {
  private ai: any;
  private session: any;
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

  constructor(private config: LiveAPIConfig) {
    if (!config.apiKey) {
      throw new Error("An API Key must be set for LiveAPI to function.");
    }
    // Initialize SDK
    this.ai = new GoogleGenAI({ 
      apiKey: config.apiKey, 
      apiVersion: "v1beta",
      vertexai: false
    });
    
    // The SDK sometimes defaults to "v1main" internally for WebSockets.
    // We need to force it to use "v1beta" by patching the apiClient.
    if (this.ai.apiClient) {
      this.ai.apiClient.getApiVersion = () => "v1beta";
      
      // If there's an internal clientOptions, force it there too
      if (this.ai.apiClient.clientOptions) {
        this.ai.apiClient.clientOptions.apiVersion = "v1beta";
      }
    }

    console.log("Initialized GoogleGenAI forced to v1beta");
    if (this.ai.apiClient) {
      console.log("SDK ApiClient resolved apiVersion:", this.ai.apiClient.getApiVersion());
      console.log("SDK ApiClient resolved websocketBaseUrl:", this.ai.apiClient.getWebsocketBaseUrl());
    }
  }

  private analyser: AnalyserNode | null = null;
  private volumeInterval: any = null;

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

    // Pre-initialize and resume AudioContext in user-gesture thread for Android/iOS compatibility
    try {
      const AudioCtx = window.AudioContext || (window as any).webkitAudioContext;
      if (AudioCtx) {
        this.audioContext = new AudioCtx({ sampleRate: 16000 });
        if (this.audioContext.state === 'suspended') {
          await this.audioContext.resume();
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

    try {
      console.log("Connecting to Live API with model:", this.config.model);
      const modelName = this.config.model.startsWith("models/") ? this.config.model : `models/${this.config.model}`;
      
      this.onTranscript("System: Negotiating connection with voice engine...");
      
      this.session = await this.ai.live.connect({
        model: modelName,
        config: {
          systemInstruction: { parts: [{ text: this.config.systemInstruction || "" }] },
          tools: this.config.tools || [{ googleSearch: {} }],
          responseModalities: [Modality.AUDIO],
          speechConfig: {
            voiceConfig: { prebuiltVoiceConfig: { voiceName: "Aoede" } },
          },
          inputAudioTranscription: {},
          outputAudioTranscription: {},
        },
        callbacks: {
          onopen: async () => {
            console.log("Live API Session Opened");
            this.isConnected = true;
            await this.setupAudioSystem();
            onOpen();
          },
          onmessage: async (message: any) => {
            console.log("Live API Message Received:", JSON.stringify(message).slice(0, 200) + "...");
            await this.handleMessage(message);
          },
          onclose: () => {
            console.log("Live API Session Closed");
            this.disconnect();
            onClose();
          },
          onerror: (error: any) => {
            console.error("Live API Session Error:", error);
            this.isConnected = false;
            let msg = "Network error - please check your API key";
            if (error?.message) msg = error.message;
            onError(new Error(msg));
            this.disconnect();
            onClose();
          },
        },
      });

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

      // 2. Handle Audio Transcriptions (supporting both modern and legacy fields)
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
          // Fallback logic if role is missing
          role = (this.isPlaying || this.audioQueue.length > 0) ? "model" : "user";
        }

        // Standardize roles to Tania/Pesala prefixes
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

    // 4. Handle Tool Calls (At root or inside serverContent)
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
      if (functionResponses.length > 0) {
        console.log("Sending Tool Responses:", functionResponses);
        this.session.sendToolResponse({ functionResponses });
      }
    }
  }

  private async setupAudioSystem() {
    console.log("Setting up audio system...");
    
    // Setup Microphone
    try {
      if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
        console.warn("Browser does not support getUserMedia. Proceeding without microphone.");
      } else {
        // Debug: List devices to see if any mic exists
        try {
          const devices = await navigator.mediaDevices.enumerateDevices();
          const hasMic = devices.some(d => d.kind === 'audioinput');
          console.log("Available audio input devices:", devices.filter(d => d.kind === 'audioinput'));
          
          if (!hasMic) {
            console.warn("No audio input devices found via enumerateDevices.");
          }
        } catch (e) {
          console.warn("Could not enumerate devices:", e);
        }

        try {
          this.stream = await navigator.mediaDevices.getUserMedia({ 
            audio: {
              echoCancellation: true,
              noiseSuppression: true,
              autoGainControl: true
            } 
          });
        } catch (err: any) {
          console.warn("Microphone with full constraints failed, retrying with basic audio:true...", err);
          this.stream = await navigator.mediaDevices.getUserMedia({ audio: true });
        }
      }
    } catch (err: any) {
      console.error("Microphone Access Error:", err);
      // We do not throw here, so the user can still use text input and hear output
      this.onTranscript("System: Microphone not found or access denied. You can still type to chat.");
    }

    // Reuse pre-initialized context or create a new one as fallback
    if (!this.audioContext) {
      const AudioCtx = window.AudioContext || (window as any).webkitAudioContext;
      if (AudioCtx) {
        try {
          this.audioContext = new AudioCtx({ sampleRate: 16000 });
        } catch (err) {
          console.warn("Fallback to default sampleRate in manual context creation:", err);
          this.audioContext = new AudioCtx();
        }
      }
    }
    
    // Ensure context is running (browsers often suspend it)
    if (this.audioContext && this.audioContext.state === 'suspended') {
      try {
        await this.audioContext.resume();
      } catch (err) {
        console.error("Failed to resume AudioContext:", err);
      }
    }

    // Setup Analyser for output lip-sync (hearing Tania)
    this.analyser = this.audioContext.createAnalyser();
    this.analyser.fftSize = 256;
    this.analyser.smoothingTimeConstant = 0.1; 
    this.analyser.minDecibels = -90;
    this.analyser.maxDecibels = -10;
    this.analyser.connect(this.audioContext.destination);

    // Only setup recording processor if we have a stream
    if (this.stream) {
      const source = this.audioContext.createMediaStreamSource(this.stream);
      this.processor = this.audioContext.createScriptProcessor(4096, 1, 1);
      
      this.processor.onaudioprocess = (e) => {
        if (!this.isConnected || !this.session) return;
        
        try {
          const inputData = e.inputBuffer.getChannelData(0);
          const currentSampleRate = e.inputBuffer.sampleRate;
          
          // Downsample high rate mic to 16000 for server-side compatibility
          const downsampledData = this.downsampleBuffer(inputData, currentSampleRate, 16000);
          
          // Check if there's actual audio signal (not just silence)
          let maxVal = 0;
          for (let i = 0; i < downsampledData.length; i++) {
            const abs = Math.abs(downsampledData[i]);
            if (abs > maxVal) maxVal = abs;
          }
          this.isUserSpeaking = maxVal > 0.001;

          const pcmData = this.float32ToInt16(downsampledData);
          const base64Data = this.arrayBufferToBase64(pcmData.buffer);
          
          this.session.sendRealtimeInput({
            audio: { data: base64Data, mimeType: 'audio/pcm;rate=16000' }
          });
        } catch (err) {
          console.error("Error sending audio input:", err);
        }
      };

      source.connect(this.processor);
      this.processor.connect(this.audioContext.destination);
    }

    // Start volume tracking loop
    this.startVolumeTracking();
  }

  private startVolumeTracking() {
    const dataArray = new Float32Array(this.analyser!.frequencyBinCount);
    const track = () => {
      if (!this.isConnected || !this.analyser) return;
      
      this.analyser.getFloatTimeDomainData(dataArray);
      let sum = 0;
      for (let i = 0; i < dataArray.length; i++) {
        sum += dataArray[i] * dataArray[i];
      }
      const rms = Math.sqrt(sum / dataArray.length);
      
      // Debug: console.log("RMS Volume:", rms);
      this.onVolumeChange(rms);
      this.volumeInterval = requestAnimationFrame(track);
    };
    this.volumeInterval = requestAnimationFrame(track);
  }

  private handleAudioOutput(base64: string) {
    const audioData = this.base64ToFloat32(base64);
    
    if (!this.audioContext || !this.analyser) return;

    const buffer = this.audioContext.createBuffer(1, audioData.length, 24000);
    buffer.getChannelData(0).set(audioData);

    const source = this.audioContext.createBufferSource();
    source.buffer = buffer;
    source.connect(this.analyser);

    const currentTime = this.audioContext.currentTime;
    
    // Smooth sample-accurate pipeline scheduling optimized for Android/mobile devices
    const isMobile = /Android|iPhone|iPad/i.test(navigator.userAgent);
    const schedulingOffset = isMobile ? 0.15 : 0.05; // 150ms on mobile covers Android bluetooth/thread latency, 50ms on desktop

    if (this.nextStreamTime < currentTime) {
      this.nextStreamTime = currentTime + schedulingOffset;
    }

    source.start(this.nextStreamTime);
    
    this.playingSources.push(source);
    source.onended = () => {
      this.playingSources = this.playingSources.filter(s => s !== source);
    };

    // Increment start time by precise buffer duration
    this.nextStreamTime += buffer.duration;
  }

  private stopPlayback() {
    this.playingSources.forEach(source => {
      try {
        source.stop();
      } catch (err) {
        // Handle source states gracefully
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

  private analyzeVolume(data: Float32Array) {
    let sum = 0;
    for (let i = 0; i < data.length; i++) {
      sum += data[i] * data[i];
    }
    const rms = Math.sqrt(sum / data.length);
    this.onVolumeChange(rms);
  }

  sendText(text: string) {
    if (this.session && this.isConnected) {
      this.session.sendRealtimeInput({ text });
    }
  }

  disconnect() {
    if (!this.isConnected && !this.session && !this.audioContext) return;

    this.isConnected = false;
    this.isPlaying = false;
    this.audioQueue = [];

    if (this.session) {
      try {
        this.session.close();
      } catch (e) {
        console.error("Error closing session:", e);
      }
      this.session = null;
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
}
