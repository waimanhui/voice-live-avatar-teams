# Voice Live Avatar with Teams Integration

> Based on: https://github.com/microsoft-foundry/voicelive-samples/tree/main/python/voice-live-avatar

A Python (FastAPI) + HTML/CSS/JavaScript sample that demonstrates the **Azure AI Voice Live API** with an animated avatar and optional Microsoft Teams meeting integration. The Voice Live SDK runs entirely on the server; the browser handles UI, audio capture/playback, avatar video rendering, and Teams calling.

## Architecture

```
┌────────────────────────────┐         ┌─────────────────────────┐         ┌──────────────────┐
│     Browser (Frontend)     │◄──WS───►│  Python Server (FastAPI)│◄──SDK──►│ Azure Voice Live │
│                            │         │                         │         │     Service      │
│  • Mic capture (PCM16)     │         │  • Session management   │         └──────────────────┘
│  • Audio playback (PCM16)  │         │  • SDK calls            │                  │
│  • Avatar video (WebRTC)   │◄──WebRTC (peer-to-peer)───────────────────────────────┘
│  • Settings / chat UI      │         │  • Event relay          │
│  • Teams ACS calling       │         │  • SDP relay            │
└──────────────┬─────────────┘         │  • ACS token minting    │
               │ ACS SDK               └─────────────────────────┘
               ▼
     ┌──────────────────┐
     │  Teams Meeting   │  ← AI avatar joins as a participant
     │  (via ACS SDK)   │    bidirectional audio via Web Audio API
     └──────────────────┘
```

**Key design points:**

- The Python backend bridges the browser WebSocket to the Azure Voice Live SDK. All SDK operations (session lifecycle, audio forwarding, event processing) happen in Python.
- Avatar video streams **peer-to-peer** from Azure directly to the browser via WebRTC — the Python server only relays the SDP offer/answer.
- Teams integration runs **entirely in the browser** using the Azure Communication Services (ACS) Calling SDK and the Web Audio API. The Python server only mints a short-lived ACS token.

---

## Prerequisites

- **Python 3.10+**
- An **Azure AI Services** resource in a [supported region](#avatar-supported-regions)
- An **Azure Communication Services** resource (required only for Teams integration)

### Avatar supported regions

The avatar feature is available in: Southeast Asia, North Europe, West Europe, Sweden Central, South Central US, East US 2, West US 2.

---

## Setup

### 1. Install Python dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure environment variables

Copy the template and fill in your values:

```powershell
Copy-Item .env.template .env
```

```env
# Required
AZURE_VOICELIVE_ENDPOINT=https://<your-resource>.cognitiveservices.azure.com/
AZURE_VOICELIVE_API_KEY=<your-api-key>

# Optional defaults (pre-filled in the UI)
VOICELIVE_MODEL=gpt-4o-realtime
VOICELIVE_VOICE=en-US-AvaMultilingualNeural

# Required only for Teams calling
AZURE_COMMUNICATION_CONNECTION_STRING=endpoint=https://<acs-resource>.communication.azure.com/;accesskey=<key>
TEAMS_MEETING_LINK=https://teams.microsoft.com/l/meetup-join/...
```

> All `.env` values are automatically pre-filled in the browser UI on page load via `/api/config`.

### 3. Run the server

```bash
uvicorn app:app --host 0.0.0.0 --port 3000 --reload
```

### 4. Open the browser

Navigate to [http://localhost:3000](http://localhost:3000).

---

## Docker

```bash
# Build
docker build -t voice-live-avatar-python .

# Run
docker run --rm -p 3000:3000 voice-live-avatar-python
```

Then open [http://localhost:3000](http://localhost:3000).

---

## Using the sample

### Connect

1. Under **Connection Settings**, enter your **Azure AI Services Endpoint** and **Subscription Key** (or Entra token for agent modes).
2. *(Optional)* Configure avatar, voice, and conversation settings (see below).
3. Click **Connect**. Once connected, click **Turn on microphone** and start speaking.

### Settings overview

| Section | Key fields |
|---|---|
| **Connection Settings** | Endpoint, Subscription Key / Entra Token, Mode (model vs agent), Model |
| **Conversation Settings** | Speech recognition model, turn detection, noise suppression, echo cancellation |
| **Voice Configuration** | Voice name, type (standard / custom / personal), speed, temperature |
| **Avatar Configuration** | Enable avatar, type (prebuilt / photo / custom), output mode (WebRTC / WebSocket), background image URL |
| **Scene Settings** | Photo-avatar only: Zoom, Position X/Y, Rotation X/Y/Z, Amplitude — adjustable live while connected |
| **Teams Integration** | Meeting link, display name, audio source, send avatar video toggle |

### Developer mode

Toggle **Developer Mode** (top of page) to reveal the chat transcript, system log messages, and a text input for sending messages without the microphone.

---

## Teams Integration

The sample can join a Microsoft Teams meeting as an AI avatar participant with bidirectional audio.

### Additional prerequisites

- An **Azure Communication Services** resource with a connection string (see `.env` above)
- A Teams meeting link (the meeting must already be created)
- The `static/acs-calling.js` bundle must be patched once — see [ACS SDK patch](#acs-sdk-patch)

### Joining a meeting

1. Expand the **Teams Integration** panel in the sidebar.
2. Confirm the meeting link and display name (default: `AI Avatar`).
3. Choose the outgoing audio source:
   - **Avatar audio** *(default)* — the AI avatar's voice is sent to Teams participants.
   - **Audio file** — a local audio file plays into the Teams call (with optional looping).
4. Click **Join Teams**. If no Voice Live session is active, one starts automatically.
5. When avatar is enabled, the avatar video is sent into the Teams meeting automatically. Uncheck **Send avatar video to Teams** to stop.
6. Click **Leave Teams** to hang up and clean up all resources.

### Audio architecture

The single shared `playbackContext` (24 kHz `AudioContext`) is the foundation of the entire audio graph. It is created at Teams join time (or at Voice Live session start, whichever comes first) and persists for the duration of the call.

#### Teams → Voice Live (remote participants speak to the AI)

```
Teams call (remote participants)
    │
    ▼
call.remoteAudioStreams[0].getMediaStream()   ← requires patched ACS SDK
    │
    ▼
AudioContext (24 kHz)  →  createMediaStreamSource()
    │
    ▼
AudioWorklet: PCM16Processor                 ← Float32 → Int16 PCM in real time
    │
    ▼
WebSocket → Python server → Voice Live API
```

#### Voice Live → Teams (AI speaks to meeting participants)

Two sub-paths depending on whether avatar (WebRTC) is active:

**Non-avatar mode** (PCM audio delivered by server):

```
Voice Live API → Python → WebSocket (audio_data, base64 PCM16)
    │
    ▼
handleAudioDelta() — PCM16 → Float32 → BufferSource (playbackContext)
    │
    ├──→ captureNode → analyserNode → AudioContext.destination   (local speakers)
    └──→ captureNode → _teamsBridgeDest (MediaStreamDestination)
                            │
                            ▼
                    LocalAudioStream → Teams call
```

**Avatar mode** (audio delivered via WebRTC):

```
Azure Voice Live  ──WebRTC──►  browser (ontrack, audio stream)
    │
    ▼
createMediaStreamSource(webrtcStream, playbackContext)
    │
    ├──→ analyserNode → AudioContext.destination   (local speakers)
    └──→ _teamsBridgeDest (MediaStreamDestination, same playbackContext)
                            │
                            ▼
                    LocalAudioStream → Teams call
```

**Key design:** The call is joined with a silent (zero-gain oscillator) feed so the participant appears unmuted from the start. When real audio arrives, the silent feed is disconnected and replaced by the actual source — all within the **same `AudioContext` instance**, avoiding cross-context resampling and relay overhead.

### ACS SDK patch

`static/acs-calling.js` ships with raw media stream access disabled. Run this once from the project root:

```powershell
python rebuild-acs/patch_acs.py
```

To rebuild the bundle from scratch from npm:

```powershell
cd rebuild-acs
python build_acs.py                    # uses @azure/communication-calling@1.42.1
python build_acs.py --version 1.43.0   # pin a different version
```

`build_acs.py` downloads from npm, bundles with esbuild, applies the patch automatically, and writes `static/acs-calling.js`.

---

## Project Structure

```
voice-live-avatar-teams/
├── app.py                    # FastAPI server: WebSocket endpoint, /api/acs-token, static serving
├── voice_handler.py          # Voice Live SDK session lifecycle, event processing, audio forwarding
├── requirements.txt          # Python dependencies
├── Dockerfile                # Container configuration
├── .env.template             # Environment variable template (copy to .env and fill in values)
├── .gitignore                # Git ignore rules
├── README.md                 # This file
├── static/
│   ├── index.html            # Main UI page
│   ├── style.css             # Styles
│   ├── app.js                # Client JS: audio, WebRTC, Teams bridge, UI state
│   └── acs-calling.js        # ACS Calling SDK bundle (patched — see ACS SDK patch section)
└── rebuild-acs/
    ├── build_acs.py          # Download ACS SDK from npm, bundle, patch, write static/acs-calling.js
    └── patch_acs.py          # Apply the allowAccessRawMediaStream patch to an existing bundle
```

> **Not committed:** `.env` (local secrets), `__pycache__/` — both covered by `.gitignore`.

---

## HTTP & WebSocket API

### REST endpoints

| Method | Path | Description |
|---|---|---|
| `GET` | `/health` | Health check — returns `{"status": "healthy"}` |
| `GET` | `/api/config` | Default config values from environment variables |
| `GET` | `/api/acs-token` | Mint a short-lived ACS user identity + `voip`-scoped token |
| `WS` | `/ws/{client_id}` | Main Voice Live session WebSocket |

### WebSocket — Frontend → Backend

| `type` | Description |
|---|---|
| `start_session` | Start a Voice Live session; carries full config object |
| `stop_session` | Stop the active session |
| `audio_chunk` | Microphone audio (base64 PCM16, 24 kHz) |
| `send_text` | Send a text message to the assistant |
| `avatar_sdp_offer` | Forward the browser WebRTC SDP offer for avatar setup |
| `interrupt` | Cancel the current assistant response |
| `update_scene` | Update photo avatar scene parameters (live while connected) |

### WebSocket — Backend → Frontend

| `type` | Description |
|---|---|
| `session_started` | Session ready; carries `sessionId` and echoed `config` |
| `session_error` | Error starting or during session |
| `ice_servers` | ICE server credentials for the avatar WebRTC peer connection |
| `avatar_sdp_answer` | Server SDP answer (base64 JSON) for avatar WebRTC setup |
| `audio_data` | Assistant audio chunk (base64 PCM16, 24 kHz) |
| `video_data` | Avatar video chunk in WebSocket mode (base64 fMP4) |
| `transcript_delta` | Streaming transcript fragment |
| `transcript_done` | Completed transcript with role and item ID |
| `text_delta` | Streaming text response fragment |
| `text_done` | Text response completed |
| `response_created` | New assistant response started |
| `response_done` | Assistant response fully completed |
| `speech_started` | User started speaking (triggers barge-in) |
| `speech_stopped` | User stopped speaking |
| `avatar_connecting` | Avatar WebRTC connection being established |
| `session_closed` | Session ended by the server |

---

## Deployment

The recommended hosting platform is [Azure Container Apps](https://learn.microsoft.com/azure/container-apps/overview).

### Push image to Azure Container Registry

```bash
docker tag voice-live-avatar-python <registry>.azurecr.io/voice-live-avatar-python:latest
docker push <registry>.azurecr.io/voice-live-avatar-python:latest
```

### Deploy to Azure Container Apps

Follow [Deploy from an existing container image](https://learn.microsoft.com/azure/container-apps/quickstart-portal). Set the environment variables from your `.env` file as Container App secrets/environment variables.

> **WebRTC note:** Avatar video uses a direct WebRTC peer connection between the browser and Azure. Ensure the Container App ingress is HTTPS (required by browser WebRTC) and that the app is reachable from the public internet.
