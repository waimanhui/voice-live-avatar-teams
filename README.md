# Voice Live Avatar with Teams Integration

> Based on: https://github.com/microsoft-foundry/voicelive-samples/tree/main/python/voice-live-avatar

A Python (FastAPI) + browser-based sample that demonstrates the **Azure AI Voice Live API** with a real-time animated avatar and optional Microsoft Teams meeting integration.

---

## Table of Contents

1. [Architecture Overview](#architecture-overview)
2. [Prerequisites](#prerequisites)
3. [Setup](#setup)
4. [Docker](#docker)
5. [How It Works](#how-it-works)
   - [Operation Modes](#operation-modes)
   - [Voice Types](#voice-types)
   - [Avatar Types](#avatar-types)
6. [Using the Sample](#using-the-sample)
   - [Quick Start](#quick-start)
   - [Settings Reference](#settings-reference)
7. [Teams Integration](#teams-integration)
8. [Project Structure](#project-structure)
9. [HTTP & WebSocket API](#http--websocket-api)
10. [Deployment](#deployment)
11. [Uncovered Features](#uncovered-features)

---

## Architecture Overview

```
┌─────────────────────────────────────┐        ┌──────────────────────────┐        ┌──────────────────────┐
│         Browser (Frontend)          │◄──WS──►│  Python Server (FastAPI) │◄──SDK─►│  Azure Voice Live    │
│                                     │        │                          │        │  Service             │
│  • Mic capture (PCM16, 24 kHz)      │        │  • Session lifecycle      │        └──────────────────────┘
│  • Audio playback (PCM16, 24 kHz)   │        │  • Audio forwarding       │                 │
│  • Avatar video (WebRTC)       ◄────────────────────────────── WebRTC (peer-to-peer) ───────┘
│  • Settings / chat UI               │        │  • Event relay            │
│  • Teams ACS calling                │        │  • SDP relay              │
└──────────────────┬──────────────────┘        │  • ACS token minting      │
                   │ ACS Calling SDK            └──────────────────────────┘
                   ▼
         ┌──────────────────┐
         │  Teams Meeting   │  ← AI avatar joins as a named participant
         │  (via ACS SDK)   │    bidirectional audio via Web Audio API
         └──────────────────┘
```


https://github.com/user-attachments/assets/e60f0881-036b-4f8a-b00c-b45d07b1b4f4


### Key design points

| Component | Responsibility |
|---|---|
| **Python server** (`app.py`) | WebSocket bridge; routes audio, events, and SDP between browser and Azure SDK; mints ACS tokens |
| **`VoiceSessionHandler`** (`voice_handler.py`) | Owns the Azure Voice Live SDK connection; builds session config; processes all server events |
| **Browser** (`app.js`) | Audio capture/playback pipeline (AudioWorklet, 24 kHz PCM16); WebRTC peer connection for avatar video; Teams meeting join via ACS Calling SDK |
| **Avatar video** | Streams **peer-to-peer** from Azure directly to the browser via WebRTC — the Python server only relays the SDP offer/answer, never touches video data |
| **Teams audio** | Runs **entirely in the browser** using the ACS Calling SDK and Web Audio API; Python server only mints a short-lived ACS token |

---

## Prerequisites

- **Python 3.10+**
- An **Azure AI Services** resource in a [supported region](#avatar-supported-regions)
- An **Azure Communication Services** resource *(required only for Teams integration)*

### Avatar supported regions

Avatar streaming is available in: Southeast Asia, North Europe, West Europe, Sweden Central, South Central US, East US 2, West US 2.

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

# Authentication — choose one:
#   Option A: API key
AZURE_VOICELIVE_API_KEY=<your-api-key>
#   Option B: Leave blank — DefaultAzureCredential is used automatically:
#     - Local dev: run `az login` once; no key needed.
#     - Azure-hosted: assign "Cognitive Services User" role to the managed identity.

# Optional — pre-fills the UI on page load
VOICELIVE_MODEL=gpt-4o-realtime
VOICELIVE_VOICE=en-US-AvaMultilingualNeural

# Required only for Teams integration — choose one:
#   Option A: Connection string (access key)
AZURE_COMMUNICATION_CONNECTION_STRING=endpoint=https://<acs-resource>.communication.azure.com/;accesskey=<key>
#   Option B: Endpoint URL — DefaultAzureCredential (az login / managed identity)
AZURE_COMMUNICATION_ENDPOINT=https://<acs-resource>.communication.azure.com/
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

# Run (pass environment variables)
docker run --rm -p 3000:3000 --env-file .env voice-live-avatar-python
```

Then open [http://localhost:3000](http://localhost:3000).

---

## How It Works

### Operation Modes

Select the mode under **Connection Settings** before connecting.

| Mode | Description | Required fields |
|---|---|---|
| `model` | Connects directly to an Azure AI model (GPT Realtime or a cascaded model) | Endpoint, Subscription Key (or `az login` / managed identity), Model |
| `agent` | Uses a pre-deployed agent referenced by its ID | Endpoint, Entra Token, Agent Project Endpoint, Agent ID |
| `agent-v2` | Uses a pre-deployed agent referenced by its name | Endpoint, Entra Token, Agent Project Endpoint, Agent Name |

**Available models** (for `model` mode):

| Category | Models |
|---|---|
| GPT Realtime (native audio) | `gpt-4o-realtime`, `gpt-4o-mini-realtime` |
| Cascaded (text LLM + TTS) | `gpt-4.1`, `gpt-4.1-mini`, `gpt-4.1-nano`, `gpt-4o`, `gpt-4o-mini`, `phi4-mm`, `phi4-mini` |

> **GPT Realtime** models process audio natively (lower latency, more natural). **Cascaded** models convert speech-to-text first, then generate a text response, then synthesise speech.

### Voice Types

| Type | Description |
|---|---|
| **Standard** | Azure Neural TTS voices, including multilingual (e.g. `en-US-AvaMultilingualNeural`) and DragonHD high-fidelity voices |
| **OpenAI** | OpenAI voices: alloy, ash, ballad, coral, echo, sage, shimmer, verse |
| **Custom** | A voice trained in Azure Custom Voice |
| **Personal** | A cloned voice using Azure Personal Voice (Dragon HD pipeline) |

### Avatar Types

| Type | Description | Output modes |
|---|---|---|
| **Prebuilt** | One of 15 pre-built characters × styles (e.g. `Lisa-casual-sitting`) | WebRTC, WebSocket |
| **Photo** | Animate a still photo using the VASA-1 model | WebRTC only |
| **Custom** | A custom-trained avatar | WebRTC, WebSocket |

**Output modes:**

- **WebRTC** — Avatar video streams peer-to-peer from Azure to the browser. Lower latency; recommended for Teams integration (video bridging).
- **WebSocket** — Avatar video is delivered as fMP4 chunks through the Python server WebSocket. Works in environments where WebRTC is blocked.

---

## Using the Sample

### Quick Start

1. Open [http://localhost:3000](http://localhost:3000).
2. Under **Connection Settings**, enter your **Azure AI Services Endpoint** and **Subscription Key** (or leave the key blank if you have run `az login`).
3. *(Optional)* Enable avatar, adjust voice settings, etc.
4. Click **Connect**.
5. Once connected, click **Turn on microphone** and start speaking.

Toggle **Developer Mode** (top-right) to reveal the chat transcript, system log messages, and a text input for sending messages without the microphone.

### Settings Reference

#### Connection Settings

| Field | Description |
|---|---|
| Mode | `model` / `agent` / `agent-v2` — see [Operation Modes](#operation-modes) |
| Endpoint | `https://<resource>.cognitiveservices.azure.com/` |
| Subscription Key | Azure AI Services API key. Optional — if blank, the server uses `DefaultAzureCredential` (`az login` locally, managed identity in Azure). |
| Entra Token | Microsoft Entra access token (required for `agent` / `agent-v2` modes) |
| Agent Project Endpoint | Azure AI Foundry project endpoint (agent modes only) |
| Agent ID / Name | Identifier for the deployed agent (agent modes only) |
| Model | LLM used for the session (see [Available Models](#operation-modes)) |

#### Conversation Settings

| Field | Description |
|---|---|
| Speech Recognition Model | `azure-speech` (standard) or `mai-ears-1` (enhanced accuracy) |
| Recognition Language | Primary language for speech recognition (13 options) |
| Noise Suppression | Remove background noise from microphone input |
| Echo Cancellation | Suppress acoustic echo from microphone input |
| Turn Detection | `Server VAD` — silence-based end-of-turn; `Azure Semantic VAD` — meaning-based end-of-turn |
| Filler Words | Remove filler words (um, uh, etc.) from user input before LLM processing |
| EOU Detection | End-of-utterance detection model for Semantic VAD (`none` / `semantic_detection_v1`) |
| Instructions | System prompt / persona instructions for the model |
| Proactive Response | Model greets the user automatically on session start |
| Temperature | Response randomness (0 = deterministic, 2 = very creative) |

#### Voice Configuration

| Field | Description |
|---|---|
| Voice Type | `standard` / `custom` / `personal` / OpenAI voices |
| Voice | Specific voice name (e.g. `en-US-AvaMultilingualNeural`) |
| Voice Temperature | Expressiveness of the voice (standard/personal voices only) |
| Voice Speed | Speaking rate multiplier |

#### Avatar Configuration

| Field | Description |
|---|---|
| Enable Avatar | Toggle avatar on/off |
| Avatar Type | `Prebuilt` / `Photo` / `Custom` |
| Output Mode | `WebRTC` (recommended) or `WebSocket` (fMP4 stream) |
| Avatar Character | Character × style for prebuilt avatars |
| Photo Avatar Character | Character for photo avatars |
| Custom Avatar Name | Name of a custom-trained avatar |
| Background Image URL | Optional background image for the avatar |

#### Scene Settings *(Photo avatar only)*

These sliders are available only when a photo avatar is enabled. Changes apply **live** while connected — no reconnect needed.

| Slider | Range | Description |
|---|---|---|
| Zoom | 70%–100% | Camera zoom level |
| Position X / Y | −50%–+50% | Horizontal / vertical offset |
| Rotation X / Y / Z | −30°–+30° | 3D head rotation |
| Amplitude | 10%–100% | Head motion intensity |

#### Teams Integration

| Field | Description |
|---|---|
| Teams Meeting Link | `https://teams.microsoft.com/l/meetup-join/...` |
| Display Name | Name shown in the Teams participant list (default: `AI Avatar`) |
| Audio Source | `Avatar audio` — AI voice into Teams; `Audio file` — play a local audio file into Teams |
| Send avatar video | Send the avatar WebRTC video stream into the Teams meeting (WebRTC output mode only) |
| W / H / FPS | Canvas resolution and frame rate for the video stream sent to Teams |
| Loop | Loop the audio file (Audio file mode only) |

---

## Teams Integration

The sample can join a Microsoft Teams meeting as an AI avatar participant with **bidirectional audio**.

### Additional prerequisites

- An **Azure Communication Services** resource with a connection string (see `.env` above)
- A Teams meeting link (the meeting must already be created)
- The `static/acs-calling.js` bundle must be patched once — see [ACS SDK patch](#acs-sdk-patch)

### Joining a meeting

1. Expand the **Teams Integration** panel in the sidebar.
2. Enter the meeting link and a display name.
3. Choose the outgoing audio source:
   - **Avatar audio** *(default)* — the AI avatar's voice is sent to Teams participants.
   - **Audio file** — a local audio file plays into the Teams call (with optional looping).
4. Click **Join Teams**. If no Voice Live session is active, one starts automatically.
5. *(Optional)* Check **Send avatar video to Teams** to have the avatar's face appear as the participant's video feed (requires WebRTC output mode).
6. Click **Leave Teams** to hang up and release all resources.

### Audio architecture

The single shared `playbackContext` (24 kHz `AudioContext`) is the foundation of the entire audio graph. It is created at Teams join time (or at Voice Live session start, whichever comes first) and persists for the duration of the call.

#### Teams → Voice Live (remote participants speak to the AI)

```
Teams call (remote participants)
    │
    ▼
remoteAudioStream.getMediaStream()     ← requires patched ACS SDK
    │
    ▼
playbackContext (24 kHz AudioContext)
    → createMediaStreamSource()
    │
    ▼
AudioWorklet: PCM16Processor           ← Float32 → Int16 in real time
    │
    ▼
WebSocket → Python server → Voice Live API
```

#### Voice Live → Teams (AI speaks to meeting participants)

Two paths depending on whether avatar WebRTC is active:

**Non-avatar (PCM from server):**

```
Voice Live → Python → WebSocket (audio_data, base64 PCM16)
    │
    ▼
handleAudioDelta()  →  BufferSource (playbackContext)
    │
    ├──→ analyserNode → speakers (local playback)
    └──→ _teamsBridgeDest (MediaStreamDestination)
              │
              ▼
        LocalAudioStream → Teams call
```

**Avatar / WebRTC mode:**

```
Azure Voice Live ──WebRTC──► browser (ontrack, audio stream)
    │
    ▼
createMediaStreamSource(webrtcStream, playbackContext)
    │
    ├──→ analyserNode → speakers (local playback)
    └──→ _teamsBridgeDest (MediaStreamDestination, same playbackContext)
              │
              ▼
        LocalAudioStream → Teams call
```

> **Key design:** The call is joined with a silent oscillator feed so the AI participant appears unmuted from the very start. When real audio arrives, the silent feed is disconnected and replaced by the actual source — all within the **same `AudioContext` instance**, avoiding cross-context resampling.

### ACS SDK patch

`static/acs-calling.js` ships with raw media stream access disabled. Run this once from the project root to enable `remoteAudioStream.getMediaStream()`:

```powershell
python rebuild-acs/patch_acs.py
```

To rebuild the bundle from scratch from npm:

```powershell
cd rebuild-acs
python build_acs.py                    # uses @azure/communication-calling@1.42.1
python build_acs.py --version 1.43.0   # pin a different version
```

`build_acs.py` downloads the ACS SDK from npm, bundles it with esbuild, applies the patch automatically, and writes `static/acs-calling.js`.

> **Note:** The patch must be re-applied after each ACS SDK upgrade.

---

## Project Structure

```
voice-live-avatar-teams/
├── app.py                    # FastAPI server: WebSocket, REST endpoints, static file serving
├── voice_handler.py          # Azure Voice Live SDK session lifecycle and event processing
├── requirements.txt          # Python dependencies
├── Dockerfile                # Container image definition
├── .env.template             # Environment variable template — copy to .env and fill in values
├── .gitignore
├── README.md
├── static/
│   ├── index.html            # Single-page UI (sidebar panels + main content area)
│   ├── style.css             # All UI styles
│   ├── app.js                # Client-side JS: audio pipeline, WebRTC, Teams bridge, UI logic
│   └── acs-calling.js        # Patched ACS Calling SDK bundle (see ACS SDK patch)
└── rebuild-acs/
    ├── build_acs.py          # Download, bundle, and patch ACS SDK from npm
    └── patch_acs.py          # Re-apply the allowAccessRawMediaStream patch to an existing bundle
```

### `app.py` — FastAPI server

- Serves the static frontend at `/`
- Accepts one WebSocket connection per browser client at `/ws/{client_id}`
- Creates a `VoiceSessionHandler` per client and runs it as an `asyncio` background task
- Mints short-lived ACS `voip` tokens at `/api/acs-token`
- Returns environment variable defaults at `/api/config` (pre-fills the UI on load)
- Supports multiple simultaneous clients — each has an independent session, audio stream, and task

### `voice_handler.py` — Voice Live session handler

- `VoiceSessionHandler` manages the full lifecycle of one Azure Voice Live SDK session
- Builds a `RequestSession` object from the client config: voice, avatar, turn detection, SR options, noise/echo settings, proactive greeting
- Connects via `azure.ai.voicelive.aio.connect()` and processes all `ServerEventType` events in a `while` loop (manual `recv()` — individual event errors don't kill the session)
- Relays audio deltas, transcripts, text responses, WebRTC SDP, and video chunks to the browser over WebSocket
- Handles function calls end-to-end: receives `CONVERSATION_ITEM_CREATED`, waits for arguments, executes built-in tools (`get_time`, `get_weather`, `calculate`), returns output, triggers `response.create()`
- Proactive greeting: for WebRTC avatar sessions the greeting is deferred until `SESSION_AVATAR_CONNECTING` so the avatar is visible before the AI speaks

### `static/app.js` — Client-side logic

- **Audio capture:** `AudioWorklet` (`PCM16Processor`) samples the microphone at 24 kHz, encodes to PCM16, and sends base64 frames over the WebSocket
- **Audio playback:** Received `audio_data` frames are decoded PCM16 → Float32 and scheduled on a shared `playbackContext` (24 kHz `AudioContext`)
- **WebRTC:** Creates a `RTCPeerConnection`, sends an SDP offer via WebSocket, receives the answer, and renders the avatar `<video>` element when the track arrives
- **WebSocket mode avatar:** Decodes incoming base64 fMP4 chunks into a `<video>` via `MediaSource`
- **Teams bridge:** Joins a Teams meeting via the ACS Calling SDK, routes AI audio (PCM or WebRTC) to a `MediaStreamDestination`, and captures remote audio through an `AudioWorklet` for forwarding to Voice Live
- **Canvas video:** Draws the avatar `<video>` onto a `<canvas>` element and sends it into Teams as a `LocalVideoStream` (uses `setInterval` to remain active when the browser tab is in the background)

---

## HTTP & WebSocket API

### REST endpoints

| Method | Path | Description |
|---|---|---|
| `GET` | `/health` | Health check — returns `{"status": "healthy"}` |
| `GET` | `/api/config` | Default config from environment variables; also returns `acsConfigured: bool` |
| `GET` | `/api/acs-token` | Mint an ACS user identity + `voip`-scoped token |
| `WS` | `/ws/{client_id}` | Main Voice Live session WebSocket |

### WebSocket — Frontend → Backend

| `type` | Payload | Description |
|---|---|---|
| `start_session` | `config: {...}` | Start a Voice Live session with the given configuration |
| `stop_session` | — | Stop the active session |
| `audio_chunk` | `data: <base64 PCM16>` | Microphone audio frame (24 kHz, mono) |
| `send_text` | `text: <string>` | Send a text message to the assistant |
| `avatar_sdp_offer` | `clientSdp: <string>` | Forward the browser WebRTC SDP offer for avatar setup |
| `interrupt` | — | Cancel the current assistant response |
| `update_scene` | `avatar: {...}` | Update photo avatar scene parameters live |

### WebSocket — Backend → Frontend

| `type` | Payload | Description |
|---|---|---|
| `session_started` | `sessionId`, `config` | Session ready |
| `session_error` | `error` | Error starting or during the session |
| `session_closed` | — | Session ended by the server |
| `ice_servers` | `iceServers: [...]` | ICE credentials for WebRTC peer connection setup |
| `avatar_sdp_answer` | `serverSdp` | Server SDP answer for WebRTC avatar setup |
| `audio_data` | `data` (base64 PCM16), `sampleRate` | Assistant audio chunk |
| `audio_done` | — | End of audio for the current response turn |
| `video_data` | `delta` (base64 fMP4) | Avatar video chunk (WebSocket output mode only) |
| `transcript_delta` | `role`, `delta` | Streaming transcript fragment |
| `transcript_done` | `role`, `transcript`, `itemId` | Full transcript for a turn |
| `text_delta` | `delta` | Streaming text response fragment |
| `text_done` | `text` | Full text response |
| `response_created` | `responseId` | New assistant response started |
| `response_done` | — | Assistant response fully completed |
| `speech_started` | `itemId` | User started speaking (triggers barge-in) |
| `speech_stopped` | — | User stopped speaking |
| `function_call_started` | `functionName`, `callId` | Function call being executed |
| `function_call_result` | `functionName`, `callId`, `result` | Function call completed |
| `error` | `error` | Error event from Voice Live API |

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

---

## Uncovered Features

This sample is a working prototype, not a production-ready application. The tables below describe capabilities that are not yet implemented, along with the recommended path to add them. Entries marked **🔴 Hard limit** require a fundamentally different architecture (server-side bot) and cannot be unlocked within the current browser-side ACS approach.

### Teams — Identity

| Feature | How to add it |
|---|---|
| Join with a real Teams identity (avatar photo, display name, roster presence) | Register an [Azure Bot](https://learn.microsoft.com/azure/bot-service/bot-service-overview) via the Bot Framework / Teams Bot SDK instead of an anonymous ACS guest |
| Access meeting chat, participant roster, and live transcription | **🔴 Hard limit** for raw ACS calling. Requires [Microsoft Graph API](https://learn.microsoft.com/graph/api/resources/teams-api-overview) or the Bot Framework |

### Teams — Meeting Management

| Feature | How to add it |
|---|---|
| Create or schedule meetings programmatically | Use [Graph API `POST /me/onlineMeetings`](https://learn.microsoft.com/graph/api/application-post-onlinemeetings) from the server |
| Lobby handling — detect and surface waiting state | Poll `teamsCall.state` for `"InLobby"`, show UI feedback, and use Graph API to admit the guest |
| Multiple simultaneous Teams meetings per browser tab | Refactor all Teams singleton state in `app.js` into a `Map` keyed by session ID |

### Teams — Audio

| Feature | How to add it |
|---|---|
| Raw media stream access without SDK patching | **🔴 Hard limit** until Microsoft exposes `remoteAudioStream.getMediaStream()` in the public ACS Calling SDK. Re-run `patch_acs.py` after each SDK upgrade in the meantime |
| Per-participant audio separation | **🔴 Hard limit** for browser ACS. Requires the server-side [Teams Bot Media SDK](https://learn.microsoft.com/microsoftteams/platform/bots/calls-and-meetings/real-time-media-concepts) |
| Reliable echo cancellation across all browsers and OS | No complete solution in the browser. Server-side echo cancellation (`useEC` toggle) is a more reliable fallback |

### Teams — Video

| Feature | How to add it |
|---|---|
| Send avatar video to Teams in WebSocket output mode | Not possible in the current architecture; only the WebRTC output mode exposes a capturable video track |
| AI receives and renders incoming video from Teams participants | **🔴 Hard limit** for browser ACS. Requires the server-side [Teams Bot Media SDK](https://learn.microsoft.com/microsoftteams/platform/bots/calls-and-meetings/real-time-media-concepts) |

### Teams — Reliability

| Feature | How to add it |
|---|---|
| Automatic reconnect after unexpected disconnect | Implement a `stateChanged` → `"Disconnected"` handler with exponential back-off retry calling `joinTeamsMeeting()` |
| Session survives browser tab closure | Move Teams logic to a server-side bot (Bot Framework + Teams Bot Media SDK); the bot process is independent of the browser tab |

### Audio

| Feature | How to add it |
|---|---|
| Guaranteed 24 kHz microphone capture | After creating `AudioContext`, verify `audioContext.sampleRate === 24000` and warn the user; add a `ScriptProcessorNode` resampler as fallback for browsers that ignore the hint |

### Server — Security & Reliability

| Feature | How to add it |
|---|---|
| WebSocket authentication and session ownership | Issue a signed JWT from `/api/session-token` and verify it on the first WebSocket message; reject mismatched `client_id` values |
| Idle session timeout | Start an `asyncio` timer on connect; reset it on each `audio_chunk`; call `cleanup_client()` on expiry |
| WebSocket message size and rate limiting | Enforce a max frame size in `websocket_endpoint` and drop clients that exceed a per-second audio chunk rate |
| Automatic client-side WebSocket reconnect | Implement exponential back-off reconnect in `ws.onclose` with a maximum retry count |
| Restrict CORS to known origins | Set `allow_origins` to the specific deployed origin(s) instead of `"*"` in production |
| Future-proof avatar scene updates | Track the upstream SDK for a public `session.update` API; until then, the `update_avatar_scene` call uses a private SDK internal and may break silently on SDK upgrades |
| Configurable session setup timeout | Expose `_wait_for_event` timeout via an env var (e.g. `VOICELIVE_SESSION_TIMEOUT_S`); consider a higher default for avatar sessions |

### Conversation & Platform

| Feature | How to add it |
|---|---|
| Persistent conversation history across sessions | Store transcripts in Azure Cosmos DB or Table Storage; reload the last N turns on session start via `conversation.item.create` |
| Server-side bot capabilities (roster control, lobby admit, meeting chat, per-speaker media) | Implement a Bot Framework bot registered via Azure Bot Service; migrate Teams logic off the browser entirely |
