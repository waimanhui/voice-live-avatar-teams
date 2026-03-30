"""
Voice Live Avatar - Python Backend
FastAPI server that bridges browser WebSocket with Azure Voice Live SDK.
Avatar video streams via WebRTC directly to browser; audio and events are relayed through this server.
"""

import asyncio
import json
import logging
import os
from contextlib import asynccontextmanager
from typing import Dict

import uvicorn
from azure.core.credentials import AzureKeyCredential
from dotenv import load_dotenv
from fastapi import FastAPI, WebSocket, WebSocketDisconnect
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
from fastapi.staticfiles import StaticFiles

from voice_handler import VoiceSessionHandler

load_dotenv()

# Logging with color
class ColorFormatter(logging.Formatter):
    """Custom formatter that adds ANSI color codes to log output."""
    COLORS = {
        logging.DEBUG:    "\033[36m",     # Cyan
        logging.INFO:     "\033[32m",     # Green
        logging.WARNING:  "\033[33m",     # Yellow
        logging.ERROR:    "\033[31m",     # Red
        logging.CRITICAL: "\033[1;31m",   # Bold Red
    }
    RESET = "\033[0m"
    BOLD = "\033[1m"
    DIM = "\033[2m"
    WHITE = "\033[97m"

    def format(self, record):
        color = self.COLORS.get(record.levelno, self.RESET)
        # Timestamp in dim, level in color+bold, name in dim, message in white
        timestamp = self.formatTime(record, self.datefmt)
        return (
            f"{self.DIM}{timestamp}{self.RESET} "
            f"{color}{self.BOLD}{record.levelname:<8}{self.RESET} "
            f"{self.DIM}{record.name}{self.RESET} "
            f"{self.WHITE}{record.getMessage()}{self.RESET}"
        )

handler = logging.StreamHandler()
handler.setFormatter(ColorFormatter())
logging.basicConfig(
    level=logging.INFO,
    handlers=[handler],
)
logger = logging.getLogger(__name__)

# WSMsgType.CLOSING (256) is a normal part of graceful WebSocket shutdown.
# The SDK logs it as WARNING because it falls through its recv_bytes() type
# switch.  Demote the entire SDK patch logger to ERROR so this noise is hidden.
logging.getLogger("azure.ai.voicelive.aio._patch").setLevel(logging.ERROR)

# aiohttp emits 'Unclosed client session' / 'Unclosed connector' through the
# asyncio logger when its __del__ finalizer fires after the event loop is
# closed.  With uvicorn --reload this happens on every Ctrl+C because the
# reloader kills the worker process before the lifespan finally block can
# complete.  The process is already exiting so this is purely cosmetic.
# We install a custom filter rather than raising the whole asyncio logger
# level so genuine asyncio errors are still visible.
class _SuppressAiohttpUnclosed(logging.Filter):
    def filter(self, record: logging.LogRecord) -> bool:
        msg = record.getMessage()
        return "Unclosed client session" not in msg and "Unclosed connector" not in msg

logging.getLogger("asyncio").addFilter(_SuppressAiohttpUnclosed())

# Track active sessions per client
active_sessions: Dict[str, VoiceSessionHandler] = {}
active_tasks: Dict[str, asyncio.Task] = {}


@asynccontextmanager
async def lifespan(app: FastAPI):
    logger.info("Voice Live Avatar server starting...")
    try:
        yield
    finally:
        # Cleanly close every active session on shutdown using the same
        # cleanup_client() path as a normal disconnect.  This ensures
        # handler.stop() closes the aiohttp ClientSession/TCPConnector
        # BEFORE the tasks are cancelled, preventing 'Unclosed client session'.
        client_ids = list(active_sessions.keys()) + [
            cid for cid in active_tasks if cid not in active_sessions
        ]
        for client_id in client_ids:
            try:
                await cleanup_client(client_id)
            except Exception:
                pass
        logger.info("Voice Live Avatar server stopped.")


app = FastAPI(
    title="Voice Live Avatar",
    description="Python backend for Azure Voice Live with Avatar support",
    version="1.0.0",
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.middleware("http")
async def no_cache_static(request, call_next):
    """Disable caching for static assets during development."""
    response = await call_next(request)
    # Only add no-cache headers outside of production so browsers can cache
    # JS/CSS in production deployments (improves load performance).
    if os.getenv("ENV", "development") != "production":
        path = request.url.path
        if path.endswith((".js", ".css", ".html")) or path == "/":
            response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
            response.headers["Pragma"] = "no-cache"
            response.headers["Expires"] = "0"
    return response


@app.get("/health")
async def health_check():
    return {"status": "healthy", "service": "voice-live-avatar"}


@app.get("/api/config")
async def get_config():
    """Return default configuration to the frontend."""
    acs_connection_string = os.getenv("AZURE_COMMUNICATION_CONNECTION_STRING", "")
    acs_endpoint = os.getenv("AZURE_COMMUNICATION_ENDPOINT", "")
    acs_configured = bool(
        (acs_connection_string and acs_connection_string.startswith("endpoint="))
        or acs_endpoint
    )
    return {
        "model": os.getenv("VOICELIVE_MODEL", "gpt-4o-realtime"),
        "voice": os.getenv("VOICELIVE_VOICE", "en-US-AvaMultilingualNeural"),
        "endpoint": os.getenv("AZURE_VOICELIVE_ENDPOINT", ""),
        "apiKey": os.getenv("AZURE_VOICELIVE_API_KEY", ""),
        "acsConnectionString": acs_connection_string,
        "teamsMeetingLink": os.getenv("TEAMS_MEETING_LINK", ""),
        "acsConfigured": acs_configured,
    }


@app.get("/api/acs-token")
async def get_acs_token():
    """Generate a short-lived Azure Communication Services user token for Teams calling.

    Credential resolution order (mirrors the Voice Live credential logic):
      1. Connection string (AZURE_COMMUNICATION_CONNECTION_STRING) — access-key path
      2. DefaultAzureCredential + endpoint URL (AZURE_COMMUNICATION_ENDPOINT)
           - Local dev:    az login
           - Azure-hosted: managed identity
           Requires the identity to have the "Contributor" or
           "Communication Services Owner" role on the ACS resource.
    """
    from fastapi import HTTPException
    from azure.communication.identity import CommunicationIdentityClient

    connection_string = os.getenv("AZURE_COMMUNICATION_CONNECTION_STRING", "")
    acs_endpoint = os.getenv("AZURE_COMMUNICATION_ENDPOINT", "")

    if not connection_string and not acs_endpoint:
        raise HTTPException(
            status_code=500,
            detail=(
                "ACS is not configured. Set AZURE_COMMUNICATION_CONNECTION_STRING "
                "or AZURE_COMMUNICATION_ENDPOINT (for az login / managed identity)."
            ),
        )

    try:
        if connection_string:
            client = CommunicationIdentityClient.from_connection_string(connection_string)
            logger.info("ACS: using connection string (access key)")
            user, token_response = client.create_user_and_token(scopes=["voip"])
        else:
            from azure.identity.aio import DefaultAzureCredential as AsyncDefaultAzureCredential
            from azure.communication.identity.aio import CommunicationIdentityClient as AsyncCommunicationIdentityClient
            async with AsyncDefaultAzureCredential() as aio_credential:
                async with AsyncCommunicationIdentityClient(acs_endpoint, aio_credential) as async_client:
                    logger.info("ACS: using DefaultAzureCredential (az login / managed identity)")
                    user, token_response = await async_client.create_user_and_token(scopes=["voip"])

        expires_on = token_response.expires_on
        return {
            "userId": user.properties["id"],
            "token": token_response.token,
            "expiresOn": expires_on.isoformat() if hasattr(expires_on, "isoformat") else str(expires_on),
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to issue ACS token: {e}")


@app.post("/api/acs-delete-user")
async def delete_acs_user(body: dict):
    """Delete an ephemeral ACS identity after a call ends.

    Prevents unused identities from accumulating in the ACS resource.
    Accepts: { "userId": "<acs-user-id>" }
    """
    from fastapi import HTTPException
    from azure.communication.identity import CommunicationIdentityClient, CommunicationUserIdentifier

    user_id = body.get("userId", "").strip()
    if not user_id:
        raise HTTPException(status_code=400, detail="userId is required")

    connection_string = os.getenv("AZURE_COMMUNICATION_CONNECTION_STRING", "")
    acs_endpoint = os.getenv("AZURE_COMMUNICATION_ENDPOINT", "")

    if not connection_string and not acs_endpoint:
        raise HTTPException(
            status_code=500,
            detail="ACS is not configured. Set AZURE_COMMUNICATION_CONNECTION_STRING or AZURE_COMMUNICATION_ENDPOINT.",
        )

    try:
        if connection_string:
            client = CommunicationIdentityClient.from_connection_string(connection_string)
            client.delete_user(CommunicationUserIdentifier(user_id))
        else:
            from azure.identity.aio import DefaultAzureCredential as AsyncDefaultAzureCredential
            from azure.communication.identity.aio import CommunicationIdentityClient as AsyncCommunicationIdentityClient
            async with AsyncDefaultAzureCredential() as aio_credential:
                async with AsyncCommunicationIdentityClient(acs_endpoint, aio_credential) as async_client:
                    await async_client.delete_user(CommunicationUserIdentifier(user_id))
        logger.info(f"ACS: deleted user identity {user_id}")
        return {"deleted": user_id}
    except Exception as e:
        # Non-fatal: log and return a soft error so the client doesn't surface it as a hard failure
        logger.warning(f"ACS: failed to delete user {user_id}: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to delete ACS user: {e}")


@app.websocket("/ws/{client_id}")
async def websocket_endpoint(websocket: WebSocket, client_id: str):
    """Main WebSocket endpoint for voice session communication."""
    await websocket.accept()
    logger.info(f"Client {client_id} connected")

    try:
        while True:
            data = await websocket.receive_text()
            message = json.loads(data)
            await handle_message(client_id, message, websocket)
    except WebSocketDisconnect:
        logger.info(f"Client {client_id} disconnected")
    except Exception as e:
        logger.error(f"WebSocket error for {client_id}: {e}")
    finally:
        await cleanup_client(client_id)


async def handle_message(client_id: str, message: dict, websocket: WebSocket):
    """Route incoming WebSocket messages."""
    msg_type = message.get("type")

    if msg_type == "start_session":
        await start_session(client_id, message.get("config", {}), websocket)

    elif msg_type == "stop_session":
        await stop_session(client_id)

    elif msg_type == "audio_chunk":
        handler = active_sessions.get(client_id)
        if handler:
            await handler.send_audio(message.get("data", ""))

    elif msg_type == "send_text":
        handler = active_sessions.get(client_id)
        if handler:
            await handler.send_text_message(message.get("text", ""))

    elif msg_type == "avatar_sdp_offer":
        handler = active_sessions.get(client_id)
        if handler:
            await handler.send_avatar_sdp_offer(message.get("clientSdp", ""))

    elif msg_type == "interrupt":
        handler = active_sessions.get(client_id)
        if handler:
            await handler.interrupt()

    elif msg_type == "update_scene":
        handler = active_sessions.get(client_id)
        if handler:
            await handler.update_avatar_scene(message.get("avatar", {}))

    else:
        logger.warning(f"Unknown message type: {msg_type}")


async def start_session(client_id: str, config: dict, websocket: WebSocket):
    """Start a new Voice Live session for a client."""
    # Clean up any existing session
    await cleanup_client(client_id)

    # Prefer credentials from frontend config, fall back to env vars
    endpoint = config.get("endpoint", "").strip() or os.getenv("AZURE_VOICELIVE_ENDPOINT", "")
    api_key = config.get("apiKey", "").strip() or os.getenv("AZURE_VOICELIVE_API_KEY", "")
    entra_token = config.get("entraToken", "").strip()

    if not endpoint:
        await send_ws_message(websocket, {
            "type": "session_error",
            "error": "Azure AI Services Endpoint is required. Provide it in the UI or set AZURE_VOICELIVE_ENDPOINT.",
        })
        return

    # Credential resolution order:
    #   1. Entra token — passed explicitly from the browser (agent / agent-v2 modes)
    #   2. API key   — AZURE_VOICELIVE_API_KEY env var or entered in the UI
    #   3. DefaultAzureCredential — no key configured; covers:
    #        - Local dev:  az login (Azure CLI credential)
    #        - Azure-hosted: managed identity automatically
    if entra_token:
        from azure.core.credentials import AccessToken
        import time

        class _StaticTokenCredential:
            """Wraps a raw token string as an async TokenCredential."""
            def __init__(self, token: str):
                self._token = token
            async def get_token(self, *scopes, **kwargs):
                return AccessToken(self._token, int(time.time()) + 3600)
            async def close(self): pass
            async def __aenter__(self): return self
            async def __aexit__(self, *args): pass

        credential = _StaticTokenCredential(entra_token)
        logger.info("Using static Entra token for Voice Live connection")
    elif api_key:
        credential = AzureKeyCredential(api_key)
        logger.info("Using API key for Voice Live connection")
    else:
        # No API key — use DefaultAzureCredential:
        #   locally this picks up `az login`; in Azure it picks up managed identity.
        try:
            from azure.identity.aio import DefaultAzureCredential
            credential = DefaultAzureCredential()
            logger.info("No API key configured — using DefaultAzureCredential (az login / managed identity)")
        except ImportError:
            await send_ws_message(websocket, {
                "type": "session_error",
                "error": "No credentials provided. Set AZURE_VOICELIVE_API_KEY or run `az login` for local dev.",
            })
            return

    async def send_message(msg: dict):
        try:
            await websocket.send_text(json.dumps(msg))
        except Exception as e:
            logger.error(f"Error sending to {client_id}: {e}")

    handler = VoiceSessionHandler(
        client_id=client_id,
        endpoint=endpoint,
        credential=credential,
        send_message=send_message,
        config=config,
    )
    active_sessions[client_id] = handler

    # Run session in background task
    task = asyncio.create_task(handler.start())
    active_tasks[client_id] = task
    logger.info(f"Session started for {client_id}")


async def stop_session(client_id: str):
    """Stop an active session."""
    await cleanup_client(client_id)


async def cleanup_client(client_id: str):
    """Clean up session and task for a client.

    Order matters:
    1. handler.stop() closes the SDK connection (WebSocket + aiohttp session)
       eagerly, while the event loop is still healthy.  This is the step that
       prevents 'Unclosed client session' / 'Unclosed connector' warnings.
    2. task.cancel() is then safe to call — there is nothing left to close, so
       the CancelledError just unblocks whatever await the task is stuck on and
       lets it exit cleanly.
    """
    handler = active_sessions.pop(client_id, None)
    if handler:
        await handler.stop()  # closes aiohttp session BEFORE task.cancel()

    task = active_tasks.pop(client_id, None)
    if task and not task.done():
        task.cancel()
        try:
            await task
        except (asyncio.CancelledError, Exception):
            pass


async def send_ws_message(websocket: WebSocket, message: dict):
    """Send a JSON message via WebSocket."""
    try:
        await websocket.send_text(json.dumps(message))
    except Exception as e:
        logger.error(f"Error sending WebSocket message: {e}")


# Mount static files (frontend)
static_path = os.path.join(os.path.dirname(__file__), "static")
if os.path.exists(static_path):
    app.mount("/", StaticFiles(directory=static_path, html=True), name="static")
else:
    @app.get("/")
    async def root():
        return {"message": "Voice Live Avatar - static files not found. Place frontend in ./static/"}


if __name__ == "__main__":
    uvicorn.run("app:app", host="0.0.0.0", port=3000, reload=True, log_level="info")
