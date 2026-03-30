/**
 * Voice Live Avatar - Client-side JavaScript
 * Handles audio capture (AudioWorklet 24kHz PCM16), WebSocket communication,
 * WebRTC avatar video, and UI state management.
 */

// ===== State =====
let ws = null;
let audioContext = null;
let workletNode = null;
let mediaStream = null;
let playbackContext = null;
let playbackBufferQueue = [];
let nextPlaybackTime = 0;
let isConnected = false;
let isConnecting = false;
let isRecording = false;
let audioChunksSent = 0;
let isDeveloperMode = false;
let avatarEnabled = false;
let peerConnection = null;
let avatarVideoElement = null;
let isSpeaking = false;
let avatarOutputMode = 'webrtc';
let cachedIceServers = null;
let peerConnectionQueue = [];

// Volume animation state
let analyserNode = null;
let analyserDataArray = null;
let micAnalyserNode = null;
let micAnalyserDataArray = null;
let recordAnimationFrameId = null;
let playChunkAnimationFrameId = null;

// WebSocket video playback (MediaSource Extensions)
let mediaSource = null;
let sourceBuffer = null;
let videoChunksQueue = [];
let videoChunkCount = 0;
let pendingWsVideoElement = null;
let mediaSourceObjectURL = null;  // saved so cleanupWebSocketVideo can revoke it

// ICE timeout IDs — stored so cleanupWebRTC can cancel them and avoid leaks
let _icePrepareTimeoutId = null;
let _iceSetupTimeoutId = null;

// Accumulated assistant response text for streaming display
let pendingAssistantText = '';

// Stable-per-tab client ID (survives hot reload, cleared on tab close)
const clientId = (() => {
    const key = 'voicelive_client_id';
    let id = sessionStorage.getItem(key);
    if (!id) {
        id = crypto.randomUUID();
        sessionStorage.setItem(key, id);
    }
    return id;
})();

// ===== Settings Persistence (localStorage) =====
const SETTINGS_KEY = 'voicelive_settings';
// IDs of settings fields to persist (excludes credentials and server-supplied values)
const PERSISTED_SETTINGS = [
    'mode', 'model', 'voiceType', 'voiceName', 'voiceSpeed', 'voiceTemperature',
    'voiceDeploymentId', 'customVoiceName', 'personalVoiceName', 'personalVoiceModel',
    'avatarEnabled', 'isPhotoAvatar', 'isCustomAvatar', 'avatarName', 'photoAvatarName',
    'customAvatarName', 'avatarOutputMode', 'avatarBackgroundImageUrl',
    'useNS', 'useEC', 'turnDetectionType', 'removeFillerWords', 'srModel',
    'recognitionLanguage', 'eouDetectionType', 'instructions', 'temperature',
    'enableProactive', 'toolGetTime', 'toolGetWeather', 'toolCalculate',
    'teamsDisplayName', 'sceneZoom', 'scenePositionX', 'scenePositionY',
    'sceneRotationX', 'sceneRotationY', 'sceneRotationZ', 'sceneAmplitude',
];

function saveSettings() {
    const saved = {};
    for (const id of PERSISTED_SETTINGS) {
        const el = document.getElementById(id);
        if (!el) continue;
        if (el.type === 'checkbox') {
            saved[id] = el.checked;
        } else if (el.tagName === 'TEXTAREA') {
            saved[id] = el.value;
        } else {
            saved[id] = el.value;
        }
    }
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(saved));
}

function restoreSettings() {
    let saved;
    try { saved = JSON.parse(localStorage.getItem(SETTINGS_KEY)); } catch (_) {}
    if (!saved) return;
    for (const [id, val] of Object.entries(saved)) {
        const el = document.getElementById(id);
        if (!el) continue;
        if (el.type === 'checkbox') {
            el.checked = Boolean(val);
        } else {
            el.value = val;
        }
    }
    // Re-sync range display labels after restoring slider values
    ['temperature', 'voiceTemperature', 'voiceSpeed',
     'sceneZoom', 'scenePositionX', 'scenePositionY',
     'sceneRotationX', 'sceneRotationY', 'sceneRotationZ', 'sceneAmplitude'
    ].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.dispatchEvent(new Event('input', { bubbles: true }));
    });
}
document.addEventListener('DOMContentLoaded', () => {
    setupUIBindings();
    updateConditionalFields();
    updateControlStates();
    fetchServerConfig().then(() => {
        restoreSettings();
        updateConditionalFields();
        updateControlStates();
    });
    // Persist settings on any input/change event within the sidebar
    const sidebar = document.getElementById('sidebar');
    if (sidebar) {
        sidebar.addEventListener('input', saveSettings);
        sidebar.addEventListener('change', saveSettings);
    }
});

// ===== Server Config =====
async function fetchServerConfig() {
    try {
        const resp = await fetch('/api/config');
        const config = await resp.json();
        if (config.endpoint) document.getElementById('endpoint').value = config.endpoint;
        if (config.apiKey) document.getElementById('apiKey').value = config.apiKey;
        if (config.model) document.getElementById('model').value = config.model;
        if (config.voice) document.getElementById('voiceName').value = config.voice;
        if (config.teamsMeetingLink) document.getElementById('teamsMeetingLink').value = config.teamsMeetingLink;
        // Warn if ACS is not configured so Teams joining will fail at token fetch time
        if (config.acsConfigured === false) {
            const teamsStatus = document.getElementById('teamsStatus');
            if (teamsStatus) {
                teamsStatus.textContent = 'ACS not configured on server — Teams calling unavailable. Set AZURE_COMMUNICATION_CONNECTION_STRING in .env.';
                teamsStatus.style.display = 'block';
                teamsStatus.style.color = '#c00';
            }
        }
    } catch (e) {
        console.log('No server config available, using defaults');
    }
}

// ===== UI Bindings =====
function setupUIBindings() {
    // Mode change
    document.getElementById('mode').addEventListener('change', updateConditionalFields);
    // Model change
    document.getElementById('model').addEventListener('change', updateConditionalFields);
    // Voice type change
    document.getElementById('voiceType').addEventListener('change', updateConditionalFields);
    // Voice name change
    document.getElementById('voiceName').addEventListener('change', updateConditionalFields);
    // Avatar enabled
    document.getElementById('avatarEnabled').addEventListener('change', updateConditionalFields);
    // Photo avatar
    document.getElementById('isPhotoAvatar').addEventListener('change', updateConditionalFields);
    // Custom avatar
    document.getElementById('isCustomAvatar').addEventListener('change', updateConditionalFields);
    // Developer mode
    document.getElementById('developerMode').addEventListener('change', (e) => {
        isDeveloperMode = e.target.checked;
        updateDeveloperModeLayout();
    });
    // Turn detection type
    document.getElementById('turnDetectionType').addEventListener('change', updateConditionalFields);
    // SR Model
    document.getElementById('srModel').addEventListener('change', updateConditionalFields);

    // Range sliders - display values
    setupRangeDisplay('temperature', 'tempValue', v => v);
    setupRangeDisplay('voiceTemperature', 'voiceTempValue', v => v);
    setupRangeDisplay('voiceSpeed', 'voiceSpeedValue', v => v + '%');
    setupRangeDisplay('sceneZoom', 'sceneZoomLabel', v => 'Zoom: ' + v + '%');
    setupRangeDisplay('scenePositionX', 'scenePositionXLabel', v => 'Position X: ' + v + '%');
    setupRangeDisplay('scenePositionY', 'scenePositionYLabel', v => 'Position Y: ' + v + '%');
    setupRangeDisplay('sceneRotationX', 'sceneRotationXLabel', v => 'Rotation X: ' + v + ' deg');
    setupRangeDisplay('sceneRotationY', 'sceneRotationYLabel', v => 'Rotation Y: ' + v + ' deg');
    setupRangeDisplay('sceneRotationZ', 'sceneRotationZLabel', v => 'Rotation Z: ' + v + ' deg');
    setupRangeDisplay('sceneAmplitude', 'sceneAmplitudeLabel', v => 'Amplitude: ' + v + '%');

    // Scene sliders: send real-time updates when connected
    const sceneSliders = ['sceneZoom', 'scenePositionX', 'scenePositionY',
        'sceneRotationX', 'sceneRotationY', 'sceneRotationZ', 'sceneAmplitude'];
    sceneSliders.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('input', throttledUpdateAvatarScene);
    });

    // Accordion behavior: only one settings group open at a time
    // Teams Integration is excluded — it stays open independently
    const settingsGroups = document.querySelectorAll('.sidebar .settings-group:not(#teamsSettingsGroup)');
    settingsGroups.forEach(group => {
        group.addEventListener('toggle', () => {
            if (group.open) {
                settingsGroups.forEach(other => {
                    if (other !== group && other.open) {
                        other.removeAttribute('open');
                    }
                });
            }
        });
    });
}

function setupRangeDisplay(sliderId, displayId, formatter) {
    const slider = document.getElementById(sliderId);
    const display = document.getElementById(displayId);
    if (slider && display) {
        slider.addEventListener('input', () => {
            display.textContent = formatter(slider.value);
        });
    }
}

// ===== Photo Avatar Scene Update =====
let lastSceneUpdate = 0;
const SCENE_THROTTLE_MS = 50;

function throttledUpdateAvatarScene() {
    const now = Date.now();
    if (now - lastSceneUpdate < SCENE_THROTTLE_MS) return;
    lastSceneUpdate = now;
    updateAvatarScene();
}

function updateAvatarScene() {
    if (!isConnected || !ws || ws.readyState !== WebSocket.OPEN) return;
    if (!document.getElementById('isPhotoAvatar')?.checked) return;
    if (!document.getElementById('avatarEnabled')?.checked) return;

    const isCustom = document.getElementById('isCustomAvatar')?.checked || false;
    const avatarName = isCustom
        ? document.getElementById('customAvatarName')?.value || ''
        : document.getElementById('photoAvatarName')?.value || 'Anika';
    const parts = avatarName.split('-');
    const character = parts[0].toLowerCase();
    const style = parts.slice(1).join('-') || undefined;

    const scene = {
        zoom: parseInt(document.getElementById('sceneZoom').value) / 100,
        position_x: parseInt(document.getElementById('scenePositionX').value) / 100,
        position_y: parseInt(document.getElementById('scenePositionY').value) / 100,
        rotation_x: parseInt(document.getElementById('sceneRotationX').value) * Math.PI / 180,
        rotation_y: parseInt(document.getElementById('sceneRotationY').value) * Math.PI / 180,
        rotation_z: parseInt(document.getElementById('sceneRotationZ').value) * Math.PI / 180,
        amplitude: parseInt(document.getElementById('sceneAmplitude').value) / 100,
    };

    const avatar = {
        type: 'photo-avatar',
        model: 'vasa-1',
        character: character,
        scene: scene,
    };
    if (isCustom) {
        avatar.customized = true;
    } else if (style) {
        avatar.style = style;
    }

    ws.send(JSON.stringify({
        type: 'update_scene',
        avatar: avatar,
    }));
}

// ===== Conditional Field Visibility =====
function updateConditionalFields() {
    const mode = document.getElementById('mode').value;
    const model = document.getElementById('model').value;
    const voiceType = document.getElementById('voiceType').value;
    const voiceName = document.getElementById('voiceName').value;
    const avatarEnabledChecked = document.getElementById('avatarEnabled').checked;
    const isPhotoAvatar = document.getElementById('isPhotoAvatar').checked;
    const isCustomAvatar = document.getElementById('isCustomAvatar').checked;
    const turnDetectionType = document.getElementById('turnDetectionType').value;
    const srModel = document.getElementById('srModel').value;

    // Cascaded models
    const cascadedModels = ['gpt-4.1', 'gpt-4.1-mini', 'gpt-4.1-nano', 'gpt-4o', 'gpt-4o-mini', 'phi4-mm', 'phi4-mini'];
    const isCascaded = cascadedModels.includes(model);
    const isRealtime = model && model.includes('realtime');

    // Mode: agent vs model -> show/hide fields
    const isAgent = mode === 'agent' || mode === 'agent-v2';
    show('agentFields', isAgent);
    show('modelField', !isAgent);
    show('instructionsField', !isAgent);
    show('temperatureField', !isAgent);

    // Agent ID vs Agent Name
    show('agentIdField', mode === 'agent');
    show('agentNameField', mode === 'agent-v2');

    // Subscription key vs Entra token (agents = entra, model = subscription key)
    show('subscriptionKeyField', !isAgent);
    show('entraTokenField', isAgent);

    // Cascaded-only fields
    show('srModelField', !isAgent && isCascaded);
    show('recognitionLanguageField', !isAgent && isCascaded && srModel !== 'mai-ears-1');
    show('eouDetectionField', !isAgent && isCascaded);

    // Filler words (semantic VAD)
    show('fillerWordsField', turnDetectionType === 'azure_semantic_vad');

    // Voice type variants
    show('standardVoiceField', voiceType === 'standard');
    show('customVoiceFields', voiceType === 'custom');
    show('personalVoiceFields', voiceType === 'personal');

    // Voice temperature (DragonHD or personal voice)
    const isDragonHD = voiceName && voiceName.includes('DragonHD');
    const isPersonal = voiceType === 'personal';
    show('voiceTempField', isDragonHD || isPersonal);

    // Avatar settings
    show('avatarSettings', avatarEnabledChecked);
    show('standardAvatarField', !isPhotoAvatar && !isCustomAvatar);
    show('photoAvatarField', isPhotoAvatar && !isCustomAvatar);
    show('customAvatarField', isCustomAvatar);
    show('photoAvatarSceneSettings', isPhotoAvatar);
}

function show(id, visible) {
    const el = document.getElementById(id);
    if (el) el.style.display = visible ? '' : 'none';
}

// ===== Sidebar Toggle (mobile) =====
function toggleSidebar() {
    document.getElementById('sidebar').classList.toggle('open');
}

// ===== Chat =====
function addMessage(role, text, isDev = false) {
    if (isDev && !isDeveloperMode) return;
    const messagesEl = document.getElementById('messages');
    const msgDiv = document.createElement('div');
    msgDiv.className = `message ${isDev ? 'dev' : role}`;

    if (!isDev) {
        const roleSpan = document.createElement('div');
        roleSpan.className = 'message-role';
        roleSpan.textContent = role === 'user' ? 'You' : role === 'assistant' ? 'Assistant' : 'System';
        msgDiv.appendChild(roleSpan);
    }

    const contentDiv = document.createElement('div');
    contentDiv.className = 'message-content';
    contentDiv.textContent = text;
    msgDiv.appendChild(contentDiv);

    messagesEl.appendChild(msgDiv);
    scrollChatToBottom();
    updateClearChatButton();
    return contentDiv;
}

function updateLastAssistantMessage(text) {
    const messages = document.querySelectorAll('.message.assistant .message-content');
    if (messages.length > 0) {
        messages[messages.length - 1].textContent = text;
        scrollChatToBottom();
    }
}

function scrollChatToBottom() {
    const chatArea = document.getElementById('chatArea');
    chatArea.scrollTop = chatArea.scrollHeight;
}

function clearChat() {
    const messages = document.getElementById('messages');
    if (messages.children.length === 0) return;
    messages.innerHTML = '';
    updateClearChatButton();
}

function updateClearChatButton() {
    const btn = document.getElementById('clearChatBtn');
    const messages = document.getElementById('messages');
    if (!btn || !messages) return;
    const hasMessages = messages.children.length > 0;
    btn.disabled = !hasMessages;
    btn.style.opacity = hasMessages ? '' : '0.5';
}

// ===== Gather Config =====
function gatherConfig() {
    const mode = document.getElementById('mode').value;
    const model = document.getElementById('model').value;
    const voiceType = document.getElementById('voiceType').value;
    const isPhotoAvatar = document.getElementById('isPhotoAvatar').checked;
    const isCustomAvatar = document.getElementById('isCustomAvatar').checked;

    const voiceSpeed = parseFloat(document.getElementById('voiceSpeed').value) / 100;

    const config = {
        mode: mode,
        model: model,
        voiceType: voiceType,
        voiceName: document.getElementById('voiceName').value,
        voiceSpeed: voiceSpeed,
        voiceTemperature: parseFloat(document.getElementById('voiceTemperature').value),
        voiceDeploymentId: document.getElementById('voiceDeploymentId').value,
        customVoiceName: document.getElementById('customVoiceName').value,
        personalVoiceName: document.getElementById('personalVoiceName').value,
        personalVoiceModel: document.getElementById('personalVoiceModel').value,
        avatarEnabled: document.getElementById('avatarEnabled').checked,
        isPhotoAvatar: isPhotoAvatar,
        isCustomAvatar: isCustomAvatar,
        avatarName: isCustomAvatar
            ? document.getElementById('customAvatarName').value
            : isPhotoAvatar
                ? document.getElementById('photoAvatarName').value
                : document.getElementById('avatarName').value,
        avatarOutputMode: document.getElementById('avatarOutputMode').value,
        avatarBackgroundImageUrl: document.getElementById('avatarBackgroundImageUrl').value,
        useNS: document.getElementById('useNS').checked,
        useEC: document.getElementById('useEC').checked,
        turnDetectionType: document.getElementById('turnDetectionType').value,
        removeFillerWords: document.getElementById('removeFillerWords').checked,
        srModel: document.getElementById('srModel').value,
        recognitionLanguage: document.getElementById('recognitionLanguage').value,
        eouDetectionType: document.getElementById('eouDetectionType').value,
        instructions: document.getElementById('instructions').value,
        temperature: parseFloat(document.getElementById('temperature').value),
        enableProactive: document.getElementById('enableProactive').checked,
        // Agent fields
        agentId: document.getElementById('agentId').value,
        agentName: document.getElementById('agentName').value,
        agentProjectName: document.getElementById('agentProjectName').value,
        // Tool toggles
        tools: [
            ...(document.getElementById('toolGetTime')?.checked ? ['get_time'] : []),
            ...(document.getElementById('toolGetWeather')?.checked ? ['get_weather'] : []),
            ...(document.getElementById('toolCalculate')?.checked ? ['calculate'] : []),
        ],
    };

    // Photo avatar scene settings
    if (isPhotoAvatar) {
        config.photoScene = {
            zoom: parseInt(document.getElementById('sceneZoom').value),
            positionX: parseInt(document.getElementById('scenePositionX').value),
            positionY: parseInt(document.getElementById('scenePositionY').value),
            rotationX: parseInt(document.getElementById('sceneRotationX').value),
            rotationY: parseInt(document.getElementById('sceneRotationY').value),
            rotationZ: parseInt(document.getElementById('sceneRotationZ').value),
            amplitude: parseInt(document.getElementById('sceneAmplitude').value),
        };
    }

    return config;
}

// ===== Connection =====
async function toggleConnection() {
    if (isConnecting) return;
    if (isConnected) {
        await disconnect();
    } else {
        await connectSession();
    }
}

async function connectSession() {
    const endpoint = document.getElementById('endpoint').value.trim();
    const mode = document.getElementById('mode').value;
    const isAgent = mode === 'agent' || mode === 'agent-v2';

    if (!endpoint) {
        addMessage('system', 'Please enter Azure AI Services Endpoint');
        return;
    }

    // Validate credentials
    const apiKey = document.getElementById('apiKey')?.value.trim();
    const entraToken = document.getElementById('entraToken')?.value.trim();

    if (isAgent && !entraToken) {
        addMessage('system', 'Please enter Entra ID Token');
        return;
    }

    setConnecting(true);
    const connectingMsgEl = addMessage('system', 'Connecting to session...');
    if (connectingMsgEl) connectingMsgEl.closest?.('.message')?.setAttribute('id', 'connectingStatusMsg');

    try {
        // Open WebSocket to Python backend
        const protocol = location.protocol === 'https:' ? 'wss:' : 'ws:';
        ws = new WebSocket(`${protocol}//${location.host}/ws/${clientId}`);

        ws.onopen = () => {
            const config = gatherConfig();
            // Send credentials to server
            config.endpoint = endpoint;
            if (isAgent) {
                config.entraToken = entraToken;
            } else {
                config.apiKey = apiKey;
            }
            ws.send(JSON.stringify({ type: 'start_session', config }));
        };

        ws.onmessage = (event) => {
            const msg = JSON.parse(event.data);
            handleServerMessage(msg);
        };

        ws.onerror = (err) => {
            console.error('WebSocket error', err);
            addMessage('system', 'WebSocket error');
            setConnecting(false);
        };

        ws.onclose = () => {
            console.log('WebSocket closed');
            if (isConnected) {
                addMessage('system', 'Disconnected');
            }
            handleDisconnect();
        };

    } catch (err) {
        console.error('Connection error', err);
        addMessage('system', 'Failed to connect: ' + err.message);
        setConnecting(false);
    }
}

async function disconnect() {
    if (ws && ws.readyState === WebSocket.OPEN) {
        ws.send(JSON.stringify({ type: 'stop_session' }));
    }
    handleDisconnect();
}

function handleDisconnect() {
    isConnected = false;
    isConnecting = false;
    isRecording = false;
    audioChunksSent = 0;
    avatarEnabled = false;
    teamsAutoConnected = false;

    stopAudioCapture();
    stopAudioPlayback();
    cleanupWebRTC();
    cleanupWebSocketVideo();
    updateSoundWaveAnimation();

    // Prepare next peer connection for faster reconnection
    if (cachedIceServers) {
        preparePeerConnection(cachedIceServers);
    }

    if (ws) {
        try { ws.close(); } catch (e) {}
        ws = null;
    }

    updateConnectionUI();
    updateDeveloperModeLayout();
}

// ===== Handle Server Messages =====
function handleServerMessage(msg) {
    const type = msg.type;

    switch (type) {
        case 'session_started':
            onSessionStarted(msg);
            break;
        case 'session_error':
            addMessage('system', 'Error: ' + (msg.error || 'Unknown error'));
            setConnecting(false);
            break;
        case 'ice_servers':
            // Only setup WebRTC when avatar output mode is webrtc
            if (avatarOutputMode === 'webrtc') {
                setupWebRTC(msg.iceServers);
            }
            break;
        case 'avatar_sdp_answer':
            handleAvatarSdpAnswer(msg.serverSdp);
            break;
        case 'audio_data':
            handleAudioDelta(msg.data);
            break;
        case 'transcript_done':
            if (msg.role === 'user') {
                // Update existing placeholder by itemId, or add new message
                const itemId = msg.itemId;
                if (itemId) {
                    const existing = document.querySelector(`.message.user[data-item-id="${itemId}"] .message-content`);
                    if (existing) {
                        existing.textContent = msg.transcript;
                        scrollChatToBottom();
                        break;
                    }
                }
                addMessage('user', msg.transcript);
            } else if (msg.role === 'assistant') {
                // Finalize the streaming assistant message (don't create a new one)
                if (msg.transcript) {
                    const assistantMsgs = document.querySelectorAll('.message.assistant .message-content');
                    if (assistantMsgs.length > 0) {
                        assistantMsgs[assistantMsgs.length - 1].textContent = msg.transcript;
                    }
                    pendingAssistantText = '';
                }
            }
            break;
        case 'transcript_delta':
            if (msg.role === 'assistant') {
                onAssistantDelta(msg.delta);
            }
            break;
        case 'text_delta':
            onAssistantDelta(msg.delta);
            break;
        case 'text_done':
            // Text response complete - already accumulated via deltas
            break;
        case 'speech_started':
            onSpeechStarted(msg.itemId);
            break;
        case 'speech_stopped':
            onSpeechStopped();
            break;
        case 'response_created':
            pendingAssistantText = '';
            addMessage('assistant', '');
            isSpeaking = true;
            break;
        case 'response_done':
            isSpeaking = false;
            // Don't stop play-chunk animation here - the animation loop
            // will self-terminate when all buffered audio finishes playing
            break;
        case 'session_closed':
            addMessage('system', 'Session closed');
            handleDisconnect();
            break;
        case 'avatar_connecting':
            addMessage('system', 'Avatar connecting...');
            break;
        case 'video_data':
            handleVideoChunk(msg.delta);
            break;
        case 'stop_playback':
            // Server interrupted playback (e.g. barge-in) — flush local buffers
            stopAudioPlayback();
            isSpeaking = false;
            addMessage('system', '[Playback stopped' + (msg.reason ? ': ' + msg.reason : '') + ']', true);
            break;
        case 'function_call_started':
            addMessage('system', `[Tool] ${msg.functionName} called (id: ${msg.callId})`, true);
            break;
        case 'function_call_result':
            addMessage('system', `[Tool] ${msg.functionName} → ${JSON.stringify(msg.result)}`, true);
            break;
        case 'function_call_error':
            addMessage('system', `[Tool] ${msg.functionName} error: ${msg.error}`, true);
            break;
        case 'audio_done':
            // Server audio output complete for this turn
            addMessage('system', '[Audio done]', true);
            break;
        case 'error':
            addMessage('system', 'Server error: ' + (msg.error || msg.message || 'Unknown error'));
            break;
        case 'conversation_item':
            // User/assistant text items forwarded by the backend
            if (msg.role === 'assistant' && msg.text) {
                addMessage('assistant', msg.text);
            }
            break;
        default:
            // Log unknown events in dev mode
            if (isDeveloperMode) {
                console.log('Unhandled:', type, msg);
            }
    }
}

function onAssistantDelta(text) {
    pendingAssistantText += text;
    const messages = document.querySelectorAll('.message.assistant .message-content');
    if (messages.length > 0) {
        messages[messages.length - 1].textContent = pendingAssistantText;
        scrollChatToBottom();
    } else {
        // Fallback: create new message if none exists
        addMessage('assistant', pendingAssistantText);
    }
}

async function onSessionStarted(msg) {
    isConnected = true;
    isConnecting = false;
    updateConnectionUI();

    // Replace the 'Connecting...' placeholder with the confirmed session message.
    const connectingMsg = document.getElementById('connectingStatusMsg');
    if (connectingMsg) {
        connectingMsg.querySelector('.message-content').textContent =
            'Session started — click the microphone button to begin.';
        connectingMsg.removeAttribute('id');
    }

    // Show appropriate content area
    avatarEnabled = msg.config?.avatarEnabled || false;
    avatarOutputMode = msg.config?.avatarOutputMode || 'webrtc';
    const isPhotoAvatarSession = document.getElementById('isPhotoAvatar')?.checked || false;
    const avatarContainer = document.getElementById('avatarVideoContainer');
    if (avatarContainer) {
        avatarContainer.classList.toggle('photo-avatar', isPhotoAvatarSession);
    }
    updateDeveloperModeLayout();

    // If avatar is enabled with websocket mode, set up MediaSource video playback
    if (avatarEnabled && avatarOutputMode === 'websocket') {
        setupWebSocketVideoPlayback(isPhotoAvatarSession);
    }

    // Show record button for non-dev mode
    document.getElementById('recordContainer').style.display = '';

    // Start audio capture but leave mic off by default
    await startAudioCapture();
    isRecording = false;
    stopRecordAnimation();
    resetVolumeCircle();
    updateMicUI();

    // If this session was auto-started by a Teams join, bridge audio both ways
    if (teamsAutoConnected && teamsCall && teamsCall.state === 'Connected') {
        teamsAutoConnected = false;
        _bridgeAvatarAudioToTeams();
        startTeamsAudioCapture();
    } else if (teamsCall && teamsCall.state === 'Connected') {
        // Voice Live was manually reconnected while Teams was already active.
        // Reset stale bridge state so new nodes can be wired up cleanly.
        _teamsBridging = false;
        if (playbackContext?._teamsBridgeNode) {
            try { playbackContext._teamsBridgeNode.disconnect(); } catch (_) {}
            playbackContext._teamsBridgeNode = null;
        }
        if (playbackContext?._webrtcBridgeSource) {
            try { playbackContext._webrtcBridgeSource.disconnect(); } catch (_) {}
            playbackContext._webrtcBridgeSource = null;
        }

        if (avatarEnabled) {
            // Avatar mode: audio arrives via WebRTC ontrack → _bridgeWebRTCAudioToTeams.
            // Restore the silent oscillator to keep the Teams audio track live until
            // the new ontrack event fires.
            if (_teamsBridgeDest && playbackContext && !_teamsBridgeSilentGain) {
                const silentOsc = playbackContext.createOscillator();
                _teamsBridgeSilentGain = playbackContext.createGain();
                _teamsBridgeSilentGain.gain.value = 0;
                silentOsc.connect(_teamsBridgeSilentGain);
                _teamsBridgeSilentGain.connect(_teamsBridgeDest);
                silentOsc.start();
            }
            // If the video track is already available (reconnect case where ontrack
            // already fired), start it now instead of waiting for another ontrack.
            if (_teamsAvatarVideoTrack && _teamsAvatarVideoTrack.readyState !== 'ended') {
                _teamsVideoPending = false;
                const videoChk = document.getElementById('teamsSendVideo');
                if (videoChk) videoChk.checked = true;
                _startTeamsVideo();
            }
        } else {
            // Non-avatar mode: audio arrives as PCM via handleAudioDelta.
            // Re-establish the PCM capture bridge immediately.
            _bridgeAvatarAudioToTeams();
        }
        startTeamsAudioCapture();
    }
}

// ===== UI State =====
function setConnecting(connecting) {
    isConnecting = connecting;
    updateConnectionUI();
}

function updateConnectionUI() {
    const btn = document.getElementById('connectBtn');
    const text = document.getElementById('connectBtnText');

    btn.classList.remove('connected', 'connecting');
    if (isConnected) {
        btn.classList.add('connected');
        text.textContent = 'Disconnect';
    } else if (isConnecting) {
        btn.classList.add('connecting');
        text.textContent = 'Connecting...';
    } else {
        text.textContent = 'Connect';
    }

    // Disable connect button while connecting
    btn.disabled = isConnecting;

    // Scene Settings title: show "(Live Adjustable)" when connected
    const sceneTitle = document.getElementById('sceneSettingsTitle');
    if (sceneTitle) {
        sceneTitle.textContent = isConnected ? 'Scene Settings (Live Adjustable)' : 'Scene Settings';
    }

    // Update all control disabled states
    updateControlStates();

    // Mic buttons
    updateMicUI();
}

// ===== Control Enable/Disable States =====
// Controls that should be disabled when connected (locked during session)
const SETTINGS_CONTROLS = [
    // Connection Settings
    'mode', 'endpoint', 'apiKey', 'entraToken',
    'agentProjectName', 'agentId', 'agentName', 'model',
    // Conversation Settings
    'srModel', 'recognitionLanguage',
    'useNS', 'useEC', 'turnDetectionType', 'removeFillerWords',
    'eouDetectionType', 'instructions', 'enableProactive',
    'temperature', 'voiceTemperature', 'voiceSpeed',
    // Voice Configuration
    'voiceType', 'voiceDeploymentId', 'customVoiceName',
    'personalVoiceName', 'personalVoiceModel', 'voiceName',
    // Avatar Configuration
    'avatarEnabled', 'isPhotoAvatar', 'avatarOutputMode',
    'isCustomAvatar', 'avatarName', 'photoAvatarName',
    'customAvatarName', 'avatarBackgroundImageUrl',
];

// Controls that should be disabled when NOT connected (chat interaction)
const CHAT_CONTROLS = [
    'textInput',
];

function updateControlStates() {
    // Disable all settings controls when connected
    for (const id of SETTINGS_CONTROLS) {
        const el = document.getElementById(id);
        if (el) el.disabled = isConnected;
    }

    // Disable chat controls when NOT connected
    for (const id of CHAT_CONTROLS) {
        const el = document.getElementById(id);
        if (el) el.disabled = !isConnected;
    }

    // Mic button (developer mode) - disabled when not connected
    const micBtn = document.getElementById('micBtn');
    if (micBtn) micBtn.disabled = !isConnected;

    // Send button - disabled when not connected
    const sendBtns = document.querySelectorAll('.send-btn');
    sendBtns.forEach(btn => btn.disabled = !isConnected);

    // Record button (non-developer mode footer) - disabled when not connected
    const recordBtn = document.getElementById('recordBtn');
    if (recordBtn) recordBtn.disabled = !isConnected;

    // Teams join button - always enabled (does not require Voice Live session)
    const teamsJoinBtn = document.getElementById('teamsJoinBtn');
    if (teamsJoinBtn && !teamsCall) teamsJoinBtn.disabled = false;
}

function updateDeveloperModeLayout() {
    const contentArea = document.getElementById('contentArea');
    const avatarVideoContainer = document.getElementById('avatarVideoContainer');
    const volumeAnimation = document.getElementById('volumeAnimation');
    const chatArea = document.getElementById('chatArea');
    const inputArea = document.getElementById('inputArea');
    const footerArea = document.getElementById('footerArea');

    if (isDeveloperMode) {
        // Developer mode: show input area, hide footer
        inputArea.style.display = '';
        footerArea.style.display = 'none';

        if (isConnected && avatarEnabled) {
            // Avatar + developer: side-by-side layout (avatar + chat)
            contentArea.classList.add('developer-layout');
            avatarVideoContainer.style.display = '';
            chatArea.style.display = '';
            volumeAnimation.style.display = 'none';
        } else if (isConnected) {
            // No avatar + developer: side-by-side layout (robot + chat)
            contentArea.classList.add('developer-layout');
            avatarVideoContainer.style.display = 'none';
            chatArea.style.display = '';
            volumeAnimation.style.display = '';
        } else {
            // Not connected: just show chat
            contentArea.classList.remove('developer-layout');
            avatarVideoContainer.style.display = 'none';
            chatArea.style.display = '';
            volumeAnimation.style.display = 'none';
        }
    } else {
        // Normal mode: show footer, hide input area
        inputArea.style.display = 'none';
        footerArea.style.display = '';
        contentArea.classList.remove('developer-layout');

        if (isConnected && avatarEnabled) {
            // Avatar + normal: only avatar video, no chat
            avatarVideoContainer.style.display = '';
            chatArea.style.display = 'none';
            volumeAnimation.style.display = 'none';
        } else if (isConnected) {
            // No avatar + normal: only robot, no chat
            avatarVideoContainer.style.display = 'none';
            chatArea.style.display = 'none';
            volumeAnimation.style.display = '';
        } else {
            // Not connected: show chat history
            avatarVideoContainer.style.display = 'none';
            chatArea.style.display = '';
            volumeAnimation.style.display = 'none';
        }
    }
}

let soundWaveIntervalId = null;

function updateSoundWaveAnimation() {
    const leftWave = document.getElementById('soundWaveLeft');
    const rightWave = document.getElementById('soundWaveRight');

    if (isConnected && avatarEnabled && isRecording && !isDeveloperMode) {
        // Create sound wave bars if not already present
        if (leftWave && leftWave.children.length === 0) {
            for (let i = 0; i < 10; i++) {
                const bar = document.createElement('div');
                bar.className = 'bar';
                bar.id = `item-${i}`;
                bar.style.height = '2px';
                leftWave.appendChild(bar);
            }
        }
        if (rightWave && rightWave.children.length === 0) {
            for (let i = 10; i < 20; i++) {
                const bar = document.createElement('div');
                bar.className = 'bar';
                bar.id = `item-${i}`;
                bar.style.height = '2px';
                rightWave.appendChild(bar);
            }
        }
        // Start animation
        if (!soundWaveIntervalId) {
            soundWaveIntervalId = setInterval(() => {
                for (let i = 0; i < 20; i++) {
                    const ele = document.getElementById(`item-${i}`);
                    const height = 50 * Math.sin((Math.PI / 20) * i) * Math.random();
                    if (ele) {
                        ele.style.transition = 'height 0.15s ease';
                        ele.style.height = `${Math.max(2, height)}px`;
                    }
                }
            }, 150);
        }
        if (leftWave) leftWave.style.display = '';
        if (rightWave) rightWave.style.display = '';
    } else {
        // Stop animation, hide waves
        if (soundWaveIntervalId) {
            clearInterval(soundWaveIntervalId);
            soundWaveIntervalId = null;
        }
        if (leftWave) leftWave.style.display = 'none';
        if (rightWave) rightWave.style.display = 'none';
    }
}
function updateMicUI() {
    const micBtn = document.getElementById('micBtn');
    const recordBtn = document.getElementById('recordBtn');

    // Toggle recording class
    if (micBtn) micBtn.classList.toggle('recording', isRecording);
    if (recordBtn) recordBtn.classList.toggle('recording', isRecording);

    // Toggle icon visibility: show off-icon when not recording, on-icon when recording
    document.querySelectorAll('.mic-off-icon').forEach(el => {
        el.style.display = isRecording ? 'none' : '';
    });
    document.querySelectorAll('.mic-on-icon').forEach(el => {
        el.style.display = isRecording ? '' : 'none';
    });

    // Update label text
    const label = document.querySelector('.microphone-label');
    if (label) {
        label.textContent = isRecording ? 'Turn off microphone' : 'Turn on microphone';
    }

    // Update sound wave visibility
    updateSoundWaveAnimation();
}

// ===== Audio Capture (24kHz PCM16 via AudioWorklet) =====
async function startAudioCapture() {
    try {
        mediaStream = await navigator.mediaDevices.getUserMedia({
            audio: {
                channelCount: 1,
                sampleRate: 24000,
                echoCancellation: true,
                noiseSuppression: true,
            }
        });
        audioContext = new AudioContext({ sampleRate: 24000 });
        console.log('[Audio] AudioContext created, actual sampleRate:', audioContext.sampleRate);

        const source = audioContext.createMediaStreamSource(mediaStream);
        workletNode = await createPCM16WorkletNode(audioContext, 'pcm16-mic-processor');

        // Create analyser for mic volume visualization
        const micAnalyser = audioContext.createAnalyser();
        micAnalyser.fftSize = 2048;
        micAnalyser.smoothingTimeConstant = 0.85;
        const micDataArray = new Uint8Array(micAnalyser.frequencyBinCount);

        workletNode.port.onmessage = (e) => {
            if (!isConnected || !isRecording || !ws || ws.readyState !== WebSocket.OPEN) return;
            const base64 = arrayBufferToBase64(e.data);
            audioChunksSent++;
            if (audioChunksSent <= 3 || audioChunksSent % 100 === 0) {
                console.log(`[Audio] Sending chunk #${audioChunksSent}, size=${base64.length}`);
            }
            ws.send(JSON.stringify({ type: 'audio_chunk', data: base64 }));
        };

        source.connect(workletNode);
        source.connect(micAnalyser);
        workletNode.connect(audioContext.destination);

        // Store mic analyser so volume animation can use it
        micAnalyserNode = micAnalyser;
        micAnalyserDataArray = micDataArray;
        analyserNode = micAnalyser;
        analyserDataArray = micDataArray;
        startVolumeAnimation('record');

        console.log('[Audio] Capture started (24kHz PCM16)');
    } catch (err) {
        console.error('Audio capture error', err);
        addMessage('system', 'Microphone access denied or not available');
    }
}

function stopAudioCapture() {
    stopRecordAnimation();
    micAnalyserNode = null;
    micAnalyserDataArray = null;
    if (workletNode) { try { workletNode.disconnect(); } catch (e) {} workletNode = null; }
    if (audioContext) { try { audioContext.close(); } catch (e) {} audioContext = null; }
    if (mediaStream) { mediaStream.getTracks().forEach(t => t.stop()); mediaStream = null; }
    resetVolumeCircle();
}

// ===== Audio Playback (24kHz PCM16) =====

/** Ensure playbackContext has an analyserNode connected to its destination. */
function _ensurePlaybackAnalyser() {
    if (!analyserNode || analyserNode.context !== playbackContext) {
        analyserNode = playbackContext.createAnalyser();
        analyserNode.fftSize = 2048;
        analyserNode.smoothingTimeConstant = 0.85;
        analyserDataArray = new Uint8Array(analyserNode.frequencyBinCount);
        analyserNode.connect(playbackContext.destination);
        nextPlaybackTime = 0;
    }
}

function handleAudioDelta(base64Data) {
    if (!base64Data) return;
    if (!playbackContext) {
        playbackContext = new AudioContext({ sampleRate: 24000 });
        nextPlaybackTime = 0;
    }
    // Ensure an analyser is attached — may be missing if playbackContext was
    // created early (e.g. at Teams join) before the analyser was wired up.
    _ensurePlaybackAnalyser();
    // Ensure context is running (may be suspended by browser autoplay policy)
    if (playbackContext.state === 'suspended') {
        playbackContext.resume();
    }
    // If Teams is connected but bridge not yet set up, do it now (guard re-entry).
    // Skip if the WebRTC bridge is already active — avatar mode routes audio via ontrack,
    // not through PCM chunks, so there is nothing useful to capture here.
    if (teamsCall && teamsCall.state === 'Connected'
            && !playbackContext._teamsBridgeNode
            && !playbackContext._webrtcBridgeSource
            && !_teamsBridging) {
        _bridgeAvatarAudioToTeams();
    }
    const arrayBuffer = base64ToArrayBuffer(base64Data);
    const int16 = new Int16Array(arrayBuffer);
    const float32 = new Float32Array(int16.length);
    for (let i = 0; i < int16.length; i++) {
        float32[i] = int16[i] / 32768;
    }
    const buffer = playbackContext.createBuffer(1, float32.length, 24000);
    buffer.getChannelData(0).set(float32);
    const source = playbackContext.createBufferSource();
    source.buffer = buffer;

    if (playbackContext._teamsBridgeNode) {
        // Bridge active: source → captureNode → analyserNode → speakers
        //                                   → _teamsBridgeDest → Teams
        source.connect(playbackContext._teamsBridgeNode);
    } else {
        // No bridge (or bridge still setting up): go direct to speakers
        source.connect(analyserNode);
    }

    const now = playbackContext.currentTime;
    if (nextPlaybackTime < now) nextPlaybackTime = now;
    source.start(nextPlaybackTime);
    nextPlaybackTime += buffer.duration;

    // Start playback volume animation (only if not already running)
    if (!playChunkAnimationFrameId) {
        startVolumeAnimation('play-chunk');
    }
}

function stopAudioPlayback() {
    stopPlayChunkAnimation();
    // Don't close playbackContext if Teams is using it for the audio bridge
    if (playbackContext && !teamsCall) {
        try { playbackContext.close(); } catch (e) {}
        playbackContext = null;
    }
    playbackBufferQueue = [];
    nextPlaybackTime = 0;
    // Switch back to mic analyser if mic is on
    if (isRecording && micAnalyserNode) {
        analyserNode = micAnalyserNode;
        analyserDataArray = micAnalyserDataArray;
        startVolumeAnimation('record');
    } else if (!teamsCall) {
        // Don't null the analyserNode if Teams bridge is using it
        analyserNode = null;
        analyserDataArray = null;
        resetVolumeCircle();
    }
}

// ===== Volume Animation =====
function startVolumeAnimation(animationType) {
    if (animationType === 'record') {
        stopPlayChunkAnimation();
    } else {
        stopPlayChunkAnimation();
        stopRecordAnimation();
    }
    const isRecord = animationType === 'record';
    const calculateVolume = () => {
        if (analyserNode && analyserDataArray) {
            analyserNode.getByteFrequencyData(analyserDataArray);
            const volume = Array.from(analyserDataArray).reduce((acc, v) => acc + v, 0) / analyserDataArray.length;
            updateVolumeCircle(volume, animationType);
        }

        if (isRecord) {
            // Stop record animation if mic was turned off
            if (!isRecording) {
                recordAnimationFrameId = null;
                resetVolumeCircle();
                return;
            }
            recordAnimationFrameId = requestAnimationFrame(calculateVolume);
        } else {
            // For playback: self-terminate when response is done AND audio finished
            if (!isSpeaking && (!playbackContext || playbackContext.currentTime >= nextPlaybackTime + 0.3)) {
                playChunkAnimationFrameId = null;
                // Switch back to mic animation or reset
                if (isRecording && micAnalyserNode) {
                    analyserNode = micAnalyserNode;
                    analyserDataArray = micAnalyserDataArray;
                    startVolumeAnimation('record');
                } else {
                    analyserNode = null;
                    analyserDataArray = null;
                    resetVolumeCircle();
                }
                return;
            }
            playChunkAnimationFrameId = requestAnimationFrame(calculateVolume);
        }
    };
    calculateVolume();
}

function stopRecordAnimation() {
    if (recordAnimationFrameId) {
        cancelAnimationFrame(recordAnimationFrameId);
        recordAnimationFrameId = null;
    }
}

function stopPlayChunkAnimation() {
    if (playChunkAnimationFrameId) {
        cancelAnimationFrame(playChunkAnimationFrameId);
        playChunkAnimationFrameId = null;
    }
}

function stopVolumeAnimation() {
    stopRecordAnimation();
    stopPlayChunkAnimation();
}

function updateVolumeCircle(volume, animationType) {
    const circle = document.getElementById('volumeCircle');
    if (!circle) return;
    const minSize = 160;
    const size = minSize + volume;
    circle.style.backgroundColor = animationType === 'record' ? 'lightgray' : 'lightblue';
    circle.style.width = size + 'px';
    circle.style.height = size + 'px';
}

function resetVolumeCircle() {
    const circle = document.getElementById('volumeCircle');
    if (!circle) return;
    circle.style.width = '';
    circle.style.height = '';
    circle.style.backgroundColor = '';
}

// ===== WebSocket Video Playback (MediaSource Extensions) =====
function setupWebSocketVideoPlayback(isPhotoAvatar) {
    // Clean any existing video
    cleanupWebSocketVideo();
    const container = document.getElementById('avatarVideo');
    if (container) container.innerHTML = '';

    // Create video element
    const videoElement = document.createElement('video');
    videoElement.id = 'ws-video';
    videoElement.autoplay = true;
    videoElement.playsInline = true;

    if (isPhotoAvatar) {
        videoElement.style.borderRadius = '10%';
    }
    videoElement.style.width = 'auto';
    videoElement.style.height = isDeveloperMode ? 'auto' : '';
    videoElement.style.objectFit = 'cover';
    videoElement.style.display = 'block';

    videoElement.addEventListener('canplay', () => {
        videoElement.play().catch(e => console.error('Play error:', e));
    });

    // fMP4 codec: H.264 video + AAC audio
    const FMP4_MIME_CODEC = 'video/mp4; codecs="avc1.42E01E, mp4a.40.2"';

    if (!MediaSource.isTypeSupported(FMP4_MIME_CODEC)) {
        console.error('MediaSource fMP4 codec not supported');
        addMessage('system', 'WebSocket video playback not supported in this browser. Please use WebRTC mode.');
        return;
    }

    mediaSource = new MediaSource();
    mediaSourceObjectURL = URL.createObjectURL(mediaSource);
    videoElement.src = mediaSourceObjectURL;

    mediaSource.addEventListener('sourceopen', () => {
        try {
            if (mediaSource.readyState === 'open') {
                sourceBuffer = mediaSource.addSourceBuffer(FMP4_MIME_CODEC);
                sourceBuffer.addEventListener('updateend', () => {
                    processVideoChunkQueue();
                });
            }
        } catch (e) {
            console.error('Error creating SourceBuffer:', e);
        }
    });

    // Append to container
    if (container) {
        container.appendChild(videoElement);
    } else {
        pendingWsVideoElement = videoElement;
    }
}

function handleVideoChunk(base64Data) {
    if (!base64Data) return;
    videoChunkCount++;
    if (videoChunkCount <= 5 || videoChunkCount % 100 === 0) {
        console.log(`[VIDEO] chunk #${videoChunkCount}, length=${base64Data.length}, mediaSource=${mediaSource?.readyState}, sourceBuffer=${!!sourceBuffer}`);
    }
    try {
        videoChunksQueue.push(base64ToArrayBuffer(base64Data));
        // Cap the queue to prevent unbounded memory growth when the SourceBuffer
        // is temporarily busy (e.g. slow device or long network stall).
        const MAX_VIDEO_QUEUE = 300;
        if (videoChunksQueue.length > MAX_VIDEO_QUEUE) {
            const dropped = videoChunksQueue.length - MAX_VIDEO_QUEUE;
            videoChunksQueue.splice(0, dropped);
            console.warn(`[VIDEO] Queue overflow — dropped ${dropped} stale chunk(s)`);
        }
        processVideoChunkQueue();
    } catch (e) {
        console.error('Error handling video chunk:', e);
    }
}

function processVideoChunkQueue() {
    if (!sourceBuffer || sourceBuffer.updating || !mediaSource || mediaSource.readyState !== 'open') {
        return;
    }
    const next = videoChunksQueue.shift();
    if (!next) return;
    try {
        sourceBuffer.appendBuffer(next);
    } catch (e) {
        console.error('Error appending video chunk:', e);
    }
}

function cleanupWebSocketVideo() {
    videoChunksQueue = [];
    videoChunkCount = 0;
    if (sourceBuffer && mediaSource) {
        try {
            if (mediaSource.readyState === 'open' && !sourceBuffer.updating) {
                mediaSource.endOfStream();
            }
        } catch (e) {
            console.error('Error ending MediaSource stream:', e);
        }
    }
    sourceBuffer = null;
    mediaSource = null;
    if (mediaSourceObjectURL) {
        URL.revokeObjectURL(mediaSourceObjectURL);
        mediaSourceObjectURL = null;
    }
    pendingWsVideoElement = null;
}

// ===== WebRTC for Avatar =====

/**
 * Wires the ontrack handler onto a peer connection.
 * Called by both preparePeerConnection() (pre-warmed PC) and setupWebRTC() (live PC).
 * @param {RTCPeerConnection} pc        - The peer connection to configure.
 * @param {HTMLElement|null}  container - The #avatarVideo DOM element to append media players into.
 */
function _setupPCOntrack(pc, container) {
    pc.ontrack = (event) => {
        const mediaPlayer = document.createElement(event.track.kind);
        mediaPlayer.id = event.track.kind;
        mediaPlayer.srcObject = event.streams[0];
        mediaPlayer.autoplay = false;
        mediaPlayer.addEventListener('loadeddata', () => { mediaPlayer.play(); });
        if (container) container.appendChild(mediaPlayer);

        if (event.track.kind === 'video') {
            avatarVideoElement = mediaPlayer;
            mediaPlayer.style.width = '0.1%';
            mediaPlayer.style.height = '0.1%';
            mediaPlayer.onplaying = () => {
                setTimeout(() => {
                    mediaPlayer.style.width = '';
                    mediaPlayer.style.height = '';
                }, 0);
            };
            _bridgeWebRTCVideoToTeams(event.track);
        } else if (event.track.kind === 'audio') {
            _bridgeWebRTCAudioToTeams(event.streams[0], mediaPlayer);
        }
    };
}

// Prepare a peer connection ahead of time so ICE candidates are pre-gathered.
// This avoids the ICE gathering delay when the user starts a new session.
function preparePeerConnection(iceServers) {
    const iceConfig = iceServers.map(s => ({
        urls: s.urls,
        username: s.username || undefined,
        credential: s.credential || undefined,
    }));

    const pc = new RTCPeerConnection({ iceServers: iceConfig });
    let iceGatheringDone = false;

    // NOTE: do NOT call _setupPCOntrack here — the #avatarVideo container
    // reference would be stale by the time this PC is actually dequeued
    // and used. _setupPCOntrack is called in setupWebRTC() instead.

    pc.onicecandidate = (event) => {
        if (!event.candidate && !iceGatheringDone) {
            iceGatheringDone = true;
            if (_icePrepareTimeoutId !== null) {
                clearTimeout(_icePrepareTimeoutId);
                _icePrepareTimeoutId = null;
            }
            peerConnectionQueue.push(pc);
            console.log('[' + new Date().toISOString() + '] ICE gathering done, new peer connection prepared.');
            // Keep only the latest prepared connection
            if (peerConnectionQueue.length > 1) {
                const old = peerConnectionQueue.shift();
                try { old.close(); } catch (e) {}
            }
        }
    };

    // Add transceivers for video and audio
    pc.addTransceiver('video', { direction: 'sendrecv' });
    pc.addTransceiver('audio', { direction: 'sendrecv' });

    // Listen for data channel events
    pc.addEventListener('datachannel', (event) => {
        const dataChannel = event.channel;
        dataChannel.onmessage = (e) => {
            console.log('[' + new Date().toISOString() + '] WebRTC event received: ' + e.data);
        };
        dataChannel.onclose = () => {
            console.log('Data channel closed');
        };
    });
    pc.createDataChannel('eventChannel');

    pc.createOffer().then(offer => {
        return pc.setLocalDescription(offer);
    }).then(() => {
        // Timeout fallback: if ICE gathering hasn't completed after 10 seconds, push anyway
        _icePrepareTimeoutId = setTimeout(() => {
            _icePrepareTimeoutId = null;
            if (!iceGatheringDone) {
                iceGatheringDone = true;
                peerConnectionQueue.push(pc);
                console.log('[' + new Date().toISOString() + '] ICE gathering timed out, peer connection prepared with available candidates.');
                if (peerConnectionQueue.length > 1) {
                    const old = peerConnectionQueue.shift();
                    try { old.close(); } catch (e) {}
                }
            }
        }, 10000);
    }).catch(err => {
        console.error('preparePeerConnection offer error', err);
    });
}

function setupWebRTC(iceServers) {
    if (peerConnection) cleanupWebRTC();

    // Cache ICE servers for future peer connection preparation
    cachedIceServers = iceServers;

    // Clear existing video container
    const container = document.getElementById('avatarVideo');
    if (container) container.innerHTML = '';

    if (peerConnectionQueue.length > 0) {
        // Use cached peer connection with pre-gathered ICE candidates
        peerConnection = peerConnectionQueue.shift();
        console.log('[' + new Date().toISOString() + '] Using cached peer connection with pre-gathered ICE candidates.');

        // Wire track handler now against the live container (not the stale prep-time one)
        _setupPCOntrack(peerConnection, container);

        // Send SDP offer immediately (no need to wait for ICE gathering)
        const sdpJson = JSON.stringify(peerConnection.localDescription);
        const sdpBase64 = btoa(sdpJson);
        console.log('[SDP] Sending cached base64 SDP, starts with:', sdpBase64.substring(0, 40));
        ws.send(JSON.stringify({ type: 'avatar_sdp_offer', clientSdp: sdpBase64 }));
        console.log('[WebRTC] Cached SDP offer sent (base64)');

        // Prepare next peer connection for future use
        preparePeerConnection(iceServers);
        return;
    }

    // No cached peer connection available (first connection), create one from scratch
    const iceConfig = iceServers.map(s => ({
        urls: s.urls,
        username: s.username || undefined,
        credential: s.credential || undefined,
    }));

    peerConnection = new RTCPeerConnection({ iceServers: iceConfig });

    // Wire up track handler (video → avatar display + Teams bridge; audio → Teams bridge)
    _setupPCOntrack(peerConnection, container);

    let iceGatheringDone = false;
    peerConnection.onicecandidate = (event) => {
        if (!event.candidate && !iceGatheringDone) {
            iceGatheringDone = true;
            // ICE gathering complete, send SDP offer now
            const sdpJson = JSON.stringify(peerConnection.localDescription);
            const sdpBase64 = btoa(sdpJson);
            console.log('[SDP] Sending base64 SDP, starts with:', sdpBase64.substring(0, 40));
            ws.send(JSON.stringify({ type: 'avatar_sdp_offer', clientSdp: sdpBase64 }));
            console.log('[WebRTC] SDP offer sent (base64)');
        }
    };

    // Add transceivers for video and audio
    peerConnection.addTransceiver('video', { direction: 'sendrecv' });
    peerConnection.addTransceiver('audio', { direction: 'sendrecv' });

    // Listen for data channel events
    peerConnection.addEventListener('datachannel', (event) => {
        const dataChannel = event.channel;
        dataChannel.onmessage = (e) => {
            console.log('[' + new Date().toISOString() + '] WebRTC event received: ' + e.data);
        };
        dataChannel.onclose = () => {
            console.log('Data channel closed');
        };
    });
    peerConnection.createDataChannel('eventChannel');

    peerConnection.createOffer().then(offer => {
        return peerConnection.setLocalDescription(offer);
    }).then(() => {
        // Timeout fallback: send SDP after 10 seconds if ICE gathering hasn't completed
        _iceSetupTimeoutId = setTimeout(() => {
            _iceSetupTimeoutId = null;
            if (!iceGatheringDone) {
                iceGatheringDone = true;
                const sdpJson = JSON.stringify(peerConnection.localDescription);
                const sdpBase64 = btoa(sdpJson);
                console.log('[SDP] Sending base64 SDP (timeout), starts with:', sdpBase64.substring(0, 40));
                ws.send(JSON.stringify({ type: 'avatar_sdp_offer', clientSdp: sdpBase64 }));
                console.log('[WebRTC] SDP offer sent after timeout (base64)');
            }
        }, 10000);
    }).catch(err => {
        console.error('WebRTC offer error', err);
        addMessage('system', 'WebRTC setup failed');
    });

    // Prepare a peer connection for future use
    preparePeerConnection(iceServers);
}

function handleAvatarSdpAnswer(serverSdpBase64) {
    if (!peerConnection || !serverSdpBase64) return;
    try {
        // Server SDP is base64-encoded JSON: {"type":"answer","sdp":"..."}
        const serverSdpJson = atob(serverSdpBase64);
        const serverSdpObj = JSON.parse(serverSdpJson);
        peerConnection.setRemoteDescription(new RTCSessionDescription(serverSdpObj)).then(() => {
            console.log('[WebRTC] Remote SDP set');
        }).catch(err => {
            console.error('SDP answer error', err);
        });
    } catch (e) {
        console.error('Failed to parse server SDP', e);
    }
}

function cleanupWebRTC() {
    // Cancel any pending ICE timeout timers
    if (_icePrepareTimeoutId !== null) {
        clearTimeout(_icePrepareTimeoutId);
        _icePrepareTimeoutId = null;
    }
    if (_iceSetupTimeoutId !== null) {
        clearTimeout(_iceSetupTimeoutId);
        _iceSetupTimeoutId = null;
    }

    // Stop the canvas draw loop immediately so the rAF chain dies.
    if (_teamsCanvasSrcVideo) {
        _teamsCanvasSrcVideo._teamsActive = false;
        try { _teamsCanvasSrcVideo.pause(); } catch (_) {}
        _teamsCanvasSrcVideo.remove();
        _teamsCanvasSrcVideo = null;
    }
    _teamsCanvas = null;

    // If Teams video was active, tell ACS to stop it now.
    // We MUST call stopVideo() before nulling _teamsLocalVideoStream so that
    // the next _startTeamsVideo() call can issue startVideo() without ACS
    // throwing "video is already started".
    if (teamsCall && _teamsLocalVideoStream) {
        const streamToStop = _teamsLocalVideoStream;
        _teamsLocalVideoStream = null;  // null first so no other code re-uses it
        teamsCall.stopVideo(streamToStop).catch(() => {});  // fire-and-forget; ignore if already stopped
    } else {
        _teamsLocalVideoStream = null;
    }
    _teamsAvatarVideoTrack = null;  // old WebRTC track is dead after reconnect

    if (peerConnection) {
        try { peerConnection.close(); } catch (e) {}
        peerConnection = null;
    }
    if (avatarVideoElement) {
        avatarVideoElement.srcObject = null;
        avatarVideoElement = null;
    }
    const container = document.getElementById('avatarVideo');
    if (container) container.innerHTML = '';
}

// ===== Mic Toggle =====
function toggleMicrophone() {
    if (!isConnected) return;
    isRecording = !isRecording;
    updateMicUI();
    // Start/stop volume animation based on mic state
    if (isRecording && micAnalyserNode) {
        analyserNode = micAnalyserNode;
        analyserDataArray = micAnalyserDataArray;
        startVolumeAnimation('record');
    } else if (!isRecording) {
        stopRecordAnimation();
        resetVolumeCircle();
    }
}

// ===== Send Text =====
function sendTextMessage() {
    const input = document.getElementById('textInput');
    const text = input.value.trim();
    if (!text || !isConnected || !ws) return;

    addMessage('user', text);
    ws.send(JSON.stringify({ type: 'send_text', text }));
    input.value = '';
}

// ===== Speech Events (sound wave animation) =====
function onSpeechStarted(itemId) {
    isSpeaking = true;
    // Stop assistant audio playback (barge-in) in speech-only mode
    stopAudioPlayback();
    // For WebSocket avatar mode: fully reset the MediaSource pipeline so the
    // stale decoder state is torn down and the video element stops rendering
    // the old response.  Simply clearing the queue is not enough — the
    // SourceBuffer may already have decoded frames queued in the browser's
    // media pipeline.  Reinitialising forces a clean slate.
    if (avatarOutputMode === 'websocket') {
        const isPhoto = document.getElementById('avatarType')?.value === 'photo';
        cleanupWebSocketVideo();
        setupWebSocketVideoPlayback(isPhoto);
    }
    // Add user placeholder message (will be updated when transcription completes)
    if (itemId) {
        const contentDiv = addMessage('user', '...');
        if (contentDiv) {
            contentDiv.closest('.message').setAttribute('data-item-id', itemId);
        }
    }
}

function onSpeechStopped() {
    pendingAssistantText = '';
    isSpeaking = false;
}

// ===== Utilities =====
function arrayBufferToBase64(buffer) {
    // Process in 8 kB chunks to avoid call-stack overflow on large buffers
    const bytes = new Uint8Array(buffer);
    const CHUNK = 8192;
    let binary = '';
    for (let i = 0; i < bytes.length; i += CHUNK) {
        binary += String.fromCharCode.apply(null, bytes.subarray(i, i + CHUNK));
    }
    return btoa(binary);
}

function base64ToArrayBuffer(base64) {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
        bytes[i] = binary.charCodeAt(i);
    }
    return bytes.buffer;
}

// ===== Shared Audio Utility =====
/**
 * Creates an AudioWorklet node that encodes float32 audio to PCM16 chunks
 * and posts each chunk via its MessagePort.
 *
 * @param {AudioContext} ctx      - The AudioContext to create the node in.
 * @param {string}       name     - Unique processor name (must be unique per AudioContext).
 * @returns {Promise<AudioWorkletNode>}
 */
async function createPCM16WorkletNode(ctx, name) {
    const processorCode = `
class PCM16Processor extends AudioWorkletProcessor {
    constructor() {
        super();
        this.bufferSize = 2400; // 100ms at 24 kHz
        this.buffer = new Float32Array(this.bufferSize);
        this.offset = 0;
    }
    process(inputs) {
        const input = inputs[0];
        if (!input || !input[0]) return true;
        const data = input[0];
        for (let i = 0; i < data.length; i++) {
            this.buffer[this.offset++] = data[i];
            if (this.offset >= this.bufferSize) {
                const pcm16 = new Int16Array(this.bufferSize);
                for (let j = 0; j < this.bufferSize; j++) {
                    const s = Math.max(-1, Math.min(1, this.buffer[j]));
                    pcm16[j] = s < 0 ? s * 0x8000 : s * 0x7FFF;
                }
                this.port.postMessage(pcm16.buffer, [pcm16.buffer]);
                this.buffer = new Float32Array(this.bufferSize);
                this.offset = 0;
            }
        }
        return true;
    }
}
registerProcessor('${name}', PCM16Processor);
`;
    const blob = new Blob([processorCode], { type: 'application/javascript' });
    const url = URL.createObjectURL(blob);
    await ctx.audioWorklet.addModule(url);
    URL.revokeObjectURL(url);
    return new AudioWorkletNode(ctx, name);
}

// ===== Teams / Azure Communication Services Integration =====
let teamsCallAgent = null;
let teamsCall = null;
let teamsCallClient = null;
let teamsInAudioCtx = null;    // AudioContext for Teams→VoiceLive capture
let teamsInWorkletNode = null; // AudioWorklet PCM16 encoder for incoming Teams audio
let teamsAutoConnected = false; // true when VoiceLive session was auto-started by Teams join
let _teamsBridgeDest = null;       // MediaStreamDestination (in playbackContext) ACS holds from join
let _teamsBridgeSilentGain = null; // GainNode feeding silence until real audio bridges in
let _teamsBridging = false;        // true while _bridgeAvatarAudioToTeams is setting up (prevents re-entry)
let _teamsLocalVideoStream = null; // ACS LocalVideoStream for avatar video
let _teamsAvatarVideoTrack = null; // The WebRTC video track being sent to Teams
let _teamsVideoPending = false;    // true when checkbox checked but track not yet available
let _teamsCanvasSrcVideo = null;   // Hidden <video> used for canvas reframing
let _teamsCanvas = null;           // Offscreen <canvas> for 16:9 reframe
let _teamsCanvasIntervalId = null; // setInterval id for background-safe canvas draw loop
let _teamsSilentOsc = null;        // OscillatorNode kept alive until leave — must be stopped explicitly
let _teamsTokenCredential = null;  // AzureCommunicationTokenCredential — kept for token refresh
let _teamsTokenRefreshTimer = null; // setTimeout id for proactive token refresh
let _teamsFileAudioCtx = null;     // AudioContext used for file-source audio bridging
let _teamsFileBufferSource = null; // BufferSourceNode for file-source audio bridging
let teamsAudioChunkQueue = [];     // PCM chunks buffered while VoiceLive session is not yet open
let _teamsUserId = null;           // ACS user identity (saved so we can delete it after the call)
let _remoteAudioStreamsHandler = null; // named handler ref so we can call .off() with same reference

/**
 * Schedules a proactive ACS token refresh 5 minutes before the token expires.
 * Calls /api/acs-token and updates the credential via updateToken() so the
 * active call is not interrupted by token expiry (~1 hour default lifetime).
 */
function _scheduleTeamsTokenRefresh(tokenData) {
    if (_teamsTokenRefreshTimer) {
        clearTimeout(_teamsTokenRefreshTimer);
        _teamsTokenRefreshTimer = null;
    }
    if (!tokenData?.expiresOn) return;
    const expiresMs = new Date(tokenData.expiresOn).getTime();
    const refreshAt = expiresMs - Date.now() - 5 * 60 * 1000; // 5 min before expiry
    if (refreshAt <= 0) return; // already expired or too close
    _teamsTokenRefreshTimer = setTimeout(async () => {
        try {
            addMessage('system', '[Teams] Refreshing ACS token...', true);
            const resp = await fetch('/api/acs-token');
            if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
            const data = await resp.json();
            if (_teamsTokenCredential) {
                _teamsTokenCredential.updateToken(data.token);
                addMessage('system', '[Teams] ACS token refreshed ✓', true);
            }
            _scheduleTeamsTokenRefresh(data); // schedule next refresh
        } catch (e) {
            addMessage('system', `[Teams] Token refresh failed: ${e.message}`, true);
        }
    }, Math.max(refreshAt, 0));
}

function setTeamsStatus(msg, isError = false) {
    const el = document.getElementById('teamsStatus');
    if (!el) return;
    el.textContent = msg;
    el.style.display = msg ? 'block' : 'none';
    el.style.color = isError ? '#c00' : '#0a6640';
    if (msg) addMessage('system', `[Teams] ${msg}`, true);
}

async function joinTeamsMeeting() {
    const meetingLink = document.getElementById('teamsMeetingLink')?.value.trim();
    const displayName = document.getElementById('teamsDisplayName')?.value.trim() || 'AI Avatar';

    if (!meetingLink) {
        setTeamsStatus('Please enter a Teams meeting link.', true);
        return;
    }

    if (typeof AzureCommunicationCalling === 'undefined') {
        setTeamsStatus('ACS Calling SDK not loaded. Check your network connection.', true);
        return;
    }

    setTeamsStatus('Fetching ACS token...');
    document.getElementById('teamsJoinBtn').disabled = true;

    try {
        // 1. Get ACS token from backend
        const tokenResp = await fetch('/api/acs-token');
        if (!tokenResp.ok) {
            const err = await tokenResp.json();
            throw new Error(err.detail || 'Failed to get ACS token');
        }
        const tokenData = await tokenResp.json();
        const { userId, token } = tokenData;
        _teamsUserId = userId;
        addMessage('system', `[Teams] ACS token issued for user: ${userId}`, true);

        setTeamsStatus('Initializing ACS call client...');

        // 2. Create ACS CallClient and CallAgent
        const { CallClient, AzureCommunicationTokenCredential } =
            AzureCommunicationCalling;

        teamsCallClient = new CallClient();
        _teamsTokenCredential = new AzureCommunicationTokenCredential(token);
        teamsCallAgent = await teamsCallClient.createCallAgent(
            _teamsTokenCredential,
            { displayName }
        );
        // Schedule proactive token refresh 5 minutes before expiry
        _scheduleTeamsTokenRefresh(tokenData);

        setTeamsStatus('Joining Teams meeting...');

        // 3. Create playbackContext now (if not already) so the bridge MediaStreamDestination
        //    lives in the same AudioContext as the avatar audio. This is the key to a working
        //    single-context audio graph — no cross-context relay needed.
        if (!playbackContext) {
            playbackContext = new AudioContext({ sampleRate: 24000 });
            analyserNode = playbackContext.createAnalyser();
            analyserNode.fftSize = 2048;
            analyserNode.smoothingTimeConstant = 0.85;
            analyserDataArray = new Uint8Array(analyserNode.frequencyBinCount);
            analyserNode.connect(playbackContext.destination);
            nextPlaybackTime = 0;
        }
        // Resume immediately — we are inside a user gesture (button click) so this will succeed.
        // Without this the context may stay 'suspended' and the MediaStreamDestination produces silence.
        if (playbackContext.state === 'suspended') {
            await playbackContext.resume();
        }
        addMessage('system', `[Teams] playbackContext state after resume: ${playbackContext.state}`, true);

        // Create the MediaStreamDestination that ACS will hold for the lifetime of the call.
        // Initially we feed it a zero-gain oscillator so the track stays live.
        // When the avatar starts speaking, _bridgeAvatarAudioToTeams() disconnects the
        // silent gain and connects captureNode — all within the same AudioContext.
        _teamsBridgeDest = playbackContext.createMediaStreamDestination();
        _teamsSilentOsc = playbackContext.createOscillator();
        _teamsBridgeSilentGain = playbackContext.createGain();
        _teamsBridgeSilentGain.gain.value = 0;
        _teamsSilentOsc.connect(_teamsBridgeSilentGain);
        _teamsBridgeSilentGain.connect(_teamsBridgeDest);
        _teamsSilentOsc.start();

        const silentStream = new AzureCommunicationCalling.LocalAudioStream(
            _teamsBridgeDest.stream
        );

        teamsCall = teamsCallAgent.join(
            { meetingLink },
            { audioOptions: { localAudioStreams: [silentStream] } }
        );

        teamsCall.on('stateChanged', () => {
            const state = teamsCall.state;
            setTeamsStatus(`Teams: ${state}`);
            if (state === 'Connected') {
                document.getElementById('teamsJoinBtn').style.display = 'none';
                document.getElementById('teamsLeaveBtn').style.display = 'inline-flex';
                if (!isConnected) {
                    // Auto-start VoiceLive session; bridging happens in onSessionStarted
                    teamsAutoConnected = true;
                    addMessage('system', '[Teams] Auto-connecting to Voice Live...', true);
                    setTeamsStatus('Teams: Connected ✓ — starting AI session...');
                    connectSession();
                } else {
                    // VoiceLive already active — bridge immediately.
                    if (avatarEnabled && avatarVideoElement?.srcObject) {
                        // Avatar mode: WebRTC audio track is already playing.
                        // Route the existing stream through the bridge, exactly as ontrack would.
                        // The audio <video> element is already in the DOM with id="audio".
                        const audioEl = document.getElementById('audio') || avatarVideoElement;
                        _bridgeWebRTCAudioToTeams(avatarVideoElement.srcObject, audioEl);
                    } else {
                        // Non-avatar / PCM mode — set up the PCM capture bridge.
                        _bridgeAvatarAudioToTeams();
                    }
                    startTeamsAudioCapture();
                    // If avatar is active and the WebRTC video track already arrived
                    // before Teams joined, start video now. Auto-check the checkbox so
                    // the UI reflects the state and the user can uncheck to stop.
                    // _bridgeWebRTCVideoToTeams couldn't act earlier because teamsCall was null.
                    if (_teamsAvatarVideoTrack) {
                        _teamsVideoPending = false;
                        const videoChk = document.getElementById('teamsSendVideo');
                        if (videoChk) videoChk.checked = true;
                        _startTeamsVideo();
                    }
                }
            } else if (['Disconnected', 'Disconnecting'].includes(state)) {
                _onTeamsDisconnected();
            }
        });

    } catch (e) {
        console.error('Teams join error:', e);
        setTeamsStatus(`Error: ${e.message}`, true);
        document.getElementById('teamsJoinBtn').disabled = false;
    }
}

// Show/hide the file picker row when audio source radio changes
function onTeamsAudioSourceChange() {
    const val = document.querySelector('input[name="teamsAudioSource"]:checked')?.value;
    const fileRow = document.getElementById('teamsAudioFileRow');
    if (fileRow) fileRow.style.display = val === 'file' ? 'block' : 'none';
    // If the call is already live, re-bridge with the new source immediately
    if (teamsCall && teamsCall.state === 'Connected') {
        _bridgeAvatarAudioToTeams();
    }
}

/**
 * Called from ontrack when the WebRTC video track arrives.
 * If Teams is active, starts sending the avatar video track as a LocalVideoStream
 * into the ACS call. The track is reframed onto a 960x540 16:9 canvas.
 * Auto-starts video whenever Teams is connected — the user can uncheck
 * 'Send avatar video to Teams' to stop it.
 */
async function _bridgeWebRTCVideoToTeams(videoTrack) {
    _teamsAvatarVideoTrack = videoTrack;  // always cache latest track

    // If Teams isn't connected yet, just cache the track.
    // stateChanged→Connected (or onSessionStarted reconnect) will call _startTeamsVideo().
    if (!teamsCall || teamsCall.state !== 'Connected') return;

    // Teams is connected and we have a fresh track — start video automatically.
    // Auto-check the checkbox so the UI reflects the active state.
    _teamsVideoPending = false;
    const videoChk = document.getElementById('teamsSendVideo');
    if (videoChk) videoChk.checked = true;
    await _startTeamsVideo();
}

/** Start sending the cached avatar video track into the active Teams call.
 *  The raw track is reframed onto a configurable canvas so Teams shows the
 *  full avatar without cropping. The avatar is scaled to fit (contain) and
 *  centred on a white background.
 *  Draw loop uses setInterval (not requestAnimationFrame) so it keeps running
 *  when the tab is backgrounded.
 */
async function _startTeamsVideo() {
    if (!teamsCall || teamsCall.state !== 'Connected') return;
    if (!_teamsAvatarVideoTrack || _teamsAvatarVideoTrack.readyState === 'ended') {
        addMessage('system', '[Teams] No avatar video track available yet', true);
        return;
    }
    if (_teamsLocalVideoStream) return;  // already sending

    try {
        // Read resolution/fps from UI, fall back to defaults
        const OUT_W = parseInt(document.getElementById('teamsVideoWidth')?.value)  || 960;
        const OUT_H = parseInt(document.getElementById('teamsVideoHeight')?.value) || 540;
        const OUT_FPS = parseInt(document.getElementById('teamsVideoFps')?.value)  || 25;

        // Hidden video element to read pixels from the WebRTC track
        const srcVideo = document.createElement('video');
        srcVideo.srcObject = new MediaStream([_teamsAvatarVideoTrack]);
        srcVideo.muted = true;
        srcVideo.autoplay = true;
        srcVideo.playsInline = true;
        srcVideo.style.cssText = 'position:fixed;width:1px;height:1px;opacity:0;pointer-events:none';
        document.body.appendChild(srcVideo);

        // Wait until the video has real dimensions (metadata loaded)
        await new Promise(resolve => {
            if (srcVideo.videoWidth > 0) { resolve(); return; }
            srcVideo.onloadedmetadata = resolve;
            srcVideo.play().catch(() => {});
            setTimeout(resolve, 3000); // fallback timeout
        });
        if (srcVideo.videoWidth === 0) await srcVideo.play().catch(() => {});

        const canvas = document.createElement('canvas');
        canvas.width  = OUT_W;
        canvas.height = OUT_H;
        // alpha:false skips alpha compositing — reduces per-frame CPU cost so
        // the encoder gets frames faster and ramp-up is quicker.
        const ctx = canvas.getContext('2d', { alpha: false });

        // Use setInterval instead of requestAnimationFrame so the draw loop
        // keeps running when the tab is in the background (rAF is throttled).
        // Warn once if the tab goes hidden — browsers may throttle setInterval on hidden tabs too.
        const _onVisibilityChange = () => {
            if (document.hidden) {
                addMessage('system', '[Teams] Tab is hidden — canvas draw loop may be throttled by the browser; Teams video quality may degrade.', true);
            }
        };
        document.addEventListener('visibilitychange', _onVisibilityChange, { once: true });

        _teamsCanvasIntervalId = setInterval(() => {
            if (!srcVideo.parentNode) { clearInterval(_teamsCanvasIntervalId); _teamsCanvasIntervalId = null; return; }
            ctx.fillStyle = '#ffffff';
            ctx.fillRect(0, 0, OUT_W, OUT_H);
            const sw = srcVideo.videoWidth  || OUT_W;
            const sh = srcVideo.videoHeight || OUT_H;
            const scale = Math.min(OUT_W / sw, OUT_H / sh);
            const dw = sw * scale;
            const dh = sh * scale;
            ctx.drawImage(srcVideo, (OUT_W - dw) / 2, (OUT_H - dh) / 2, dw, dh);
        }, Math.round(1000 / OUT_FPS));

        // Wait for srcVideo to decode and paint at least one real frame into the
        // canvas before captureStream() so the first encoded frame is a complete
        // I-frame rather than a blank keyframe, which helps Teams ramp up faster.
        await new Promise(resolve => {
            if ('requestVideoFrameCallback' in srcVideo) {
                srcVideo.requestVideoFrameCallback(() => {
                    // Draw the first real frame synchronously before resolving
                    ctx.fillStyle = '#ffffff';
                    ctx.fillRect(0, 0, OUT_W, OUT_H);
                    const sw = srcVideo.videoWidth  || OUT_W;
                    const sh = srcVideo.videoHeight || OUT_H;
                    const scale = Math.min(OUT_W / sw, OUT_H / sh);
                    ctx.drawImage(srcVideo, (OUT_W - sw * scale) / 2, (OUT_H - sh * scale) / 2, sw * scale, sh * scale);
                    resolve();
                });
            } else {
                // Fallback: wait one interval tick
                setTimeout(resolve, Math.round(1000 / OUT_FPS));
            }
        });

        const canvasStream = canvas.captureStream(OUT_FPS);
        const { LocalVideoStream } = AzureCommunicationCalling;
        _teamsLocalVideoStream = new LocalVideoStream(canvasStream);

        // Store refs for cleanup
        _teamsCanvasSrcVideo = srcVideo;
        _teamsCanvas = canvas;

        await teamsCall.startVideo(_teamsLocalVideoStream);
        addMessage('system', `[Teams] Avatar video started ✓ (${srcVideo.videoWidth}×${srcVideo.videoHeight} → ${OUT_W}×${OUT_H} @ ${OUT_FPS}fps)`, true);
        setTeamsStatus('Teams: Connected ✓ (audio + video)');
    } catch (e) {
        addMessage('system', `[Teams] startVideo error: ${e.message}`, true);
        _teamsLocalVideoStream = null;
    }
}

/** Stop sending avatar video to Teams and clean up canvas resources. */
async function _stopTeamsVideo() {
    if (!teamsCall || !_teamsLocalVideoStream) return;
    // Stop the draw loop
    if (_teamsCanvasIntervalId) {
        clearInterval(_teamsCanvasIntervalId);
        _teamsCanvasIntervalId = null;
    }
    try {
        await teamsCall.stopVideo(_teamsLocalVideoStream);
        addMessage('system', '[Teams] Avatar video stopped', true);
    } catch (e) {
        addMessage('system', `[Teams] stopVideo error: ${e.message}`, true);
    }
    _teamsLocalVideoStream = null;
    // Clean up hidden video + canvas used for reframing
    if (_teamsCanvasSrcVideo) {
        try { _teamsCanvasSrcVideo.pause(); } catch (_) {}
        _teamsCanvasSrcVideo.remove();
        _teamsCanvasSrcVideo = null;
    }
    _teamsCanvas = null;
}

/** Called when the 'Send avatar video' checkbox changes. */
async function onTeamsSendVideoChange() {
    const checked = document.getElementById('teamsSendVideo')?.checked;
    if (checked) {
        // If Teams isn't connected yet, just mark pending — the stateChanged
        // handler or _bridgeWebRTCVideoToTeams will start video once both
        // Teams and the WebRTC track are available.
        if (!teamsCall || teamsCall.state !== 'Connected') {
            _teamsVideoPending = true;
            addMessage('system', '[Teams] Video pending — will start when Teams connects', true);
            return;
        }
        // If the WebRTC track hasn't arrived yet, try to get it from the
        // already-playing avatar video element (srcObject track)
        if (!_teamsAvatarVideoTrack && avatarVideoElement?.srcObject) {
            const tracks = avatarVideoElement.srcObject.getVideoTracks();
            if (tracks.length > 0) _teamsAvatarVideoTrack = tracks[0];
        }
        if (_teamsAvatarVideoTrack && _teamsAvatarVideoTrack.readyState !== 'ended') {
            await _startTeamsVideo();
        } else {
            // Track not here yet — set pending flag; _bridgeWebRTCVideoToTeams will pick it up
            _teamsVideoPending = true;
            addMessage('system', '[Teams] Video pending — will start when avatar track arrives', true);
        }
    } else {
        _teamsVideoPending = false;
        await _stopTeamsVideo();
    }
}

function _bridgeWebRTCAudioToTeams(webrtcStream, audioElement) {
    if (!teamsCall || !_teamsBridgeDest || !playbackContext) {
        // Teams not active — let the audio element play normally
        addMessage('system', '[Teams] WebRTC audio: Teams not active, playing locally only', true);
        return;
    }

    addMessage('system', '[Teams] WebRTC audio track received — routing through bridge to Teams', true);

    // Mute the audio element so we don't double-play through speakers;
    // Web Audio graph will handle speaker output instead.
    audioElement.muted = true;

    // Ensure context is running
    if (playbackContext.state === 'suspended') {
        playbackContext.resume();
    }

    // Source: WebRTC audio stream → playbackContext
    const webrtcSource = playbackContext.createMediaStreamSource(webrtcStream);

    // Create a fresh analyser in playbackContext for local speaker + volume viz
    const playbackAnalyser = playbackContext.createAnalyser();
    playbackAnalyser.fftSize = 2048;
    playbackAnalyser.smoothingTimeConstant = 0.85;
    playbackAnalyser.connect(playbackContext.destination);
    analyserNode = playbackAnalyser;
    analyserDataArray = new Uint8Array(playbackAnalyser.frequencyBinCount);

    // webrtcSource → analyser → speakers (local playback)
    webrtcSource.connect(playbackAnalyser);

    // webrtcSource → _teamsBridgeDest (Teams participants hear it)
    // Disconnect silent oscillator first
    if (_teamsBridgeSilentGain) {
        try { _teamsBridgeSilentGain.disconnect(_teamsBridgeDest); } catch (_) {}
        _teamsBridgeSilentGain = null;
    }
    webrtcSource.connect(_teamsBridgeDest);

    // Store on playbackContext so teardown can find and disconnect it
    playbackContext._webrtcBridgeSource = webrtcSource;
    // Clear the PCM bridge node — not needed for WebRTC path
    playbackContext._teamsBridgeNode = null;
    // Ensure re-entrant guard is clear regardless of prior state
    _teamsBridging = false;

    setTeamsStatus('Teams: Connected ✓ (bridging avatar audio via WebRTC)');
}

/**
 * Bridge audio into the active Teams call.
 *
 * Supported sources (selected via #teamsAudioSource radios):
 *   avatar  – avatar PCM output via Web Audio capture node (default)
 *   file    – local audio file decoded and played via Web Audio
 *
 * Calls teamsCall.startAudio(LocalAudioStream) to inject the track into ACS.
 */
async function _bridgeAvatarAudioToTeams() {
    if (!teamsCall) {
        addMessage('system', '[Teams] Bridge skipped — no active call', true);
        return;
    }
    // Prevent re-entrant calls (e.g. multiple handleAudioDelta firing before bridge is set)
    if (_teamsBridging) return;
    _teamsBridging = true;

    const source = document.querySelector('input[name="teamsAudioSource"]:checked')?.value ?? 'avatar';
    addMessage('system', `[Teams] Setting up audio bridge: source=${source}`, true);

    // Tear down any previous file-playback context
    if (_teamsFileAudioCtx) {
        try { _teamsFileAudioCtx.close(); } catch (_) {}
        _teamsFileAudioCtx = null;
    }
    // Stop any previous file buffer source node
    if (_teamsFileBufferSource) {
        try { _teamsFileBufferSource.stop(); _teamsFileBufferSource.disconnect(); } catch (_) {}
        _teamsFileBufferSource = null;
    }
    // Tear down previous avatar bridge node
    if (playbackContext?._teamsBridgeNode) {
        try { playbackContext._teamsBridgeNode.disconnect(); } catch (_) {}
        playbackContext._teamsBridgeNode = null;
    }

    try {
        let mediaTrack;

        if (source === 'file') {
            // ── Audio file ──────────────────────────────────────────────────────
            const fileInput = document.getElementById('teamsAudioFile');
            if (!fileInput?.files?.length) {
                setTeamsStatus('Select an audio file first.', true);
                _teamsBridging = false;
                return;
            }
            const loop = document.getElementById('teamsAudioLoop')?.checked ?? true;
            const arrayBuffer = await fileInput.files[0].arrayBuffer();

            // Decode through playbackContext (24 kHz) to avoid sample-rate mismatch.
            // playbackContext is guaranteed to exist at this point (created during joinTeamsMeeting).
            const audioBuffer = await playbackContext.decodeAudioData(arrayBuffer);

            // Tear down any previous bridge node before rebuilding
            if (playbackContext._teamsBridgeNode) {
                try { playbackContext._teamsBridgeNode.disconnect(); } catch (_) {}
                playbackContext._teamsBridgeNode = null;
            }
            if (_teamsBridgeSilentGain) {
                try { _teamsBridgeSilentGain.disconnect(_teamsBridgeDest); } catch (_) {}
                _teamsBridgeSilentGain = null;
            }

            // Reconnect silent oscillator so the Teams audio track stays live
            // while we decode and start the new file source node.
            if (_teamsSilentOsc && _teamsBridgeDest) {
                const tmpGain = playbackContext.createGain();
                tmpGain.gain.value = 0;
                _teamsBridgeSilentGain = tmpGain;
                _teamsSilentOsc.connect(tmpGain);
                tmpGain.connect(_teamsBridgeDest);
            }

            const fileSource = playbackContext.createBufferSource();
            fileSource.buffer = audioBuffer;
            fileSource.loop = loop;
            // Disconnect temp silent gain now that real audio is about to start
            if (_teamsBridgeSilentGain) {
                try { _teamsBridgeSilentGain.disconnect(_teamsBridgeDest); } catch (_) {}
                _teamsBridgeSilentGain = null;
            }
            fileSource.connect(_teamsBridgeDest);
            fileSource.start(0);
            fileSource.onended = () => {
                if (!loop) addMessage('system', '[Teams] Audio file playback ended', true);
            };
            // Store source so it can be stopped on re-bridge or leave
            _teamsFileBufferSource = fileSource;
            _teamsBridging = false;
            setTeamsStatus(`Teams: Connected ✓ (file: ${fileInput.files[0].name})`);

        } else {
            // ── Avatar audio (default) ───────────────────────────────────────────
            if (!playbackContext || !_teamsBridgeDest) {
                addMessage('system', '[Teams] Bridge error — playbackContext or bridge destination not ready', true);
                _teamsBridging = false;
                return;
            }

            // Build the graph synchronously so _teamsBridgeNode is set before
            // any concurrent handleAudioDelta call checks it.
            //
            //   BufferSource → captureNode → playbackAnalyser → destination  (local speakers)
            //                      ↓
            //              _teamsBridgeDest  (same playbackContext, same MediaStream ACS holds)
            //                      ↓
            //         Teams call ← Teams participants hear the AI

            const captureNode = playbackContext.createGain();
            captureNode.gain.value = 1.0;

            // Always use a playbackContext-owned analyser — the global analyserNode may
            // belong to audioContext (mic) which is a different AudioContext instance.
            // Cross-AudioContext .connect() calls silently fail.
            const playbackAnalyser = playbackContext.createAnalyser();
            playbackAnalyser.fftSize = 2048;
            playbackAnalyser.smoothingTimeConstant = 0.85;
            playbackAnalyser.connect(playbackContext.destination);
            captureNode.connect(playbackAnalyser);
            // Update global so volume animation still works
            analyserNode = playbackAnalyser;
            analyserDataArray = new Uint8Array(playbackAnalyser.frequencyBinCount);

            // Swap out the silent oscillator; connect real audio
            if (_teamsBridgeSilentGain) {
                try { _teamsBridgeSilentGain.disconnect(_teamsBridgeDest); } catch (_) {}
                _teamsBridgeSilentGain = null;
            }
            captureNode.connect(_teamsBridgeDest);

            // Set this SYNCHRONOUSLY before returning so handleAudioDelta
            // immediately routes new BufferSources through captureNode
            playbackContext._teamsBridgeNode = captureNode;
            _teamsBridging = false;

            // Resume context in case browser suspended it (autoplay policy)
            if (playbackContext.state === 'suspended') {
                playbackContext.resume().then(() =>
                    addMessage('system', '[Teams] playbackContext resumed', true)
                );
            }

            // Ensure the call is not muted
            if (teamsCall && teamsCall.isMuted) {
                teamsCall.unmute().catch(e =>
                    addMessage('system', `[Teams] unmute error: ${e.message}`, true)
                );
            }

            const bridgeTrack = _teamsBridgeDest.stream.getAudioTracks()[0];
            addMessage('system', `[Teams] Bridge active — track: ${bridgeTrack?.readyState}, enabled: ${bridgeTrack?.enabled}, ctx: ${playbackContext.state}`, true);
            setTeamsStatus('Teams: Connected ✓ (bridging avatar audio)');
            return;
        }

    } catch (e) {
        _teamsBridging = false;
        console.warn('Teams audio bridge error:', e);
        addMessage('system', `[Teams] Audio bridge error: ${e.message}`, true);
        setTeamsStatus(`Audio bridge error: ${e.message}`, true);
    }
}

/**
 * Shared teardown of all Teams bridge state.
 * Called by both leaveTeamsMeeting() and _onTeamsDisconnected().
 */
function _cleanupTeamsState() {
    // Cancel scheduled token refresh
    if (_teamsTokenRefreshTimer) {
        clearTimeout(_teamsTokenRefreshTimer);
        _teamsTokenRefreshTimer = null;
    }
    _teamsTokenCredential = null;

    // Stop silent oscillator
    if (_teamsSilentOsc) {
        try { _teamsSilentOsc.stop(); } catch (_) {}
        _teamsSilentOsc = null;
    }

    // Tear down avatar audio bridge
    if (playbackContext?._webrtcBridgeSource) {
        try { playbackContext._webrtcBridgeSource.disconnect(); } catch (_) {}
        playbackContext._webrtcBridgeSource = null;
    }
    if (playbackContext?._teamsBridgeNode) {
        try { playbackContext._teamsBridgeNode.disconnect(); } catch (_) {}
        playbackContext._teamsBridgeNode = null;
    }

    // Stop canvas draw loop and clean up video reframe resources
    if (_teamsCanvasIntervalId) {
        clearInterval(_teamsCanvasIntervalId);
        _teamsCanvasIntervalId = null;
    }
    if (_teamsCanvasSrcVideo) {
        _teamsCanvasSrcVideo._teamsActive = false;
        try { _teamsCanvasSrcVideo.pause(); } catch (_) {}
        _teamsCanvasSrcVideo.remove();
        _teamsCanvasSrcVideo = null;
    }
    _teamsCanvas = null;
    _teamsLocalVideoStream = null;
    _teamsAvatarVideoTrack = null;
    _teamsVideoPending = false;
    const videoChk = document.getElementById('teamsSendVideo');
    if (videoChk) videoChk.checked = false;
    _teamsBridgeDest = null;
    _teamsBridgeSilentGain = null;
    _teamsBridging = false;
    teamsAutoConnected = false;
    teamsAudioChunkQueue = [];

    // Clean up file audio context if active
    if (_teamsFileAudioCtx) {
        try { _teamsFileAudioCtx.close(); } catch (_) {}
        _teamsFileAudioCtx = null;
    }
    if (_teamsFileBufferSource) {
        try { _teamsFileBufferSource.stop(); _teamsFileBufferSource.disconnect(); } catch (_) {}
        _teamsFileBufferSource = null;
    }

    // If Voice Live is not active, close the shared playbackContext now
    if (!isConnected && playbackContext) {
        try { playbackContext.close(); } catch (_) {}
        playbackContext = null;
        analyserNode = null;
        analyserDataArray = null;
    }

    stopTeamsAudioCapture();
}

async function leaveTeamsMeeting() {
    document.getElementById('teamsLeaveBtn').style.display = 'none';
    document.getElementById('teamsJoinBtn').style.display = 'inline-flex';

    if (teamsCall) {
        try { await teamsCall.hangUp(); } catch (_) {}
        teamsCall = null;
    }
    if (teamsCallAgent) {
        try { await teamsCallAgent.dispose(); } catch (_) {}
        teamsCallAgent = null;
    }
    teamsCallClient = null;

    _cleanupTeamsState();
    setTeamsStatus('');
    addMessage('system', '[Teams] Left meeting', true);
    document.getElementById('teamsJoinBtn').disabled = false;

    // Delete the ephemeral ACS identity so it doesn't accumulate in the resource
    if (_teamsUserId) {
        const idToDelete = _teamsUserId;
        _teamsUserId = null;
        fetch('/api/acs-delete-user', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ userId: idToDelete }),
        }).catch(e => console.warn('[Teams] ACS identity delete error:', e));
    }
}

function _onTeamsDisconnected() {
    teamsCall = null;
    document.getElementById('teamsJoinBtn').style.display = 'inline-flex';
    document.getElementById('teamsLeaveBtn').style.display = 'none';
    document.getElementById('teamsJoinBtn').disabled = false;
    addMessage('system', '[Teams] Disconnected', true);

    if (teamsCallAgent) {
        teamsCallAgent.dispose().catch(() => {});
        teamsCallAgent = null;
    }

    _cleanupTeamsState();

    // Delete the ephemeral ACS identity so it doesn't accumulate in the resource
    if (_teamsUserId) {
        const idToDelete = _teamsUserId;
        _teamsUserId = null;
        fetch('/api/acs-delete-user', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ userId: idToDelete }),
        }).catch(e => console.warn('[Teams] ACS identity delete error:', e));
    }
}

// ===== Teams → Voice Live Audio Capture =====

async function startTeamsAudioCapture() {
    if (!teamsCall || teamsCall.state !== 'Connected') {
        addMessage('system', '[Teams→AI] Pending — will start when Teams connects', true);
        return;
    }

    // Voice Live session doesn't need to be open yet: the worklet onmessage
    // checks isConnected before sending, so chunks are held until session opens.

    // Stop any previous capture first
    stopTeamsAudioCapture();

    try {
        addMessage('system', '[Teams→AI] Getting remote audio stream...', true);

        // Get the mixed incoming audio MediaStream from ACS.
        // Register a permanent remoteAudioStreamsUpdated listener so that if the
        // stream is removed and replaced later (e.g. participant rejoins) the
        // capture is automatically re-attached.  The handler is cleaned up in
        // stopTeamsAudioCapture() via teamsCall.off().
        if (_remoteAudioStreamsHandler) {
            teamsCall.off('remoteAudioStreamsUpdated', _remoteAudioStreamsHandler);
        }
        _remoteAudioStreamsHandler = async ({ added }) => {
            if (added.length > 0) {
                addMessage('system', '[Teams→AI] Remote audio stream replaced — re-attaching...', true);
                stopTeamsAudioCapture(/* keepListener= */ true);
                await _attachTeamsIncomingStream(added[0]);
            }
        };
        teamsCall.on('remoteAudioStreamsUpdated', _remoteAudioStreamsHandler);

        const remoteStreams = teamsCall.remoteAudioStreams;
        if (!remoteStreams || remoteStreams.length === 0) {
            addMessage('system', '[Teams→AI] No remote streams yet — will attach when remoteAudioStreamsUpdated fires...', true);
            return;
        }
        await _attachTeamsIncomingStream(remoteStreams[0]);
    } catch (e) {
        console.error('[Teams→AI] startTeamsAudioCapture error:', e);
        addMessage('system', `[Teams→AI] Error starting audio capture: ${e.message}`, true);
    }
}

async function _attachTeamsIncomingStream(remoteAudioStream) {
    addMessage('system', '[Teams→AI] Attaching remote audio stream to VoiceLive pipeline...', true);

    const mediaStream = await remoteAudioStream.getMediaStream();
    if (!mediaStream) throw new Error('getMediaStream() returned null');

    const tracks = mediaStream.getAudioTracks();
    addMessage('system', `[Teams→AI] MediaStream obtained — ${tracks.length} audio track(s)`, true);

    // Request echo cancellation on the incoming track where the browser supports it.
    // This prevents the AI's own outgoing voice (heard by remote participants and
    // looped back through Teams) from feeding back into the Voice Live input.
    if (tracks.length > 0) {
        try {
            await tracks[0].applyConstraints({ echoCancellation: true, noiseSuppression: true });
            addMessage('system', '[Teams→AI] Echo cancellation + noise suppression applied to remote track', true);
        } catch (e) {
            addMessage('system', `[Teams→AI] applyConstraints skipped: ${e.message}`, true);
        }
    }

    // Build an AudioContext → AudioWorklet pipeline identical to startAudioCapture()
    teamsInAudioCtx = new AudioContext({ sampleRate: 24000 });

    const source = teamsInAudioCtx.createMediaStreamSource(mediaStream);
    teamsInWorkletNode = await createPCM16WorkletNode(teamsInAudioCtx, 'pcm16-teams-in-processor');

    let chunkCount = 0;
    teamsInWorkletNode.port.onmessage = (e) => {
        const base64 = arrayBufferToBase64(e.data);
        // If VoiceLive session is not yet open, buffer chunks and drain once connected
        if (!isConnected || !ws || ws.readyState !== WebSocket.OPEN) {
            teamsAudioChunkQueue.push(base64);
            // Cap queue at ~5 seconds of audio (5000ms / 20ms per chunk ≈ 250 chunks)
            if (teamsAudioChunkQueue.length > 250) teamsAudioChunkQueue.shift();
            return;
        }
        // Drain any buffered chunks first
        if (teamsAudioChunkQueue.length > 0) {
            for (const buffered of teamsAudioChunkQueue) {
                ws.send(JSON.stringify({ type: 'audio_chunk', data: buffered }));
            }
            teamsAudioChunkQueue = [];
        }
        chunkCount++;
        if (chunkCount === 1) {
            console.log('[Teams→AI] First audio chunk forwarded to Voice Live');
        }
        ws.send(JSON.stringify({ type: 'audio_chunk', data: base64 }));
    };

    source.connect(teamsInWorkletNode);
    // Do NOT connect to destination — we don't want double playback

    // Subscribe to track availability changes so we restart capture if the
    // remote audio stream is interrupted and comes back (e.g. network glitch).
    try {
        remoteAudioStream.on('isAvailableChanged', async () => {
            if (remoteAudioStream.isAvailable) {
                addMessage('system', '[Teams→AI] Remote audio stream became available — restarting capture', true);
                await _attachTeamsIncomingStream(remoteAudioStream);
            } else {
                addMessage('system', '[Teams→AI] Remote audio stream unavailable', true);
                stopTeamsAudioCapture();
            }
        });
    } catch (_) {}

    addMessage('system', '[Teams→AI] ✓ Streaming Teams audio to Voice Live', true);
    setTeamsStatus('Teams: Connected ✓ + sending audio to AI');
}

function stopTeamsAudioCapture(keepListener = false) {
    const wasRunning = !!(teamsInWorkletNode || teamsInAudioCtx);
    if (teamsInWorkletNode) {
        try { teamsInWorkletNode.disconnect(); } catch (_) {}
        teamsInWorkletNode = null;
    }
    if (teamsInAudioCtx) {
        try { teamsInAudioCtx.close(); } catch (_) {}
        teamsInAudioCtx = null;
    }
    teamsAudioChunkQueue = [];
    // Remove the permanent remoteAudioStreamsUpdated listener unless the caller
    // is re-attaching (keepListener=true), in which case the listener stays
    // registered so stream replacement continues to be handled.
    if (!keepListener && _remoteAudioStreamsHandler && teamsCall) {
        try { teamsCall.off('remoteAudioStreamsUpdated', _remoteAudioStreamsHandler); } catch (_) {}
        _remoteAudioStreamsHandler = null;
    }
    if (wasRunning) {
        addMessage('system', '[Teams→AI] Stopped streaming Teams audio to Voice Live', true);
    }
}
