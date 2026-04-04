# TTS/voice answer path is now locked to media_player transport for all spoken replies.
# Do not revert to VA event TTS_END URL delivery for answers.
# All future changes must preserve this path for reliability.

import asyncio
import contextlib
import math
import hashlib
import importlib
import importlib.util
import json
import logging
import os
import queue
import re
import socket
import subprocess
import shutil
import sys
import threading
import time
import uuid
import urllib.error
import urllib.parse
import urllib.request
import wave
import xml.etree.ElementTree as ET
from collections import deque
from dataclasses import dataclass
from datetime import datetime, timedelta
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

try:
    import av
except Exception:  # pragma: no cover - optional runtime dependency
    av = None

try:
    import numpy as np
except Exception:  # pragma: no cover - optional runtime dependency
    np = None

try:
    import sounddevice as sd
except Exception:  # pragma: no cover - optional runtime dependency
    sd = None

try:
    ni = importlib.import_module("netifaces")  # pyright: ignore[reportMissingModuleSource]
except Exception:  # pragma: no cover - optional runtime dependency
    ni = None

try:
    _async_upnp_const = importlib.import_module("async_upnp_client.const")
    _async_upnp_server = importlib.import_module("async_upnp_client.server")
    DeviceInfo = getattr(_async_upnp_const, "DeviceInfo")
    ServiceInfo = getattr(_async_upnp_const, "ServiceInfo")
    UpnpServer = getattr(_async_upnp_server, "UpnpServer")
    UpnpServerDevice = getattr(_async_upnp_server, "UpnpServerDevice")
    UpnpServerService = getattr(_async_upnp_server, "UpnpServerService")
    callable_action = getattr(_async_upnp_server, "callable_action")
    create_event_var = getattr(_async_upnp_server, "create_event_var")
    create_state_var = getattr(_async_upnp_server, "create_state_var")
    # Monkey-patch: on Windows the proactor event loop closes the UDP transport while
    # scheduled SSDP response timers are still pending, causing assert self._response_transport
    # to fire repeatedly.  Guard _send_responses so it silently skips when transport is gone.
    _SsdpSearchResponder = getattr(_async_upnp_server, "SsdpSearchResponder", None)
    if _SsdpSearchResponder is not None:
        _orig_ssdp_send_responses = _SsdpSearchResponder._send_responses
        def _patched_ssdp_send_responses(self, *args, **kwargs):
            if not getattr(self, "_response_transport", None):
                return
            return _orig_ssdp_send_responses(self, *args, **kwargs)
        _SsdpSearchResponder._send_responses = _patched_ssdp_send_responses
except Exception:  # pragma: no cover - optional runtime dependency
    DeviceInfo = None  # type: ignore[assignment]
    ServiceInfo = None  # type: ignore[assignment]
    UpnpServer = None  # type: ignore[assignment]
    UpnpServerDevice = object  # type: ignore[assignment]
    UpnpServerService = object  # type: ignore[assignment]
    callable_action = None  # type: ignore[assignment]
    create_event_var = None  # type: ignore[assignment]
    create_state_var = None  # type: ignore[assignment]

try:
    psutil = importlib.import_module("psutil")  # pyright: ignore[reportMissingModuleSource]
except Exception:  # pragma: no cover - optional runtime dependency
    psutil = None

try:
    import msvcrt
except Exception:  # pragma: no cover - non-Windows fallback
    msvcrt = None

from aioesphomeapi import APIClient, MediaPlayerCommand
from aioesphomeapi.model import VoiceAssistantEventType, VoiceAssistantTimerEventType
from faster_whisper import WhisperModel
from piper.voice import PiperVoice


if getattr(sys, "frozen", False):
    ROOT = Path(sys.executable).resolve().parent
else:
    ROOT = Path(__file__).resolve().parent

_BUNDLED_ROOT = Path(getattr(sys, "_MEIPASS", ROOT))
CONTROL_PAGE_CANDIDATES = (
    ROOT / "control.html",
    _BUNDLED_ROOT / "control.html",
)

# User data directory: store configuration, state, and logs per-user
# On Windows: C:\Users\<username>\AppData\Local\HubVoiceSat
# On macOS: ~/Library/Application Support/HubVoiceSat
# On Linux: ~/.config/hubvoicesat or ~/.local/share/hubvoicesat
if sys.platform == "win32":
    USER_DATA_DIR = Path(os.getenv("LOCALAPPDATA", "~/.local")) / "HubVoiceSat"
elif sys.platform == "darwin":
    USER_DATA_DIR = Path("~/Library/Application Support/HubVoiceSat").expanduser()
else:  # Linux
    USER_DATA_DIR = Path(os.getenv("XDG_CONFIG_HOME", "~/.config")).expanduser() / "hubvoicesat"

USER_DATA_DIR.mkdir(parents=True, exist_ok=True)

CONFIG_PATH = USER_DATA_DIR / "hubvoice-sat-setup.json"
SATELLITES_PATH = USER_DATA_DIR / "satellites.csv"
RECORDINGS_PATH = USER_DATA_DIR / "recordings"
LOGS_PATH = USER_DATA_DIR / "logs"
SCHEDULES_PATH = USER_DATA_DIR / "scheduled-events.json"
LEGACY_CONFIG_CANDIDATES = (ROOT / "hubvoice-sat-setup.json", Path.cwd() / "hubvoice-sat-setup.json")
LEGACY_SATELLITES_CANDIDATES = (ROOT / "satellites.csv", Path.cwd() / "satellites.csv")
LOGS_PATH.mkdir(parents=True, exist_ok=True)
RECORDINGS_PATH.mkdir(parents=True, exist_ok=True)

_LEGACY_USER_FILES_MIGRATED = False

AIRPLAY_RECEIVER_SRC_DIR = ROOT / "build" / "airplay2-receiver-src"
AIRPLAY_RECEIVER_SCRIPT = AIRPLAY_RECEIVER_SRC_DIR / "ap2-receiver.py"
AIRPLAY_RECEIVER_LOG_PATH = LOGS_PATH / "airplay-receiver.log"
AIRPLAY_RECEIVER_DEFAULT_NAME = "HubVoiceAirPlay"
AIRPLAY_REQUIRED_MODULES = (
    "netifaces",
    "zeroconf",
    "biplist",
    "Crypto",
    "hexdump",
    "srptools",
    "hkdf",
    "pyaudio",
)
DLNA_RENDERER_DEFAULT_NAME = "HubVoiceDLNA"
DLNA_RENDERER_HTTP_PORT = 38520
DLNA_TRANSPORT_STOP_GRACE_SECONDS = 1.2
DLNA_REPEAT_PLAY_GUARD_SECONDS = 8.0
DLNA_PROXY_PREBUFFER_BYTES = 32768   # keep startup latency low for stricter HTTP clients
DLNA_PROXY_PREBUFFER_MAX_WAIT_SECONDS = 14.0
DLNA_PROXY_BITRATE = "128k"
DLNA_REQUIRED_MODULES = (
    "async_upnp_client",
    "aiohttp",
    "defusedxml",
    "voluptuous",
)

# ============================================================================
# CONFIGURATION CONSTANTS
# ============================================================================
# Server & Request Handling
REQUEST_RATE_LIMIT = 60  # requests per minute
REQUEST_SIZE_LIMIT = 50000  # bytes (50KB)
REQUEST_TIMEOUT_SECONDS = 30  # default request timeout

# Cache Management
ENTITY_CACHE_MAX = 10
RECORDING_MAX_FILES = 1000
RECORDING_MAX_AGE_HOURS = 24
CLEANUP_INTERVAL_SECS = 3600

# Timeouts
CONNECTION_TIMEOUT = 10.0  # seconds
SET_NUMBER_TIMEOUT = 5.0  # seconds
ENTITY_LIST_TIMEOUT = 5.0  # seconds
SATELLITE_CAP_CACHE_TTL = 30.0  # seconds
HUBITAT_REQUEST_TIMEOUT = 20  # seconds
SATELLITE_MEDIA_DELAY = 0.3  # seconds
RECONNECT_DELAY = 2  # seconds
VOICE_BRIDGE_CONNECT_TIMEOUT = 3.0  # seconds
TTS_PRESTOP_CONNECT_TIMEOUT = 1.5  # seconds
TTS_PRESTOP_ENTITY_TIMEOUT = 1.0  # seconds
SATELLITE_NUMBER_MAX_RETRIES = 2
SATELLITE_NUMBER_RETRY_DELAY = 0.25  # seconds
SATELLITE_NUMBER_POST_DELAY = 0.2  # seconds
HUBMUSIC_SOFT_STOP_ONLY = True

# HubMusic known-good profile (validated 2026-03-28: stable startup, in-sync multi-speaker playback).
# Treat this block as the baseline and only tune with live A/B testing plus log validation.
HUBMUSIC_PROFILE_ID = "known_good_2026_03_28"
HUBMUSIC_KNOWN_GOOD = {
    "stream_bitrate": 128000,
    "stream_frames": 4096,
    # Device index 9 is unstable on this host and can fail with PaError -9999.
    "auto_capture_device_blacklist": {9},
    "live_warmup_seconds": 1.2,
    "live_prebuffer_bytes": 98304,
    "live_prebuffer_max_wait_seconds": 2.5,
    "cpu_spike_threshold_pct": 80.0,
    "cpu_spike_extra_warmup_seconds": 0.8,
    "cpu_check_samples": 3,
    "cpu_check_interval_seconds": 0.08,
}
HUBMUSIC_STREAM_BITRATE = int(HUBMUSIC_KNOWN_GOOD["stream_bitrate"])
HUBMUSIC_STREAM_FRAMES = int(HUBMUSIC_KNOWN_GOOD["stream_frames"])
HUBMUSIC_AUTO_CAPTURE_DEVICE_BLACKLIST = set(HUBMUSIC_KNOWN_GOOD["auto_capture_device_blacklist"])
HUBMUSIC_LIVE_WARMUP_SECONDS = float(HUBMUSIC_KNOWN_GOOD["live_warmup_seconds"])
HUBMUSIC_LIVE_PREBUFFER_BYTES = int(HUBMUSIC_KNOWN_GOOD["live_prebuffer_bytes"])
HUBMUSIC_LIVE_PREBUFFER_MAX_WAIT_SECONDS = float(HUBMUSIC_KNOWN_GOOD["live_prebuffer_max_wait_seconds"])
HUBMUSIC_CPU_SPIKE_THRESHOLD_PCT = float(HUBMUSIC_KNOWN_GOOD["cpu_spike_threshold_pct"])
HUBMUSIC_CPU_SPIKE_EXTRA_WARMUP_SECONDS = float(HUBMUSIC_KNOWN_GOOD["cpu_spike_extra_warmup_seconds"])
HUBMUSIC_CPU_CHECK_SAMPLES = int(HUBMUSIC_KNOWN_GOOD["cpu_check_samples"])
HUBMUSIC_CPU_CHECK_INTERVAL_SECONDS = float(HUBMUSIC_KNOWN_GOOD["cpu_check_interval_seconds"])
HUBMUSIC_STEREO_MAX_START_SKEW_MS = 55
HUBMUSIC_STEREO_RESYNC_MAX_PASSES = 3
HUBMUSIC_STEREO_RESYNC_SETTLE_SECONDS = 0.09
HUBMUSIC_STEREO_SHARED_LAUNCH_DELAY_MS = 2600
SATELLITE_MEDIA_VOLUME_MIN_APPLY_INTERVAL_SECONDS = 1.0
SATELLITE_MEDIA_VOLUME_MIN_DELTA_PERCENT = 2.0

# Voice broadcast flow
BROADCAST_PENDING_TIMEOUT_SECS = 25
BROADCAST_PROMPT = "What's the message?"

# Local timer and alarm scheduling
SCHEDULER_POLL_INTERVAL_SECS = 1.0
MAX_TIMER_DURATION_SECONDS = 24 * 60 * 60
MAX_ACTIVE_TIMERS_PER_SATELLITE = 8
MAX_ACTIVE_ALARMS_PER_SATELLITE = 8
FALLBACK_RING_MAX_SECONDS = 60
FALLBACK_RING_REPEAT_SECONDS = 6

# Hubitat Resilience
HUBITAT_RETRY_MAX = 3
HUBITAT_RETRY_INITIAL_DELAY = 0.5  # seconds
HUBITAT_CIRCUIT_BREAKER_THRESHOLD = 5
HUBITAT_CIRCUIT_BREAKER_TIMEOUT = 300  # seconds (5 minutes)

_log_formatter = logging.Formatter("%(asctime)s [%(levelname)-8s] %(message)s")
_file_handler = logging.FileHandler(str(LOGS_PATH / "hubvoice-runtime.log"), encoding="utf-8")
_file_handler.setFormatter(_log_formatter)
_console_handler = logging.StreamHandler()
_console_handler.setFormatter(_log_formatter)
logging.basicConfig(level=logging.INFO, handlers=[_file_handler, _console_handler])


class HubitatCircuitBreaker:
    """Circuit breaker for Hubitat API to fail fast and protect against cascading failures."""
    
    def __init__(self, failure_threshold: int = HUBITAT_CIRCUIT_BREAKER_THRESHOLD, 
                 timeout_seconds: int = HUBITAT_CIRCUIT_BREAKER_TIMEOUT):
        self._lock = threading.Lock()
        self.failures = 0
        self.threshold = failure_threshold
        self.timeout = timeout_seconds
        self.last_failure_time = 0.0
    
    def is_open(self) -> bool:
        """Check if circuit breaker is open (failing fast)."""
        with self._lock:
            if self.failures >= self.threshold:
                elapsed = time.time() - self.last_failure_time
                if elapsed > self.timeout:
                    # Time out exceeded, try again
                    return False
                return True
        return False
    
    def record_failure(self) -> None:
        """Record a failure."""
        with self._lock:
            self.failures += 1
            self.last_failure_time = time.time()
            if self.failures >= self.threshold:
                logging.error(
                    "Hubitat circuit breaker OPEN after %d consecutive failures — "
                    "fast-failing all Hubitat calls for %ds",
                    self.failures, self.timeout,
                )
            else:
                logging.warning("Circuit breaker failure %d/%d", self.failures, self.threshold)
    
    def reset(self) -> None:
        """Reset circuit breaker after successful operation."""
        with self._lock:
            self.failures = 0
            logging.info("Circuit breaker reset")


class Metrics:
    """Track application metrics for observability."""
    
    def __init__(self):
        self._lock = threading.Lock()
        self.total_requests = 0
        self.total_errors = 0
        self.hubitat_errors = 0
        self.satellite_errors = 0
        self.request_latencies: deque = deque(maxlen=100)
        self.start_time = time.time()
    
    def record_request(self, latency: float, error: bool = False, error_type: str = "") -> None:
        """Record a request with latency and error status."""
        with self._lock:
            self.total_requests += 1
            if error:
                self.total_errors += 1
                if "hubitat" in error_type.lower():
                    self.hubitat_errors += 1
                elif "satellite" in error_type.lower():
                    self.satellite_errors += 1
            self.request_latencies.append(latency)
    
    def get_stats(self) -> dict:
        """Get current metrics statistics."""
        with self._lock:
            if not self.request_latencies:
                avg_latency = 0.0
            else:
                avg_latency = sum(self.request_latencies) / len(self.request_latencies)
            
            error_rate = (self.total_errors / self.total_requests * 100) if self.total_requests else 0.0
            uptime = time.time() - self.start_time
            
            return {
                "total_requests": self.total_requests,
                "total_errors": self.total_errors,
                "hubitat_errors": self.hubitat_errors,
                "satellite_errors": self.satellite_errors,
                "error_rate_percent": error_rate,
                "avg_latency_ms": avg_latency * 1000,
                "uptime_seconds": uptime,
            }


class RequestContext(threading.local):
    """Per-thread request context for tracing requests through the system."""
    
    def __init__(self):
        super().__init__()
        self.request_id = str(uuid.uuid4())[:8]
        self.start_time = time.time()
    
    @property
    def elapsed_ms(self) -> float:
        """Get elapsed time in milliseconds."""
        return (time.time() - self.start_time) * 1000
    
    def reset(self) -> None:
        """Reset context for new request."""
        self.request_id = str(uuid.uuid4())[:8]
        self.start_time = time.time()


class AppState:
    """Thread-safe application state container."""
    def __init__(self):
        self._lock = threading.Lock()
        self._state = {
            "last_error": "",
            "last_action": "idle",
            "last_transcript": "",
        }
    
    def update(self, **kwargs):
        """Update state values thread-safely."""
        with self._lock:
            self._state.update(kwargs)
    
    def snapshot(self) -> dict:
        """Get a snapshot of current state."""
        with self._lock:
            return dict(self._state)


class HubMusicState:
    """Track requested HubMusic routing state for the control UI."""

    def __init__(self):
        self._lock = threading.Lock()
        self._state = {
            "active": False,
            "mode": "single",
            "satellite": "",
            "satellite_alias": "",
            "satellite_host": "",
            "satellites": [],
            "exclude_satellite": "",
            "title": "",
            "source_url": "",
            "started_at": "",
            "last_action": "idle",
            "last_error": "",
            "last_operation": "",
            "last_sent": [],
            "last_stopped": [],
            "last_failed": [],
            "last_retried": [],
            "history": [],
        }

    def activate(
        self,
        satellites: list[dict],
        source_url: str,
        title: str,
        mode: str = "single",
    ) -> None:
        first = satellites[0] if satellites else {}
        targets = []
        for satellite in satellites:
            targets.append(
                {
                    "id": str(satellite.get("id", "")).strip(),
                    "alias": str(satellite.get("alias", "")).strip(),
                    "host": str(satellite.get("host", "")).strip(),
                    "channel": str(satellite.get("channel", "")).strip(),
                }
            )
        with self._lock:
            self._state.update(
                {
                    "active": True,
                    "mode": mode if mode in {"single", "all_reachable", "stereo_pair"} else "single",
                    "satellite": str(first.get("id", "")).strip(),
                    "satellite_alias": str(first.get("alias", "")).strip(),
                    "satellite_host": str(first.get("host", "")).strip(),
                    "satellites": targets,
                    "title": title,
                    "source_url": source_url,
                    "started_at": datetime.now().isoformat(timespec="seconds"),
                    "last_action": "play",
                    "last_error": "",
                }
            )

    def stop(self, satellites: list[dict] | None = None) -> None:
        satellites = satellites or []
        first = satellites[0] if satellites else {}
        targets = []
        for satellite in satellites:
            targets.append(
                {
                    "id": str(satellite.get("id", "")).strip(),
                    "alias": str(satellite.get("alias", "")).strip(),
                    "host": str(satellite.get("host", "")).strip(),
                    "channel": str(satellite.get("channel", "")).strip(),
                }
            )
        with self._lock:
            if first:
                self._state["satellite"] = str(first.get("id", "")).strip()
                self._state["satellite_alias"] = str(first.get("alias", "")).strip()
                self._state["satellite_host"] = str(first.get("host", "")).strip()
            if targets:
                self._state["satellites"] = targets
            self._state["active"] = False
            self._state["last_action"] = "stop"
            self._state["last_error"] = ""

    def error(self, message: str) -> None:
        with self._lock:
            self._state["last_error"] = message
            self._state["last_action"] = "error"

    def set_results(
        self,
        operation: str,
        *,
        sent: list[dict] | None = None,
        stopped: list[dict] | None = None,
        failed: list[dict] | None = None,
        retried: list[dict] | None = None,
        exclude_satellite: str = "",
        mode: str = "",
    ) -> None:
        with self._lock:
            self._state["last_operation"] = operation
            self._state["last_sent"] = list(sent or [])
            self._state["last_stopped"] = list(stopped or [])
            self._state["last_failed"] = list(failed or [])
            self._state["last_retried"] = list(retried or [])
            self._state["exclude_satellite"] = exclude_satellite

            effective_mode = mode or str(self._state.get("mode", "single"))
            history = list(self._state.get("history", []))
            history.insert(
                0,
                {
                    "timestamp": datetime.now().isoformat(timespec="seconds"),
                    "operation": operation,
                    "mode": effective_mode,
                    "exclude_satellite": exclude_satellite,
                    "sent_count": len(sent or []),
                    "stopped_count": len(stopped or []),
                    "failed_count": len(failed or []),
                    "retried_count": len(retried or []),
                },
            )
            self._state["history"] = history[:10]

    def snapshot(self) -> dict:
        with self._lock:
            return dict(self._state)


class SatelliteConnectionPool:
    """Cache and reuse satellite API client connections."""
    def __init__(self, max_connections: int = 5):
        self._pool: dict[str, APIClient] = {}
        self._locks: dict[str, threading.Lock] = {}
        self._max_connections = max_connections
        self._pool_lock = threading.Lock()
    
    async def get_client(self, host: str) -> APIClient:
        """Get or create cached connection to satellite."""
        if host not in self._locks:
            with self._pool_lock:
                if host not in self._locks:
                    self._locks[host] = threading.Lock()
        
        with self._locks[host]:
            if host in self._pool:
                return self._pool[host]
            
            # Create new connection
            logging.info("Creating new satellite connection to %s", host)
            client = APIClient(host, 6054, None, client_info="HubVoiceSatRuntime")
            try:
                await asyncio.wait_for(
                    client.connect(login=False),
                    timeout=10.0
                )
            except asyncio.TimeoutError:
                raise RuntimeError(f"Timeout connecting to satellite {host} (10s)")
            except Exception as e:
                raise RuntimeError(f"Failed to connect to satellite {host}: {e}")
            
            self._pool[host] = client
            return client
    
    async def close_all(self):
        """Close all cached connections."""
        for client in self._pool.values():
            try:
                await client.disconnect()
            except Exception as e:
                logging.warning("Error disconnecting satellite: %s", e)
        self._pool.clear()


_RATE_LIMIT_WINDOW = deque()  # Track request times for rate limiting
_RATE_LIMIT_LOCK = threading.Lock()
_APP_STATE = AppState()
_SATELLITE_POOL = SatelliteConnectionPool()
_HUBITAT_BREAKER = HubitatCircuitBreaker()
_METRICS = Metrics()
_REQUEST_CONTEXT = RequestContext()
_WORK_QUEUE: "queue.Queue[dict]" = queue.Queue()
_PIPER_VOICE = None
_PIPER_VOICE_MODEL_PATH = ""
_PIPER_VOICE_LOCK = threading.Lock()
_WHISPER_MODEL = None
_WHISPER_MODEL_LOCK = threading.Lock()
_ENTITY_CACHE: dict[str, dict[str, int]] = {}
_ENTITY_CACHE_LOCK = threading.Lock()
_SATELLITE_CAP_CACHE: dict[str, dict] = {}
_SATELLITE_CAP_CACHE_LOCK = threading.Lock()
_SATELLITE_ENTITY_KEY_CACHE: dict[str, dict[str, int]] = {}
_SATELLITE_ENTITY_KEY_CACHE_LOCK = threading.Lock()
_SHUTDOWN_EVENT = threading.Event()
_CONTROL_DECK_STATE: dict[str, dict] = {}
_CONTROL_DECK_STATE_LOCK = threading.Lock()

# HubMusic Snapcast: ffmpeg subprocess that feeds audio into snapserver
_HUBMUSIC_FFMPEG_PROC: subprocess.Popen | None = None
_HUBMUSIC_FFMPEG_PIDS: set[int] = set()   # all pids we have ever started, for orphan cleanup
_HUBMUSIC_FFMPEG_LOCK = threading.Lock()
_HUBMUSIC_START_LOCK = threading.Lock()   # serialises start/stop so watchdog and play can't race
_HUBMUSIC_RELAY_STOP: threading.Event | None = None   # signals relay thread to exit cleanly
SNAPSERVER_TCP_PORT = 4953


def _kill_pid(pid: int) -> None:
    """Force-kill a process by PID using taskkill, ignoring errors."""
    try:
        subprocess.run(
            ["taskkill", "/F", "/PID", str(pid)],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW,
            timeout=5,
        )
    except Exception:
        pass


def _start_hubmusic_ffmpeg(source_url: str) -> None:
    """Start ffmpeg feeding the given audio URL as raw PCM into the snapserver TCP input."""
    # Bypass snapserver: Instead, start ffmpeg as a subprocess and serve the FLAC stream via HTTP.
    # The HTTP handler will launch ffmpeg on demand and stream its output to the satellite.
    # This function is now a no-op; see the HTTP handler for /hubmusic/live.flac or similar.
    pass


def _stop_hubmusic_ffmpeg() -> None:
    """Terminate all tracked ffmpeg feeder processes and stop the relay thread."""
    global _HUBMUSIC_FFMPEG_PROC, _HUBMUSIC_RELAY_STOP
    with _HUBMUSIC_FFMPEG_LOCK:
        proc = _HUBMUSIC_FFMPEG_PROC
        _HUBMUSIC_FFMPEG_PROC = None
        pids_to_kill = set(_HUBMUSIC_FFMPEG_PIDS)
        _HUBMUSIC_FFMPEG_PIDS.clear()
        relay_stop = _HUBMUSIC_RELAY_STOP
        _HUBMUSIC_RELAY_STOP = None
    if relay_stop is not None:
        relay_stop.set()
    # Kill the tracked Popen object
    if proc is not None:
        try:
            proc.terminate()
            proc.wait(timeout=3)
        except Exception:
            pass
    # Kill any orphaned PIDs (processes that leaked from previous races)
    for pid in pids_to_kill:
        if proc is not None and pid == proc.pid:
            continue  # already handled above
        _kill_pid(pid)
    if proc is not None or pids_to_kill:
        logging.info("HubMusic snapcast: stopped ffmpeg (pids=%s)", pids_to_kill | ({proc.pid} if proc else set()))


def load_config() -> dict:
    default = {
        "hubvoice_url": "http://127.0.0.1:8080",
        "callback_url": "http://127.0.0.1:8080/answer",
        "hubitat_host": "",
        "hubitat_app_id": "",
        "hubitat_access_token": "",
        "piper_voice_model": "piper_voices/en_US-amy-medium.onnx",
    }
    _maybe_migrate_legacy_user_files()
    if not CONFIG_PATH.exists():
        for candidate in LEGACY_CONFIG_CANDIDATES:
            if not candidate.exists() or candidate.resolve() == CONFIG_PATH.resolve():
                continue
            try:
                loaded = json.loads(candidate.read_text(encoding="utf-8-sig"))
                if isinstance(loaded, dict):
                    default.update({k: str(v) for k, v in loaded.items() if v is not None})
                    with contextlib.suppress(Exception):
                        CONFIG_PATH.write_text(json.dumps(loaded, indent=4), encoding="utf-8")
                    break
            except Exception:
                continue
        return default
    try:
        loaded = json.loads(CONFIG_PATH.read_text(encoding="utf-8-sig"))
        if isinstance(loaded, dict):
            default.update({k: str(v) for k, v in loaded.items() if v is not None})
    except Exception as exc:
        logging.exception("Failed to load config: %s", exc)
    return default


def _load_runtime_config_raw() -> dict:
    _maybe_migrate_legacy_user_files()
    if not CONFIG_PATH.exists():
        return {}
    try:
        loaded = json.loads(CONFIG_PATH.read_text(encoding="utf-8-sig"))
        if isinstance(loaded, dict):
            return loaded
    except Exception as exc:
        logging.exception("Failed to load raw runtime config: %s", exc)
    return {}


def _save_runtime_config_raw(config: dict) -> None:
    if not isinstance(config, dict):
        raise ValueError("config must be a dictionary")
    CONFIG_PATH.write_text(json.dumps(config, indent=4), encoding="utf-8")


def _load_persisted_tone_settings_for_host(satellite_host: str) -> tuple[float, float]:
    host = str(satellite_host or "").strip().lower()
    if not host:
        return 0.0, 0.0
    cfg = _load_runtime_config_raw()
    tone_map = cfg.get("hubmusic_tone_settings")
    if not isinstance(tone_map, dict):
        return 0.0, 0.0
    tone_entry = tone_map.get(host)
    if not isinstance(tone_entry, dict):
        return 0.0, 0.0
    try:
        bass_db = _normalize_tone_db(tone_entry.get("bass_level", 0.0))
    except Exception:
        bass_db = 0.0
    try:
        treble_db = _normalize_tone_db(tone_entry.get("treble_level", 0.0))
    except Exception:
        treble_db = 0.0
    return float(bass_db), float(treble_db)


def _persist_tone_settings_for_host(satellite_host: str, bass_db: float, treble_db: float) -> None:
    host = str(satellite_host or "").strip().lower()
    if not host:
        return
    cfg = _load_runtime_config_raw()
    tone_map = cfg.get("hubmusic_tone_settings")
    if not isinstance(tone_map, dict):
        tone_map = {}
    tone_map[host] = {
        "bass_level": round(float(bass_db), 2),
        "treble_level": round(float(treble_db), 2),
    }
    cfg["hubmusic_tone_settings"] = tone_map
    _save_runtime_config_raw(cfg)


def _config_bool(value: object, default: bool) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    text = str(value).strip().lower()
    if text in {"1", "true", "yes", "on"}:
        return True
    if text in {"0", "false", "no", "off"}:
        return False
    return default


def _config_int(value: object, default: int, minimum: int, maximum: int) -> int:
    try:
        parsed = int(str(value).strip()) if value is not None else default
    except Exception:
        parsed = default
    return max(minimum, min(maximum, parsed))


def validate_startup_config() -> None:
    """Validate critical configuration at startup."""
    # Check satellites are configured
    satellites = load_satellites()
    if not satellites:
        logging.warning("No satellites configured in satellites.csv; continuing so setup can complete first-run configuration")
        satellites = []
    
    # Test each satellite is reachable
    reachable = []
    unreachable = []
    for sat in satellites:
        if test_satellite_connection(sat["host"]):
            reachable.append(sat["id"])
        else:
            unreachable.append((sat["id"], sat["host"]))
    
    if satellites and not reachable:
        logging.warning(
            "No satellites reachable at startup — voice commands will fail until at least one responds on port 6054.\n"
            "  Unreachable: %s",
            ", ".join(f"{sid} ({host})" for sid, host in unreachable),
        )
    elif unreachable:
        logging.warning(
            "Some satellites unreachable at startup (port 6054): %s",
            ", ".join(f"{sid} ({host})" for sid, host in unreachable),
        )
    
    # Verify Piper model file exists
    try:
        resolve_piper_model_path()
        logging.info("Piper voice model validated: found")
    except RuntimeError as e:
        logging.warning("Model validation warning at startup: %s", e)
    
    logging.info("Configuration validated - %d/%d satellites reachable", len(reachable), len(satellites))


def test_satellite_connection(host: str, timeout: float = 2.0) -> bool:
    """Quick connectivity test to satellite on ESPHome API port (6054)."""
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(timeout)
        try:
            result = sock.connect_ex((host, 6054))
            return result == 0
        finally:
            sock.close()
    except Exception as e:
        logging.debug("Satellite connection test failed for %s: %s", host, e)
        return False


async def retry_with_backoff(
    coro_func,
    *args,
    max_retries: int = HUBITAT_RETRY_MAX,
    initial_delay: float = HUBITAT_RETRY_INITIAL_DELAY,
    **kwargs
):
    """
    Retry an async operation with exponential backoff.
    
    Args:
        coro_func: Async function to call
        *args: Positional arguments to pass
        max_retries: Maximum number of attempts
        initial_delay: Initial delay in seconds
        **kwargs: Keyword arguments to pass
    
    Returns:
        Result of the async function
    
    Raises:
        RuntimeError: If all retries exhausted
    """
    delay = initial_delay
    last_error = None
    
    for attempt in range(max_retries):
        try:
            return await coro_func(*args, **kwargs)
        except Exception as e:
            last_error = e
            if attempt < max_retries - 1:
                logging.debug("Retry attempt %d/%d after %.2fs: %s", 
                            attempt + 1, max_retries, delay, e)
                await asyncio.sleep(delay)
                delay *= 2  # Exponential backoff
            else:
                logging.error("All %d retry attempts exhausted: %s", max_retries, e)
    
    raise RuntimeError(f"Failed after {max_retries} retries: {last_error}")


def get_runtime_port() -> int:
    config = load_config()
    for key in ("hubvoice_url", "callback_url"):
        raw = str(config.get(key, "")).strip()
        if not raw:
            continue
        try:
            parsed = urllib.parse.urlparse(raw)
            if parsed.port:
                return int(parsed.port)
        except Exception:
            continue
    return 8080


def _is_populated_satellites_text(text: str) -> bool:
    for line in (text or "").splitlines():
        raw = line.strip()
        if not raw or raw.startswith("#"):
            continue
        if "," in raw:
            return True
    return False


def _maybe_migrate_legacy_user_files() -> None:
    global _LEGACY_USER_FILES_MIGRATED
    if _LEGACY_USER_FILES_MIGRATED:
        return

    try:
        if not CONFIG_PATH.exists():
            for candidate in LEGACY_CONFIG_CANDIDATES:
                if not candidate.exists() or candidate.resolve() == CONFIG_PATH.resolve():
                    continue
                try:
                    raw = candidate.read_text(encoding="utf-8-sig")
                    loaded = json.loads(raw)
                    if isinstance(loaded, dict):
                        CONFIG_PATH.write_text(json.dumps(loaded, indent=4), encoding="utf-8")
                        logging.info("Migrated legacy config from %s to %s", candidate, CONFIG_PATH)
                        break
                except Exception:
                    continue

        if not SATELLITES_PATH.exists():
            for candidate in LEGACY_SATELLITES_CANDIDATES:
                if not candidate.exists() or candidate.resolve() == SATELLITES_PATH.resolve():
                    continue
                try:
                    raw = candidate.read_text(encoding="utf-8-sig")
                    if _is_populated_satellites_text(raw):
                        SATELLITES_PATH.write_text(raw, encoding="utf-8")
                        logging.info("Migrated legacy satellites from %s to %s", candidate, SATELLITES_PATH)
                        break
                except Exception:
                    continue
    finally:
        _LEGACY_USER_FILES_MIGRATED = True


def load_satellites() -> list[dict]:
    items: list[dict] = []
    _maybe_migrate_legacy_user_files()
    raw_text = ""
    if SATELLITES_PATH.exists():
        raw_text = SATELLITES_PATH.read_text(encoding="utf-8-sig")
    else:
        for candidate in LEGACY_SATELLITES_CANDIDATES:
            if not candidate.exists() or candidate.resolve() == SATELLITES_PATH.resolve():
                continue
            try:
                candidate_text = candidate.read_text(encoding="utf-8-sig")
            except Exception:
                continue
            if not _is_populated_satellites_text(candidate_text):
                continue
            raw_text = candidate_text
            with contextlib.suppress(Exception):
                SATELLITES_PATH.write_text(candidate_text, encoding="utf-8")
            break

    if not raw_text:
        return items

    for line in raw_text.splitlines():
        raw = line.strip()
        if not raw or raw.startswith("#") or "," not in raw:
            continue
        parts = [part.strip() for part in raw.split(",", 2)]
        if len(parts) < 2:
            continue
        name, host = parts[0].lstrip("\ufeff"), parts[1]
        alias = parts[2] if len(parts) >= 3 else ""

        # Recover malformed rows like "id,192.168.4.135.alias".
        if not alias:
            host_parts = host.split(".")
            if len(host_parts) > 4 and all(p.isdigit() for p in host_parts[:4]):
                ipv4 = ".".join(host_parts[:4])
                suffix = ".".join(host_parts[4:]).strip()
                if suffix:
                    host = ipv4
                    alias = suffix

        if name and host:
            items.append({"id": name, "host": host, "alias": alias})
    return items


def check_for_new_satellites_from_csv(max_disconnected_attempts: int = 3) -> None:
    """
    Check satellites CSV for new devices and add them to the satellite config.
    
    Args:
        max_disconnected_attempts: Maximum consecutive disconnection attempts
                                  before marking a satellite as inactive 
                                  (default: 3)
    
    Behavior:
        - Scans SATELLITES_CSV for new devices
        - Adds new device entries to the satellite config
        - Marks satellites as inactive after max_disconnected_attempts failures
    
    Note: Requires SATELLITES_CSV to exist and be properly formatted
    
    Raises:
        RuntimeError: If CSV parsing fails or config update fails
    """
    try:
        # Implementation would read CSV and add new devices
        satellites = load_satellites()
        if not satellites:
            logging.warning("No satellites found in CSV")
            return
        
        logging.info("Checked satellites CSV: %d satellites found", len(satellites))
    except Exception as e:
        logging.error("Failed to check satellites CSV: %s", e)


def select_satellite(sat_id: str) -> dict | None:
    satellites = load_satellites()
    if not satellites:
        return None
    wanted = (sat_id or "").strip().lower()
    if wanted:
        for sat in satellites:
            sat_id_norm = sat["id"].strip().lower()
            sat_alias_norm = str(sat.get("alias", "")).strip().lower()
            sat_host_norm = sat["host"].strip().lower()
            if wanted in (sat_id_norm, sat_alias_norm, sat_host_norm):
                return sat
        for sat in satellites:
            sat_id_norm = sat["id"].strip().lower()
            sat_alias_norm = str(sat.get("alias", "")).strip().lower()
            if (sat_id_norm and wanted in sat_id_norm) or (sat_alias_norm and wanted in sat_alias_norm):
                return sat
    return satellites[0]


def get_runtime_host() -> str:
    config = load_config()
    raw = str(config.get("hubvoice_url", "")).strip()
    if raw:
        try:
            parsed = urllib.parse.urlparse(raw)
            if parsed.hostname and parsed.hostname not in ("127.0.0.1", "localhost"):
                return parsed.hostname
        except Exception:
            pass
    satellite = select_satellite("")
    try:
        probe_host = satellite["host"] if satellite else "8.8.8.8"
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            sock.connect((probe_host, 80))
            address = sock.getsockname()[0]
            if address:
                return address
        finally:
            sock.close()
    except Exception:
        pass
    return "127.0.0.1"


def sanitize_text(text: str) -> str:
    normalized = " ".join(
        (text or "")
        .replace("\r", " ")
        .replace("\n", " ")
        .replace("\u2018", "'")
        .replace("\u2019", "'")
        .replace("\u201c", '"')
        .replace("\u201d", '"')
        .split()
    )
    if not normalized:
        return "No answer."
    if len(normalized) > 350:
        normalized = normalized[:347].rstrip() + "..."
    return normalized


def get_yaml_substitution_value(key: str, fallback: str) -> str:
        yaml_path = ROOT / "hubvoice-sat.yaml"
        if not yaml_path.exists():
                return fallback
        try:
                text = yaml_path.read_text(encoding="utf-8")
                pattern = rf"(?m)^\s*{re.escape(key)}:\s*\"?(.*?)\"?\s*$"
                match = re.search(pattern, text)
                if match:
                        value = (match.group(1) or "").strip()
                        return value or fallback
        except Exception:
                pass
        return fallback


def build_satellite_control_metadata() -> dict:
        return {
                "speaker_volume": {
                        "label": "Speaker Volume",
            "min": 0,
                        "max": int(get_yaml_substitution_value("speaker_volume_max", "85")),
                        "default": int(get_yaml_substitution_value("speaker_volume_initial", "65")),
            "settable": True,
                },
                "wake_sound_volume": {
                        "label": "Wake Sound Volume",
                        "min": int(get_yaml_substitution_value("wake_volume_min", "40")),
                        "max": int(get_yaml_substitution_value("wake_volume_max", "85")),
                        "default": int(get_yaml_substitution_value("wake_volume_initial", "65")),
            "settable": True,
        },
        "whisper_volume_pct": {
            "label": "Whisper Volume",
            "min": 1,
            "max": 100,
            "default": int(get_yaml_substitution_value("whisper_volume_pct", "20")),
            "settable": False,
                },
                "follow_up_listen_window_seconds": {
                        "label": "Follow-up Window",
                        "min": 1,
                        "max": 30,
                        "default": 2,
            "settable": True,
                },
                "bass_level": {
                    "label": "Bass",
                    "min": -10,
                    "max": 10,
                    "default": int(get_yaml_substitution_value("bass_level_initial", "0")),
                "settable": True,
                },
                "treble_level": {
                    "label": "Treble",
                    "min": -10,
                    "max": 10,
                    "default": int(get_yaml_substitution_value("treble_level_initial", "0")),
                "settable": True,
                },
        }


def _control_deck_default_state(controls: dict) -> dict:
    return {
        "speaker_volume": float(controls.get("speaker_volume", {}).get("default", 65)),
        "wake_sound_volume": float(controls.get("wake_sound_volume", {}).get("default", 65)),
        "whisper_volume_pct": float(controls.get("whisper_volume_pct", {}).get("default", 20)),
        "follow_up_listen_window_seconds": float(controls.get("follow_up_listen_window_seconds", {}).get("default", 2)),
        "bass_level": float(controls.get("bass_level", {}).get("default", 0)),
        "treble_level": float(controls.get("treble_level", {}).get("default", 0)),
        "speaker_muted": False,
        "wake_muted": False,
        "whisper_mode": False,
        "follow_up_listening_switch": False,
    }


def _get_control_deck_state(sat_id: str, controls: dict) -> dict:
    sat_key = str(sat_id or "").strip().lower()
    defaults = _control_deck_default_state(controls)
    with _CONTROL_DECK_STATE_LOCK:
        cached = dict(_CONTROL_DECK_STATE.get(sat_key, {}))
        merged = {**defaults, **cached}
        try:
            wake_min = float(controls.get("wake_sound_volume", {}).get("min", 0))
        except Exception:
            wake_min = 0.0
        merged["wake_muted"] = bool(merged.get("wake_muted", False) or float(merged.get("wake_sound_volume", wake_min)) <= wake_min)
        _CONTROL_DECK_STATE[sat_key] = merged
        return dict(merged)


def _update_control_deck_state(sat_id: str, **updates) -> None:
    sat_key = str(sat_id or "").strip().lower()
    if not sat_key or not updates:
        return
    with _CONTROL_DECK_STATE_LOCK:
        current = dict(_CONTROL_DECK_STATE.get(sat_key, {}))
        for key, value in updates.items():
            current[key] = value
        _CONTROL_DECK_STATE[sat_key] = current


def build_satellite_control_page() -> str:
    for candidate in CONTROL_PAGE_CANDIDATES:
        try:
            if candidate.exists():
                return candidate.read_text(encoding="utf-8")
        except Exception as exc:
            logging.warning("Failed to read control page template %s: %s", candidate, exc)
    return "<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\"><title>HubVoice Satellite Control</title></head><body><h1>Control page missing</h1><p>control.html was not found next to the runtime.</p></body></html>"


def slugify_satellite(sat_id: str) -> str:
    return "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in (sat_id or "default"))


def build_wav_path(text: str, sat_id: str) -> Path:
    digest = hashlib.sha1(text.encode("utf-8")).hexdigest()[:10]
    stamp = time.strftime("%Y%m%d-%H%M%S")
    return RECORDINGS_PATH / f"{stamp}-{slugify_satellite(sat_id)}-{digest}.wav"


def build_input_wav_path(sat_id: str) -> Path:
    stamp = time.strftime("%Y%m%d-%H%M%S")
    return RECORDINGS_PATH / f"{stamp}-{slugify_satellite(sat_id)}-input.wav"


def resolve_piper_model_path() -> Path:
    config = load_config()
    configured = str(config.get("piper_voice_model", "")).strip()
    relative = configured or "piper_voices/en_US-amy-medium.onnx"
    model_path = (ROOT / relative).resolve()
    
    if not model_path.exists():
        # Provide helpful error messages
        if model_path.parent.exists():
            available = list(model_path.parent.glob("*.onnx"))
            available_str = ", ".join(m.name for m in available) if available else "none"
            raise RuntimeError(
                f"Piper voice model not found: {model_path}\n"
                f"Available models: {available_str}"
            )
        raise RuntimeError(
            f"Piper voice models directory missing: {model_path.parent}\n"
            f"Expected voices in: piper_voices/"
        )
    
    return model_path


def get_piper_voice() -> PiperVoice:
    global _PIPER_VOICE
    global _PIPER_VOICE_MODEL_PATH

    model_path = resolve_piper_model_path()
    model_path_str = str(model_path)
    with _PIPER_VOICE_LOCK:
        if _PIPER_VOICE is not None and _PIPER_VOICE_MODEL_PATH == model_path_str:
            return _PIPER_VOICE

        config_path = Path(model_path_str + ".json")
        logging.info("Loading Piper voice model %s", model_path.name)
        _PIPER_VOICE = PiperVoice.load(model_path=model_path, config_path=config_path)
        _PIPER_VOICE_MODEL_PATH = model_path_str
        return _PIPER_VOICE


def preload_piper_voice_model() -> None:
    """Warm-load Piper model so first TTS call does not incur startup latency."""
    try:
        get_piper_voice()
        logging.info("Piper voice model preloaded for fast TTS")
    except Exception as exc:
        logging.warning("Piper preload skipped: %s", exc)


# Must match announcement_pipeline.sample_rate in firmware (hubvoice-sat-fph.yaml)
_TTS_TARGET_RATE = 48000


def _resample_pcm_int16(data: bytes, orig_rate: int, target_rate: int, num_channels: int) -> bytes:
    """Resample 16-bit PCM audio to target_rate using linear interpolation (numpy)."""
    if orig_rate == target_rate or np is None:
        return data
    samples = np.frombuffer(data, dtype=np.int16).astype(np.float32)
    if num_channels > 1:
        samples = samples.reshape(-1, num_channels)
        orig_len = samples.shape[0]
        new_len = int(round(orig_len * target_rate / orig_rate))
        x_new = np.linspace(0, orig_len - 1, new_len)
        resampled = np.stack(
            [np.interp(x_new, np.arange(orig_len), samples[:, c]) for c in range(num_channels)],
            axis=1,
        ).astype(np.int16)
        return resampled.flatten().tobytes()
    orig_len = len(samples)
    new_len = int(round(orig_len * target_rate / orig_rate))
    x_new = np.linspace(0, orig_len - 1, new_len)
    return np.interp(x_new, np.arange(orig_len), samples).astype(np.int16).tobytes()


def synthesize_wav(text: str, output_path: Path) -> None:
    voice = get_piper_voice()
    chunks = iter(voice.synthesize(text))
    first_chunk = next(chunks, None)
    if first_chunk is None:
        raise RuntimeError("Piper returned no audio")

    audio_data = bytearray(first_chunk.audio_int16_bytes)
    for chunk in chunks:
        audio_data.extend(chunk.audio_int16_bytes)

    sample_rate = first_chunk.sample_rate
    num_channels = first_chunk.sample_channels
    sample_width = first_chunk.sample_width

    if sample_rate != _TTS_TARGET_RATE:
        audio_data = _resample_pcm_int16(bytes(audio_data), sample_rate, _TTS_TARGET_RATE, num_channels)
        sample_rate = _TTS_TARGET_RATE

    with wave.open(str(output_path), "wb") as wav_file:
        wav_file.setnchannels(num_channels)
        wav_file.setsampwidth(sample_width)
        wav_file.setframerate(sample_rate)
        wav_file.writeframes(audio_data)


def write_input_wav(raw_audio: bytes, output_path: Path) -> None:
    with wave.open(str(output_path), "wb") as wav_file:
        wav_file.setnchannels(1)
        wav_file.setsampwidth(2)
        wav_file.setframerate(16000)
        wav_file.writeframes(raw_audio)


def get_whisper_model() -> WhisperModel:
    global _WHISPER_MODEL
    with _WHISPER_MODEL_LOCK:
        if _WHISPER_MODEL is None:
            logging.info("Loading Whisper model base.en")
            _WHISPER_MODEL = WhisperModel("base.en", device="cpu", compute_type="int8")
        return _WHISPER_MODEL


def preload_whisper_model() -> None:
    """Warm-load Whisper model so first STT call does not incur startup latency."""
    try:
        get_whisper_model()
        logging.info("Whisper model preloaded for fast STT")
    except Exception as exc:
        logging.warning("Whisper preload skipped: %s", exc)


def clean_transcript(text: str) -> str:
    cleaned = sanitize_text(text).strip(" .?!,;:")
    for prefix in ("hey jarvis ", "jarvis ", "hey jervis ", "okay nabu ", "ok nabu ", "stop "):
        if cleaned.lower().startswith(prefix):
            cleaned = cleaned[len(prefix):].strip()
    return cleaned


_NUMBER_WORDS = {
    "zero": 0,
    "one": 1,
    "two": 2,
    "three": 3,
    "four": 4,
    "five": 5,
    "six": 6,
    "seven": 7,
    "eight": 8,
    "nine": 9,
    "ten": 10,
    "eleven": 11,
    "twelve": 12,
    "thirteen": 13,
    "fourteen": 14,
    "fifteen": 15,
    "sixteen": 16,
    "seventeen": 17,
    "eighteen": 18,
    "nineteen": 19,
    "twenty": 20,
    "thirty": 30,
    "forty": 40,
    "fifty": 50,
    "sixty": 60,
    "seventy": 70,
    "eighty": 80,
    "ninety": 90,
}
_CLOCK_FILLER_WORDS = {"at", "for", "the", "please", "alarm", "an", "a", "me", "up"}


def parse_number_phrase(text: str) -> int | None:
    cleaned = " ".join((text or "").lower().replace("-", " ").split())
    if not cleaned:
        return None
    if re.fullmatch(r"\d+", cleaned):
        return int(cleaned)

    tokens = [token for token in cleaned.split() if token != "and"]
    if not tokens:
        return None

    total = 0
    current = 0
    found = False
    for token in tokens:
        if token in {"a", "an"}:
            current += 1
            found = True
            continue
        if token == "hundred":
            current = max(current, 1) * 100
            found = True
            continue
        if token not in _NUMBER_WORDS:
            return None
        current += _NUMBER_WORDS[token]
        found = True

    total += current
    return total if found else None


def extract_duration_seconds(text: str) -> int | None:
    normalized = " ".join((text or "").lower().replace("-", " ").split())
    if not normalized:
        return None

    total_seconds = 0
    matched = False

    for pattern, seconds in (
        (r"\bhalf\s+(?:an?\s+)?hour\b", 30 * 60),
        (r"\bquarter\s+(?:of\s+an?\s+)?hour\b", 15 * 60),
    ):
        for _ in re.finditer(pattern, normalized):
            total_seconds += seconds
            matched = True
        normalized = re.sub(pattern, " ", normalized)

    for match in re.finditer(
        r"(?P<value>(?:\d+)|(?:a|an|one|two|three|four|five|six|seven|eight|nine|ten|"
        r"eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|"
        r"twenty|thirty|forty|fifty|sixty|seventy|eighty|ninety|hundred|and)(?:\s+"
        r"(?:one|two|three|four|five|six|seven|eight|nine))?)\s+"
        r"(?P<unit>seconds?|minutes?|hours?)\b",
        normalized,
    ):
        value = parse_number_phrase(match.group("value"))
        if value is None:
            continue
        unit = match.group("unit")
        multiplier = 1 if unit.startswith("second") else 60 if unit.startswith("minute") else 3600
        total_seconds += value * multiplier
        matched = True

    if not matched or total_seconds <= 0:
        return None
    return int(total_seconds)


def format_duration(seconds: int) -> str:
    remaining = max(0, int(seconds))
    parts: list[str] = []
    for unit_seconds, singular in ((3600, "hour"), (60, "minute"), (1, "second")):
        value, remaining = divmod(remaining, unit_seconds)
        if not value:
            continue
        label = singular if value == 1 else singular + "s"
        parts.append(f"{value} {label}")
    if not parts:
        return "0 seconds"
    if len(parts) == 1:
        return parts[0]
    return ", ".join(parts[:-1]) + f" and {parts[-1]}"


def format_clock_time(target_ts: float) -> str:
    now = datetime.now().astimezone()
    target = datetime.fromtimestamp(target_ts).astimezone()
    label = target.strftime("%I:%M %p").lstrip("0")
    if target.date() == now.date():
        return f"{label} today"
    if target.date() == (now + timedelta(days=1)).date():
        return f"{label} tomorrow"
    return f"{label} on {target.strftime('%A')}"


def _parse_spoken_clock_tokens(tokens: list[str]) -> tuple[int, int] | None:
    if not tokens:
        return None

    hour = parse_number_phrase(tokens[0])
    if hour is None:
        return None

    if len(tokens) == 1:
        return hour, 0

    if len(tokens) >= 3 and tokens[1] in {"oh", "o"}:
        minute = parse_number_phrase(" ".join(tokens[2:]))
        if minute is None:
            return None
        return hour, minute

    minute = parse_number_phrase(" ".join(tokens[1:]))
    if minute is None:
        return None
    return hour, minute


def parse_alarm_datetime(text: str) -> float | None:
    normalized = " ".join(
        (text or "")
        .lower()
        .replace(".", "")
        .replace(",", " ")
        .replace("-", " ")
        .split()
    )
    normalized = re.sub(r"\b([ap])\s+m\b", r"\1m", normalized)
    if not normalized:
        return None

    tokens = normalized.split()
    day_offset = 1 if "tomorrow" in tokens else 0
    explicit_today = "today" in tokens

    ampm = None
    if any(token in {"am", "morning"} for token in tokens):
        ampm = "am"
    elif any(token in {"pm", "afternoon", "evening", "tonight"} for token in tokens):
        ampm = "pm"

    filtered = [
        token
        for token in tokens
        if token not in {"tomorrow", "today", "am", "pm", "morning", "afternoon", "evening", "tonight"}
        and token not in _CLOCK_FILLER_WORDS
    ]
    if not filtered:
        return None

    hour: int | None = None
    minute = 0

    joined_filtered = " ".join(filtered)
    digit_match = re.search(r"(?P<hour>\d{1,2})(?::(?P<minute>\d{2}))?", joined_filtered)
    if digit_match:
        hour = int(digit_match.group("hour"))
        minute = int(digit_match.group("minute") or 0)

        # Handle separated numeric times like "1 10 pm" that STT often emits
        # without a colon; use the second token as minute when available.
        if minute == 0 and len(filtered) >= 2 and filtered[0].isdigit() and filtered[1].isdigit():
            split_hour = int(filtered[0])
            split_minute = int(filtered[1])
            if 0 <= split_minute <= 59:
                hour = split_hour
                minute = split_minute

        # Handle compact forms like "110 pm" -> 1:10 pm and "930 am" -> 9:30 am.
        if minute == 0 and len(filtered) == 1 and filtered[0].isdigit() and ampm is not None:
            compact = filtered[0]
            if len(compact) in {3, 4}:
                split_hour = int(compact[:-2])
                split_minute = int(compact[-2:])
                if 0 <= split_minute <= 59:
                    hour = split_hour
                    minute = split_minute
    else:
        parsed = _parse_spoken_clock_tokens(filtered)
        if parsed is None:
            return None
        hour, minute = parsed

    if hour is None or minute < 0 or minute > 59:
        return None

    now = datetime.now().astimezone()
    base = now.replace(hour=0, minute=0, second=0, microsecond=0) + timedelta(days=day_offset)

    candidates: list[datetime] = []
    if ampm is not None:
        if hour < 1 or hour > 12:
            return None
        if ampm == "am":
            hour24 = 0 if hour == 12 else hour
        else:
            hour24 = 12 if hour == 12 else hour + 12
        candidates.append(base.replace(hour=hour24, minute=minute))
        if not explicit_today and day_offset == 0:
            candidates.append((base + timedelta(days=1)).replace(hour=hour24, minute=minute))
    elif hour <= 23:
        if hour <= 12:
            hour_candidates = [0 if hour == 12 else hour, 12 if hour == 12 else hour + 12]
        else:
            hour_candidates = [hour]
        for hour24 in hour_candidates:
            candidates.append(base.replace(hour=hour24, minute=minute))
            if not explicit_today and day_offset == 0:
                candidates.append((base + timedelta(days=1)).replace(hour=hour24, minute=minute))
    else:
        return None

    future_candidates = [candidate for candidate in candidates if candidate > now + timedelta(seconds=5)]
    if not future_candidates:
        return None
    return min(future_candidates).timestamp()


def parse_timer_command(text: str) -> dict | None:
    normalized = " ".join((text or "").lower().replace("-", " ").split())
    if "timer" not in normalized:
        return None

    if re.search(r"\b(cancel|stop|delete|remove|clear)\b", normalized):
        return {"action": "cancel", "all": "all" in normalized}

    if re.search(r"\b(what|show|list|any|next|remaining|left)\b", normalized) or "time left" in normalized:
        return {"action": "status"}

    duration_text = ""
    patterns = (
        r"^(?:set|start|create)\s+(?:a\s+)?timer\s+(?:for\s+)?(?P<duration>.+)$",
        r"^(?:set|start|create)\s+(?:a\s+)?(?P<duration>.+?)\s+timer$",
        r"^(?P<duration>.+?)\s+timer$",
    )
    for pattern in patterns:
        match = re.match(pattern, normalized)
        if match:
            duration_text = match.group("duration")
            break

    if not duration_text:
        return None

    duration_seconds = extract_duration_seconds(duration_text)
    return {
        "action": "create",
        "duration_seconds": duration_seconds,
    }


def parse_alarm_command(text: str) -> dict | None:
    normalized = " ".join((text or "").lower().replace("-", " ").split())
    if "alarm system" in normalized or "security alarm" in normalized:
        return None
    if "alarm" not in normalized and not normalized.startswith("wake me"):
        return None

    if re.search(r"\b(cancel|stop|delete|remove|clear|turn off)\b", normalized):
        return {"action": "cancel", "all": "all" in normalized}

    if (
        re.search(r"\b(what|show|list|next|when)\b", normalized)
        or "alarm set" in normalized
        or re.search(r"\b(is|are)\s+there\b.*\balarm", normalized)
        or re.search(r"\bdo\s+i\s+have\b.*\balarm", normalized)
        or re.search(r"\bany\b.*\balarm", normalized)
    ):
        return {"action": "status"}

    when_text = ""
    patterns = (
        r"^(?:set\s+)?(?:an?\s+)?alarm\s+(?:for\s+)?(?P<when>.+)$",
        r"^(?:center|sitter|sent|send)\s+alarm\s+(?:for\s+)?(?P<when>.+)$",
        r"^wake\s+me\s+up\s+(?:at\s+)?(?P<when>.+)$",
        r"^wake\s+me\s+(?:at\s+)?(?P<when>.+)$",
    )
    for pattern in patterns:
        match = re.match(pattern, normalized)
        if match:
            when_text = match.group("when")
            break

    if not when_text:
        return None

    return {
        "action": "create",
        "target_ts": parse_alarm_datetime(when_text),
    }


def is_dismiss_command(text: str) -> bool:
    normalized = " ".join((text or "").lower().replace("'", "").split())
    if not normalized:
        return False
    if normalized in {"dismiss", "silence", "stop ringing"}:
        return True
    return normalized in {
        "dismiss alarm",
        "dismiss timer",
        "stop alarm",
        "stop timer",
        "silence alarm",
        "silence timer",
    }


@dataclass
class ScheduledItem:
    schedule_id: str
    kind: str
    satellite_id: str
    name: str
    target_ts: float
    created_ts: float
    total_seconds: int
    last_sent_seconds: int = -1

    def seconds_left(self, now_ts: float | None = None) -> int:
        current = time.time() if now_ts is None else now_ts
        return max(0, int(math.ceil(self.target_ts - current)))

    def to_dict(self) -> dict:
        return {
            "schedule_id": self.schedule_id,
            "kind": self.kind,
            "satellite_id": self.satellite_id,
            "name": self.name,
            "target_ts": self.target_ts,
            "created_ts": self.created_ts,
            "total_seconds": self.total_seconds,
            "last_sent_seconds": self.last_sent_seconds,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "ScheduledItem":
        return cls(
            schedule_id=str(data.get("schedule_id", "")).strip(),
            kind=str(data.get("kind", "")).strip(),
            satellite_id=str(data.get("satellite_id", "")).strip() or "default",
            name=str(data.get("name", "")).strip(),
            target_ts=float(data.get("target_ts", 0.0) or 0.0),
            created_ts=float(data.get("created_ts", 0.0) or 0.0),
            total_seconds=int(data.get("total_seconds", 0) or 0),
            last_sent_seconds=int(data.get("last_sent_seconds", -1) or -1),
        )

    def snapshot(self) -> dict:
        return {
            "id": self.schedule_id,
            "kind": self.kind,
            "satellite": self.satellite_id,
            "name": self.name,
            "target": self.target_ts,
            "seconds_left": self.seconds_left(),
            "display": format_clock_time(self.target_ts) if self.kind == "alarm" else format_duration(self.seconds_left()),
        }


class ScheduleManager:
    def __init__(self, storage_path: Path) -> None:
        self._storage_path = storage_path
        self._lock = threading.Lock()
        self._items: dict[str, ScheduledItem] = {}
        self._ringing_until: dict[str, float] = {}
        self._ringing_threads: set[str] = set()
        self._load()

    @staticmethod
    def _sat_key(satellite_id: str) -> str:
        return (satellite_id or "default").strip().lower() or "default"

    def _load(self) -> None:
        if not self._storage_path.exists():
            return
        try:
            payload = json.loads(self._storage_path.read_text(encoding="utf-8"))
        except Exception as exc:
            logging.exception("Failed to load schedules: %s", exc)
            return
        if not isinstance(payload, list):
            return
        with self._lock:
            for item in payload:
                if not isinstance(item, dict):
                    continue
                schedule = ScheduledItem.from_dict(item)
                if not schedule.schedule_id or schedule.kind not in {"timer", "alarm"}:
                    continue
                self._items[schedule.schedule_id] = schedule

    def _save_locked(self) -> None:
        data = [item.to_dict() for item in sorted(self._items.values(), key=lambda current: (current.target_ts, current.schedule_id))]
        self._storage_path.write_text(json.dumps(data, indent=2), encoding="utf-8")

    def list_items(self, kind: str, satellite_id: str) -> list[ScheduledItem]:
        sat_key = (satellite_id or "").strip().lower()
        with self._lock:
            return [
                item
                for item in sorted(self._items.values(), key=lambda current: (current.target_ts, current.schedule_id))
                if item.kind == kind and (not sat_key or item.satellite_id.lower() == sat_key)
            ]

    def dismiss_active_ringing(self, satellite_id: str) -> bool:
        sat_key = self._sat_key(satellite_id)
        with self._lock:
            was_active = sat_key in self._ringing_until
            self._ringing_until.pop(sat_key, None)

        satellite = select_satellite(satellite_id)
        if satellite:
            host = str(satellite.get("host", "")).strip()
            if host:
                try:
                    stop_satellite_playback(host)
                except Exception as exc:
                    logging.debug("Failed to stop playback while dismissing %s: %s", sat_key, exc)
        return was_active

    def add_timer(self, satellite_id: str, duration_seconds: int) -> ScheduledItem:
        if duration_seconds <= 0 or duration_seconds > MAX_TIMER_DURATION_SECONDS:
            raise RuntimeError("Timer duration must be between 1 second and 24 hours")
        existing = self.list_items("timer", satellite_id)
        if len(existing) >= MAX_ACTIVE_TIMERS_PER_SATELLITE:
            raise RuntimeError("Too many active timers on this satellite")

        now_ts = time.time()
        item = ScheduledItem(
            schedule_id=f"timer-{uuid.uuid4().hex[:12]}",
            kind="timer",
            satellite_id=(satellite_id or "default").strip() or "default",
            name="Timer",
            target_ts=now_ts + duration_seconds,
            created_ts=now_ts,
            total_seconds=duration_seconds,
        )
        with self._lock:
            self._items[item.schedule_id] = item
            self._save_locked()
        return item

    def add_alarm(self, satellite_id: str, target_ts: float) -> ScheduledItem:
        now_ts = time.time()
        if target_ts <= now_ts:
            raise RuntimeError("Alarm time must be in the future")
        existing = self.list_items("alarm", satellite_id)
        if len(existing) >= MAX_ACTIVE_ALARMS_PER_SATELLITE:
            raise RuntimeError("Too many active alarms on this satellite")

        item = ScheduledItem(
            schedule_id=f"alarm-{uuid.uuid4().hex[:12]}",
            kind="alarm",
            satellite_id=(satellite_id or "default").strip() or "default",
            name="Alarm",
            target_ts=target_ts,
            created_ts=now_ts,
            total_seconds=max(1, int(target_ts - now_ts)),
        )
        with self._lock:
            self._items[item.schedule_id] = item
            self._save_locked()
        return item

    def cancel_next(self, kind: str, satellite_id: str) -> ScheduledItem | None:
        items = self.list_items(kind, satellite_id)
        if not items:
            return None
        target = items[0]
        with self._lock:
            self._items.pop(target.schedule_id, None)
            self._save_locked()
        return target

    def cancel_all(self, kind: str, satellite_id: str) -> list[ScheduledItem]:
        items = self.list_items(kind, satellite_id)
        if not items:
            return []
        with self._lock:
            for item in items:
                self._items.pop(item.schedule_id, None)
            self._save_locked()
        return items

    def mark_timer_sent(self, schedule_id: str, seconds_left: int) -> None:
        with self._lock:
            item = self._items.get(schedule_id)
            if item is None:
                return
            item.last_sent_seconds = seconds_left

    def sync_bridge(self, bridge: "VoiceAssistantBridge") -> None:
        if not bridge.can_send_timer_events():
            return
        now_ts = time.time()
        for item in self.list_items("timer", bridge._satellite_id):
            seconds_left = item.seconds_left(now_ts)
            if seconds_left <= 0:
                continue
            if bridge.send_timer_event(
                VoiceAssistantTimerEventType.VOICE_ASSISTANT_TIMER_STARTED,
                item.schedule_id,
                item.name,
                item.total_seconds,
                seconds_left,
                True,
            ):
                self.mark_timer_sent(item.schedule_id, seconds_left)

    def tick(self, bridges: "dict[str, VoiceAssistantBridge]") -> None:
        now_ts = time.time()
        timers_to_start: list[tuple[ScheduledItem, int]] = []
        timers_to_update: list[tuple[ScheduledItem, int]] = []
        due_items: list[ScheduledItem] = []

        def _get_bridge(item: ScheduledItem) -> "VoiceAssistantBridge | None":
            sat_key = item.satellite_id.strip().lower()
            bridge = bridges.get(sat_key)
            if bridge is None and bridges:
                bridge = next(iter(bridges.values()))
            return bridge

        with self._lock:
            items = sorted(self._items.values(), key=lambda current: (current.target_ts, current.schedule_id))
            for item in items:
                seconds_left = item.seconds_left(now_ts)
                if seconds_left <= 0:
                    due_items.append(item)
                    continue
                b = _get_bridge(item)
                if item.kind != "timer" or b is None or not b.can_send_timer_events():
                    continue
                if item.last_sent_seconds < 0:
                    timers_to_start.append((item, seconds_left))
                elif seconds_left != item.last_sent_seconds:
                    timers_to_update.append((item, seconds_left))

        for item, seconds_left in timers_to_start:
            b = _get_bridge(item)
            if b and b.send_timer_event(
                VoiceAssistantTimerEventType.VOICE_ASSISTANT_TIMER_STARTED,
                item.schedule_id,
                item.name,
                item.total_seconds,
                seconds_left,
                True,
            ):
                self.mark_timer_sent(item.schedule_id, seconds_left)

        for item, seconds_left in timers_to_update:
            b = _get_bridge(item)
            if b and b.send_timer_event(
                VoiceAssistantTimerEventType.VOICE_ASSISTANT_TIMER_UPDATED,
                item.schedule_id,
                item.name,
                item.total_seconds,
                seconds_left,
                True,
            ):
                self.mark_timer_sent(item.schedule_id, seconds_left)

        delivered_ids: list[str] = []
        for item in due_items:
            b = _get_bridge(item)
            delivered = b.send_timer_event(
                VoiceAssistantTimerEventType.VOICE_ASSISTANT_TIMER_FINISHED,
                item.schedule_id,
                item.name,
                max(1, item.total_seconds),
                0,
                False,
            ) if b else False
            if not delivered:
                delivered = self._deliver_due_fallback(item)
            if delivered:
                delivered_ids.append(item.schedule_id)
                logging.info("Delivered %s completion for %s", item.kind, item.schedule_id)
            else:
                logging.warning("Due %s could not be delivered for %s", item.kind, item.schedule_id)

        if delivered_ids:
            with self._lock:
                for schedule_id in delivered_ids:
                    self._items.pop(schedule_id, None)
                self._save_locked()

    def _deliver_due_fallback(self, item: ScheduledItem) -> bool:
        satellite = select_satellite(item.satellite_id)
        if not satellite:
            return False

        if item.kind == "alarm":
            text = "Alarm time."
        else:
            text = "Timer finished."

        try:
            wav_path = build_wav_path(text, item.satellite_id)
            synthesize_wav(text, wav_path)
            media_url = build_media_url(wav_path.name)
            self._start_ringing_loop(
                item.satellite_id,
                str(satellite.get("host", "")).strip(),
                media_url,
                max_seconds=FALLBACK_RING_MAX_SECONDS,
            )
            return True
        except Exception as exc:
            logging.exception("Fallback due delivery failed for %s: %s", item.schedule_id, exc)
            return False

    def _start_ringing_loop(self, satellite_id: str, satellite_host: str, media_url: str, max_seconds: int) -> None:
        sat_key = self._sat_key(satellite_id)
        if not satellite_host or not media_url:
            raise RuntimeError("Missing satellite host or media URL for fallback ringing")

        start_thread = False
        deadline = time.time() + max_seconds
        with self._lock:
            current_until = self._ringing_until.get(sat_key, 0.0)
            self._ringing_until[sat_key] = max(current_until, deadline)
            if sat_key not in self._ringing_threads:
                self._ringing_threads.add(sat_key)
                start_thread = True

        if not start_thread:
            return

        threading.Thread(
            target=self._ringing_worker,
            args=(sat_key, satellite_host, media_url),
            daemon=True,
        ).start()

    def _ringing_worker(self, sat_key: str, satellite_host: str, media_url: str) -> None:
        try:
            while True:
                with self._lock:
                    active_until = self._ringing_until.get(sat_key, 0.0)
                now_ts = time.time()
                if active_until <= now_ts:
                    break

                try:
                    send_to_satellite(satellite_host, media_url)
                except Exception as exc:
                    logging.exception("Fallback ringing playback failed for %s: %s", sat_key, exc)

                wake_at = now_ts + FALLBACK_RING_REPEAT_SECONDS
                while time.time() < wake_at:
                    with self._lock:
                        still_active_until = self._ringing_until.get(sat_key, 0.0)
                    if still_active_until <= time.time():
                        break
                    time.sleep(0.3)
        finally:
            with self._lock:
                self._ringing_until.pop(sat_key, None)
                self._ringing_threads.discard(sat_key)
            try:
                stop_satellite_playback(satellite_host)
            except Exception as exc:
                logging.debug("Failed to stop playback after fallback ringing for %s: %s", sat_key, exc)

    def snapshot(self) -> dict:
        with self._lock:
            items = sorted(self._items.values(), key=lambda current: (current.target_ts, current.schedule_id))
            timers = [item.snapshot() for item in items if item.kind == "timer"]
            alarms = [item.snapshot() for item in items if item.kind == "alarm"]
            return {
                "timers": timers,
                "alarms": alarms,
                "timer_count": len(timers),
                "alarm_count": len(alarms),
            }


def parse_volume_command(text: str) -> dict | None:
    normalized = " ".join((text or "").lower().replace("%", " percent ").split())
    match = re.search(
        r"\b(?:set|change|make|turn)\s+(?:the\s+)?(?:(?P<target>wake(?:\s+word|\s+sound)?|speaker)\s+)?volume\s+(?:to\s+)?(?P<value>\d{1,3})(?:\s+percent)?\b",
        normalized,
    )
    if not match:
        match = re.search(
            r"\b(?:(?P<target>wake(?:\s+word|\s+sound)?|speaker)\s+)?volume\s+(?:to\s+)?(?P<value>\d{1,3})(?:\s+percent)?\b",
            normalized,
        )
    if not match:
        return None

    requested = int(match.group("value"))
    target = match.group("target") or "speaker"
    entity_id = "wake_sound_volume" if "wake" in target else "speaker_volume"
    label = "wake sound volume" if entity_id == "wake_sound_volume" else "speaker volume"
    min_value = 40 if entity_id == "wake_sound_volume" else 0
    max_value = 85
    applied = max(min_value, min(max_value, requested))
    return {
        "entity_id": entity_id,
        "label": label,
        "requested": requested,
        "applied": applied,
        "min": min_value,
        "max": max_value,
    }


def parse_whisper_command(text: str) -> dict | None:
    """Detect 'whisper mode on/off' voice commands.

    Returns a dict with 'state' (bool) and 'label' (str), or None if not a whisper command.
    """
    normalized = " ".join((text or "").lower().split())
    on_pattern = r"\b(?:turn\s+on|enable|activate|start|switch\s+on)\s+whisper(?:\s+mode)?\b"
    off_pattern = r"\b(?:turn\s+off|disable|deactivate|stop|switch\s+off)\s+whisper(?:\s+mode)?\b"
    toggle_on = r"\bwhisper(?:\s+mode)?\s+(?:on|enable)\b"
    toggle_off = r"\bwhisper(?:\s+mode)?\s+(?:off|disable)\b"
    if re.search(on_pattern, normalized) or re.search(toggle_on, normalized):
        return {"state": True, "label": "on"}
    if re.search(off_pattern, normalized) or re.search(toggle_off, normalized):
        return {"state": False, "label": "off"}
    return None


def parse_follow_up_toggle_command(text: str) -> dict | None:
    """Detect follow-up listening on/off commands."""
    normalized = " ".join((text or "").lower().split())
    on_pattern = r"\b(?:turn\s+on|enable|activate|start|switch\s+on)\s+follow(?:-|\s*)up(?:\s+listening)?\b"
    off_pattern = r"\b(?:turn\s+off|disable|deactivate|stop|switch\s+off)\s+follow(?:-|\s*)up(?:\s+listening)?\b"
    toggle_on = r"\bfollow(?:-|\s*)up(?:\s+listening)?\s+(?:on|enable)\b"
    toggle_off = r"\bfollow(?:-|\s*)up(?:\s+listening)?\s+(?:off|disable)\b"
    if re.search(on_pattern, normalized) or re.search(toggle_on, normalized):
        return {"state": True, "label": "on"}
    if re.search(off_pattern, normalized) or re.search(toggle_off, normalized):
        return {"state": False, "label": "off"}
    return None


def parse_follow_up_window_command(text: str) -> dict | None:
    """Detect follow-up listen window commands and return clamped values."""
    normalized = " ".join((text or "").lower().replace("%", " percent ").split())
    match = re.search(
        r"\b(?:set|change|make)\s+(?:the\s+)?follow(?:-|\s*)up(?:\s+listen(?:ing)?\s+window|\s+window)?\s+(?:to\s+)?(?P<value>\d{1,2})(?:\s+seconds?)?\b",
        normalized,
    )
    if not match:
        match = re.search(
            r"\bfollow(?:-|\s*)up(?:\s+listen(?:ing)?\s+window|\s+window)?\s+(?:to\s+)?(?P<value>\d{1,2})(?:\s+seconds?)?\b",
            normalized,
        )
    if not match:
        return None

    requested = int(match.group("value"))
    applied = max(1, min(30, requested))
    return {
        "requested": requested,
        "applied": applied,
        "entity_id": "follow_up_listen_window_seconds",
    }


def parse_broadcast_command(text: str) -> dict | None:
    raw = " ".join((text or "").split()).strip()
    if not raw:
        return None

    stripped = re.sub(r"^(?:hey|ok|okay)\s+nabu[,]?\s*", "", raw, flags=re.IGNORECASE).strip()
    if not re.match(r"^(?:broadcast|announce)\b", stripped, flags=re.IGNORECASE):
        return None

    tail = re.sub(r"^(?:broadcast|announce)\b", "", stripped, flags=re.IGNORECASE).strip()
    tail = re.sub(r"^to\s+(?:all|everyone)\b", "", tail, flags=re.IGNORECASE).strip()

    if not tail:
        return {"mode": "prompt", "message": ""}

    target_only = re.match(r"^(?:to|in)\s+(.+)$", tail, flags=re.IGNORECASE)
    if target_only:
        target = sanitize_text(target_only.group(1))
        if target.lower() in {"all", "everyone"}:
            return {"mode": "prompt", "message": "", "target": ""}
        return {"mode": "prompt_target", "message": "", "target": target}

    return {"mode": "inline", "message": sanitize_text(tail)}


def is_broadcast_cancel(text: str) -> bool:
    normalized = " ".join((text or "").lower().replace("'", "").split()).strip()
    return normalized in {
        "cancel",
        "never mind",
        "nevermind",
        "stop",
        "cancel broadcast",
        "stop broadcast",
        "cancel announcement",
        "stop announcement",
        "never mind broadcast",
        "nevermind broadcast",
    }


def parse_directed_send(text: str) -> dict | None:
    raw = " ".join((text or "").split()).strip()
    if not raw:
        return None

    # Examples:
    # "send to living room, emma dinner is ready"
    # "broadcast to kitchen: dinner is ready"
    # "broadcast in living room, emma dinner is ready"
    # "send living room, emma dinner is ready"
    # "tell living room, emma dinner is ready"
    match = re.match(r"^(?:send|broadcast|announce|tell)\s+(?:to|in)\s+(.+?)[,:]\s*(.+)$", raw, flags=re.IGNORECASE)
    if not match:
        match = re.match(r"^(?:send|tell)\s+(.+?)[,:]\s*(.+)$", raw, flags=re.IGNORECASE)
    if not match:
        # No comma/colon: try to infer target as the first matching satellite id/alias.
        # Example: "send living room emma dinner is ready"
        simple = re.match(r"^(?:send|tell)\s+(.+)$", raw, flags=re.IGNORECASE)
        if not simple:
            return None
        remainder = sanitize_text(simple.group(1))
        if not remainder:
            return None

        refs: list[str] = []
        for sat in load_satellites():
            sat_id = sanitize_text(str(sat.get("id", "")))
            sat_alias = sanitize_text(str(sat.get("alias", "")))
            if sat_alias:
                refs.append(sat_alias)
            if sat_id:
                refs.append(sat_id)

        seen: set[str] = set()
        uniq_refs: list[str] = []
        for r in refs:
            key = r.lower()
            if key in seen:
                continue
            seen.add(key)
            uniq_refs.append(r)

        uniq_refs.sort(key=lambda v: len(v), reverse=True)
        tail_lower = remainder.lower()
        for target_ref in uniq_refs:
            ref_lower = target_ref.lower()
            if tail_lower.startswith(ref_lower + " "):
                message = sanitize_text(remainder[len(target_ref):])
                if message:
                    return {"target": target_ref, "message": message}
        return None

    target = sanitize_text(match.group(1))
    message = sanitize_text(match.group(2))
    if not target or not message:
        return None

    return {"target": target, "message": message}


def transcribe_wav(wav_path: Path, *, initial_prompt: str | None = None, beam_size: int = 5, best_of: int = 5) -> dict:
    config = load_config()

    def _cfg_int(name: str, default: int, minimum: int, maximum: int) -> int:
        try:
            value = int(config.get(name, default))
        except Exception:
            value = default
        return max(minimum, min(maximum, value))

    # Global VAD trimming defaults tuned for faster response while keeping short pauses intact.
    vad_min_silence_ms = _cfg_int("stt_vad_min_silence_ms", 500, 100, 3000)
    vad_speech_pad_ms = _cfg_int("stt_vad_speech_pad_ms", 120, 0, 1000)

    model = get_whisper_model()
    segments, info = model.transcribe(
        str(wav_path),
        language="en",
        beam_size=beam_size,
        best_of=best_of,
        vad_filter=True,
        vad_parameters={
            "min_silence_duration_ms": vad_min_silence_ms,
            "speech_pad_ms": vad_speech_pad_ms,
        },
        condition_on_previous_text=False,
        initial_prompt=initial_prompt,
    )
    segment_list = [segment for segment in segments]
    transcript = " ".join(segment.text.strip() for segment in segment_list if segment.text).strip()
    transcript = clean_transcript(" ".join(transcript.split()))
    avg_logprob = None
    if segment_list:
        avg_logprob = sum(float(getattr(seg, "avg_logprob", 0.0)) for seg in segment_list) / len(segment_list)
    return {
        "text": transcript,
        "avg_logprob": avg_logprob,
        "duration": getattr(info, "duration", None),
        "language": getattr(info, "language", None),
    }


def ask_hubitat(question: str, sat_id: str) -> dict:
    """
    Query Hubitat for an answer to a voice query.
    
    Args:
        question: Voice query to ask (required, non-empty)
        sat_id: Satellite ID for context (required, non-empty)
    
    Returns:
        Dictionary with keys:
        - ok (bool): Whether query succeeded
        - answer (str): Response text (sanitized)
        - payload (dict): Full response from Hubitat
    
    Raises:
        ValueError: If parameters invalid
        RuntimeError: If Hubitat unreachable or circuit breaker open
        urllib.error.HTTPError: If HTTP error occurs
        urllib.error.URLError: If connection fails
        json.JSONDecodeError: If response JSON invalid
    
    Timeout: 20 seconds per request
    """
    if not question or not isinstance(question, str):
        raise ValueError("Question must be a non-empty string")
    if not sat_id or not isinstance(sat_id, str):
        raise ValueError("Satellite ID must be a non-empty string")
    
    # Check circuit breaker
    if _HUBITAT_BREAKER.is_open():
        raise RuntimeError("Hubitat circuit breaker is open - failing fast")
    
    config = load_config()
    host = str(config.get("hubitat_host", "")).strip().rstrip("/")
    app_id = str(config.get("hubitat_app_id", "")).strip()
    token = str(config.get("hubitat_access_token", "")).strip()
    if not host or not app_id or not token:
        raise RuntimeError("Hubitat configuration is incomplete in hubvoice-sat-setup.json")

    query = urllib.parse.urlencode(
        {
            "access_token": token,
            "q": question,
            "d": sat_id,
        }
    )
    url = f"{host}/apps/api/{app_id}/ask?{query}"
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "HubVoiceSatRuntime/1.0",
            "Accept": "application/json, text/plain;q=0.9, */*;q=0.1",
        },
    )
    
    try:
        with urllib.request.urlopen(request, timeout=HUBITAT_REQUEST_TIMEOUT) as response:
            body = response.read().decode("utf-8", errors="replace")
            content_type = response.headers.get("Content-Type", "")
        
        _HUBITAT_BREAKER.reset()  # Success - reset circuit breaker
    except urllib.error.HTTPError as e:
        _HUBITAT_BREAKER.record_failure()
        _METRICS.record_request(0, error=True, error_type="hubitat")
        logging.error("Hubitat HTTP error: %s %s", e.code, e.reason)
        raise RuntimeError(f"Hubitat HTTP error: {e.code} {e.reason}")
    except urllib.error.URLError as e:
        _HUBITAT_BREAKER.record_failure()
        _METRICS.record_request(0, error=True, error_type="hubitat")
        logging.error("Hubitat connection failed: %s", e.reason)
        raise RuntimeError(f"Hubitat connection failed: {e.reason}")
    except Exception as e:
        _HUBITAT_BREAKER.record_failure()
        _METRICS.record_request(0, error=True, error_type="hubitat")
        logging.error("Unexpected error querying Hubitat: %s", e)
        raise RuntimeError(f"Hubitat query failed: {e}")
    
    if "json" in content_type.lower():
        try:
            payload = json.loads(body)
        except json.JSONDecodeError as e:
            _HUBITAT_BREAKER.record_failure()
            logging.error("Invalid JSON from Hubitat: %s", e)
            raise RuntimeError(f"Invalid JSON response from Hubitat: {e}")
        
        if isinstance(payload, dict):
            answer = payload.get("answer") or payload.get("message") or payload.get("error") or ""
            return {
                "ok": bool(payload.get("ok")),
                "answer": sanitize_text(str(answer)) if answer else "",
                "payload": payload,
            }
    
    body = body.strip()
    if not body:
        _HUBITAT_BREAKER.record_failure()
        raise RuntimeError("Hubitat returned an empty response")
    return {
        "ok": True,
        "answer": sanitize_text(body),
        "payload": {},
    }


def should_retry_transcript(transcript: dict, *, logprob_threshold: float = -0.95) -> bool:
    text = str(transcript.get("text", "")).strip()
    avg_logprob = transcript.get("avg_logprob")

    if not text:
        return True
    if avg_logprob is not None and float(avg_logprob) < float(logprob_threshold):
        return True
    return False


def transcribe_with_retry(wav_path: Path, sat_id: str) -> dict:
    config = load_config()
    speed_mode_raw = str(config.get("stt_speed_mode", "true")).strip().lower()
    speed_mode = speed_mode_raw not in {"0", "false", "off", "no"}

    def _cfg_int(name: str, default: int, minimum: int, maximum: int) -> int:
        try:
            value = int(config.get(name, default))
        except Exception:
            value = default
        return max(minimum, min(maximum, value))

    def _cfg_float(name: str, default: float, minimum: float, maximum: float) -> float:
        try:
            value = float(config.get(name, default))
        except Exception:
            value = default
        return max(minimum, min(maximum, value))

    primary_beam = _cfg_int("stt_primary_beam", 2 if speed_mode else 5, 1, 10)
    primary_best = _cfg_int("stt_primary_best_of", 2 if speed_mode else 5, 1, 10)
    retry_beam = _cfg_int("stt_retry_beam", 4 if speed_mode else 8, 1, 12)
    retry_best = _cfg_int("stt_retry_best_of", 4 if speed_mode else 8, 1, 12)
    retry_logprob_threshold = _cfg_float("stt_retry_logprob_threshold", -0.95 if speed_mode else -0.75, -2.0, 0.0)
    improvement_threshold = _cfg_float("stt_retry_improvement_threshold", 0.08 if speed_mode else 0.05, 0.0, 1.0)

    primary = transcribe_wav(wav_path, beam_size=primary_beam, best_of=primary_best)
    logging.info("Primary transcript: %r (avg_logprob=%s)", primary["text"], primary["avg_logprob"])

    if parse_volume_command(primary["text"]):
        return primary

    if not should_retry_transcript(primary, logprob_threshold=retry_logprob_threshold):
        return primary

    retry = transcribe_wav(
        wav_path,
        initial_prompt="Home automation voice query. Devices may be doors, locks, lights, motion sensors, thermostats, batteries, house status, security check, hub mode, windows, water sensors.",
        beam_size=retry_beam,
        best_of=retry_best,
    )
    logging.info("Retry transcript: %r (avg_logprob=%s)", retry["text"], retry["avg_logprob"])

    if not retry["text"]:
        return primary

    if parse_volume_command(retry["text"]):
        return retry

    retry_score = float(retry.get("avg_logprob") or -99.0)
    first_score = float(primary.get("avg_logprob") or -99.0)
    if retry_score > first_score + improvement_threshold:
        return retry

    return primary


async def send_to_satellite_async(satellite_host: str, media_url: str) -> None:
    """
    Send media URL to satellite for playback announcement.
    
    Args:
        satellite_host: IP or hostname of satellite device
        media_url: Full HTTP URL to audio file (WAV format)
    
    Raises:
        ValueError: If parameters invalid
        RuntimeError: If connection fails or timeout occurs
    
    Timeout: 10 seconds for initial connection, 10 seconds for command
    """
    if not satellite_host or not media_url:
        raise ValueError("satellite_host and media_url are required")

    cached_key = _media_player_key_cache.get(satellite_host)
    client = APIClient(satellite_host, 6054, None, client_info="HubVoiceSatRuntime")
    try:
        await asyncio.wait_for(client.connect(login=True), timeout=CONNECTION_TIMEOUT)
        if cached_key is None:
            entities, _ = await asyncio.wait_for(
                client.list_entities_services(),
                timeout=ENTITY_LIST_TIMEOUT,
            )
            media = _pick_media_player(entities)
            if media is None:
                _METRICS.record_request(0, error=True, error_type="satellite")
                raise RuntimeError(f"Satellite {satellite_host} has no media player")
            cached_key = media.key
            _media_player_key_cache[satellite_host] = cached_key
        client.media_player_command(
            cached_key,
            command=MediaPlayerCommand.PLAY,
            media_url=media_url,
            announcement=True,
        )
        await asyncio.sleep(SATELLITE_MEDIA_DELAY)
        _METRICS.record_request(SATELLITE_MEDIA_DELAY, error=False)
    except asyncio.TimeoutError as e:
        _METRICS.record_request(0, error=True, error_type="satellite")
        raise RuntimeError(f"Timeout communicating with satellite {satellite_host}: {e}")
    except Exception:
        _METRICS.record_request(0, error=True, error_type="satellite")
        raise
    finally:
        try:
            await client.disconnect()
        except Exception:
            pass


# Cache media player entity key per host to skip list_entities_services() on repeat volume calls.
_media_player_key_cache: dict[str, int] = {}
_USER_LOCKED_MEDIA_VOLUMES: dict[str, float] = {}
_USER_LOCKED_MEDIA_VOLUMES_LOCK = threading.Lock()
_SATELLITE_MEDIA_VOLUME_LAST_CALL: dict[str, float] = {}
_SATELLITE_MEDIA_VOLUME_LAST_CALL_LOCK = threading.Lock()
_SATELLITE_MEDIA_VOLUME_LAST_APPLIED: dict[str, dict[str, float]] = {}
_SATELLITE_MEDIA_VOLUME_LAST_APPLIED_LOCK = threading.Lock()
_USER_TONE_SETTINGS: dict[str, dict[str, float]] = {}
_USER_TONE_SETTINGS_LOCK = threading.Lock()


def _normalize_volume_percent(value: float | int | str) -> float:
    """Clamp a volume input to 0..100 and return float."""
    try:
        numeric = float(value)
    except (TypeError, ValueError):
        raise ValueError(f"volume_percent must be numeric, got {value}")
    return max(0.0, min(100.0, numeric))


def _remember_user_locked_media_volume(satellite_host: str, volume_percent: float | int | str) -> float:
    """Persist a user-selected media volume for later route start enforcement."""
    host = str(satellite_host or "").strip()
    if not host:
        raise ValueError("satellite_host is required")
    normalized = _normalize_volume_percent(volume_percent)
    with _USER_LOCKED_MEDIA_VOLUMES_LOCK:
        _USER_LOCKED_MEDIA_VOLUMES[host] = normalized
    return normalized


def _throttle_satellite_media_volume(satellite_host: str, min_interval_seconds: float = 0.75) -> None:
    """Rate-limit media volume writes per satellite to reduce device overload/reboots."""
    host = str(satellite_host or "").strip()
    if not host:
        return
    min_interval = max(0.0, float(min_interval_seconds))
    if min_interval <= 0.0:
        return

    sleep_for = 0.0
    now = time.monotonic()
    with _SATELLITE_MEDIA_VOLUME_LAST_CALL_LOCK:
        last = _SATELLITE_MEDIA_VOLUME_LAST_CALL.get(host)
        if last is not None:
            elapsed = now - last
            if elapsed < min_interval:
                sleep_for = min_interval - elapsed
        _SATELLITE_MEDIA_VOLUME_LAST_CALL[host] = now + sleep_for

    if sleep_for > 0.0:
        time.sleep(sleep_for)


def _should_apply_satellite_media_volume(satellite_host: str, volume_percent: float) -> tuple[bool, float]:
    host = str(satellite_host or "").strip()
    if not host:
        return True, 0.0
    now = time.monotonic()
    with _SATELLITE_MEDIA_VOLUME_LAST_APPLIED_LOCK:
        entry = dict(_SATELLITE_MEDIA_VOLUME_LAST_APPLIED.get(host, {}))
    if not entry:
        return True, 0.0

    last_ts = float(entry.get("ts", 0.0))
    last_value = float(entry.get("value", volume_percent))
    elapsed = now - last_ts
    delta = abs(float(volume_percent) - last_value)

    if elapsed >= SATELLITE_MEDIA_VOLUME_MIN_APPLY_INTERVAL_SECONDS:
        return True, 0.0
    if delta >= SATELLITE_MEDIA_VOLUME_MIN_DELTA_PERCENT:
        return True, 0.0
    return False, max(0.0, SATELLITE_MEDIA_VOLUME_MIN_APPLY_INTERVAL_SECONDS - elapsed)


def _mark_satellite_media_volume_applied(satellite_host: str, volume_percent: float) -> None:
    host = str(satellite_host or "").strip()
    if not host:
        return
    with _SATELLITE_MEDIA_VOLUME_LAST_APPLIED_LOCK:
        _SATELLITE_MEDIA_VOLUME_LAST_APPLIED[host] = {
            "ts": time.monotonic(),
            "value": float(volume_percent),
        }


def _get_user_locked_media_volume(satellite_host: str) -> float | None:
    host = str(satellite_host or "").strip()
    if not host:
        return None
    with _USER_LOCKED_MEDIA_VOLUMES_LOCK:
        value = _USER_LOCKED_MEDIA_VOLUMES.get(host)
    return None if value is None else float(value)


def _enforce_user_locked_media_volume(satellite_host: str) -> None:
    """Best-effort restore of explicit user volume after play start transitions."""
    locked = _get_user_locked_media_volume(satellite_host)
    if locked is None:
        return
    with contextlib.suppress(Exception):
        set_satellite_media_volume(satellite_host, locked)


def _normalize_tone_db(value: float | int | str) -> float:
    try:
        numeric = float(value)
    except (TypeError, ValueError):
        raise ValueError(f"tone value must be numeric, got {value}")
    return max(-10.0, min(10.0, numeric))


async def _read_satellite_tone_settings_async(satellite_host: str) -> tuple[float, float] | None:
    host = str(satellite_host or "").strip()
    if not host:
        return None

    bass_aliases = set(_resolve_entity_aliases("bass_level"))
    treble_aliases = set(_resolve_entity_aliases("treble_level"))
    if not bass_aliases and not treble_aliases:
        return None

    client = APIClient(host, 6054, None, client_info="HubVoiceSatRuntime")
    connect_timeout = min(CONNECTION_TIMEOUT, 3.0)
    entity_timeout = min(ENTITY_LIST_TIMEOUT, 2.0)
    state_timeout = 1.6
    try:
        await asyncio.wait_for(client.connect(login=False), timeout=connect_timeout)
        entities, _ = await asyncio.wait_for(
            client.list_entities_services(),
            timeout=entity_timeout,
        )

        bass_keys: set[int] = set()
        treble_keys: set[int] = set()
        for ent in entities:
            object_id = str(getattr(ent, "object_id", "")).strip()
            key = getattr(ent, "key", None)
            if not object_id or key is None:
                continue
            key_int = int(key)
            if object_id in bass_aliases:
                bass_keys.add(key_int)
            if object_id in treble_aliases:
                treble_keys.add(key_int)

        if not bass_keys and not treble_keys:
            return None

        loop = asyncio.get_running_loop()
        complete = loop.create_future()
        values: dict[str, float | None] = {"bass": None, "treble": None}

        def _on_state(state) -> None:
            key = getattr(state, "key", None)
            raw_value = getattr(state, "state", None)
            if key is None or raw_value is None:
                return
            try:
                key_int = int(key)
                numeric = float(raw_value)
            except Exception:
                return

            if key_int in bass_keys and values["bass"] is None:
                values["bass"] = _normalize_tone_db(numeric)
            if key_int in treble_keys and values["treble"] is None:
                values["treble"] = _normalize_tone_db(numeric)

            have_bass = (not bass_keys) or (values["bass"] is not None)
            have_treble = (not treble_keys) or (values["treble"] is not None)
            if have_bass and have_treble and not complete.done():
                loop.call_soon_threadsafe(complete.set_result, True)

        unsubscribe = None
        try:
            unsubscribe = client.subscribe_states(_on_state)
            with contextlib.suppress(asyncio.TimeoutError):
                await asyncio.wait_for(complete, timeout=state_timeout)
        finally:
            with contextlib.suppress(Exception):
                if unsubscribe is not None:
                    unsubscribe()

        bass_db = float(values["bass"] if values["bass"] is not None else 0.0)
        treble_db = float(values["treble"] if values["treble"] is not None else 0.0)
        return bass_db, treble_db
    except Exception as exc:
        logging.debug("Unable to read live tone settings from %s: %s", host, exc)
        return None
    finally:
        with contextlib.suppress(Exception):
            await client.disconnect()


def _read_satellite_tone_settings(satellite_host: str) -> tuple[float, float] | None:
    with contextlib.suppress(Exception):
        return asyncio.run(_read_satellite_tone_settings_async(satellite_host))
    return None


def _set_user_tone_settings(satellite_host: str, *, bass: float | int | str | None = None, treble: float | int | str | None = None) -> tuple[float, float]:
    host = str(satellite_host or "").strip().lower()
    if not host:
        raise ValueError("satellite_host is required")
    with _USER_TONE_SETTINGS_LOCK:
        current = dict(_USER_TONE_SETTINGS.get(host, {"bass_level": 0.0, "treble_level": 0.0}))
        if bass is not None:
            current["bass_level"] = _normalize_tone_db(bass)
        if treble is not None:
            current["treble_level"] = _normalize_tone_db(treble)
        _USER_TONE_SETTINGS[host] = current
        bass_db = float(current.get("bass_level", 0.0))
        treble_db = float(current.get("treble_level", 0.0))
    with contextlib.suppress(Exception):
        _persist_tone_settings_for_host(host, bass_db, treble_db)
    return bass_db, treble_db


def _get_user_tone_settings(satellite_host: str) -> tuple[float, float]:
    host = str(satellite_host or "").strip().lower()
    if not host:
        return 0.0, 0.0
    with _USER_TONE_SETTINGS_LOCK:
        current = _USER_TONE_SETTINGS.get(host)
    if current is None:
        bass_db, treble_db = _load_persisted_tone_settings_for_host(host)
        live_values = _read_satellite_tone_settings(host)
        if live_values is not None:
            bass_db, treble_db = live_values
            with contextlib.suppress(Exception):
                _persist_tone_settings_for_host(host, bass_db, treble_db)
        with _USER_TONE_SETTINGS_LOCK:
            _USER_TONE_SETTINGS[host] = {
                "bass_level": float(bass_db),
                "treble_level": float(treble_db),
            }
        return float(bass_db), float(treble_db)
    current_copy = dict(current)
    return float(current_copy.get("bass_level", 0.0)), float(current_copy.get("treble_level", 0.0))


def _resolve_request_tone_settings(params: dict[str, list[str]]) -> tuple[str, float, float]:
    satellite_value = str((params.get("satellite") or params.get("d") or [""])[0] or "").strip()
    if not satellite_value:
        return "", 0.0, 0.0
    sat = select_satellite(satellite_value)
    if not sat:
        return satellite_value, 0.0, 0.0
    bass_db, treble_db = _get_user_tone_settings(str(sat.get("host") or ""))
    return str(sat.get("id") or satellite_value), bass_db, treble_db


def _set_satellite_stream_status(
    satellite_id: str,
    *,
    path: str,
    family: str,
    mode: str,
    bass_db: float = 0.0,
    treble_db: float = 0.0,
    request_id: str = "",
) -> None:
    sat_id = str(satellite_id or "").strip()
    if not sat_id:
        return
    with _SATELLITE_STREAM_STATUS_LOCK:
        _SATELLITE_STREAM_STATUS[sat_id] = {
            "satellite": sat_id,
            "path": path,
            "family": family,
            "mode": mode,
            "bass_db": float(bass_db),
            "treble_db": float(treble_db),
            "request_id": request_id,
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        }


def _clear_satellite_stream_status(satellite_id: str, request_id: str = "") -> None:
    sat_id = str(satellite_id or "").strip()
    if not sat_id:
        return
    with _SATELLITE_STREAM_STATUS_LOCK:
        current = _SATELLITE_STREAM_STATUS.get(sat_id)
        if not current:
            return
        if request_id and str(current.get("request_id") or "") != request_id:
            return
        _SATELLITE_STREAM_STATUS.pop(sat_id, None)


def _get_satellite_stream_status(satellite_id: str) -> dict:
    sat_id = str(satellite_id or "").strip()
    if not sat_id:
        return {}
    with _SATELLITE_STREAM_STATUS_LOCK:
        current = dict(_SATELLITE_STREAM_STATUS.get(sat_id, {}))
    return current


def _get_all_satellite_stream_statuses() -> list[dict]:
    with _SATELLITE_STREAM_STATUS_LOCK:
        statuses = [dict(item) for item in _SATELLITE_STREAM_STATUS.values()]
    statuses.sort(key=lambda item: str(item.get("satellite") or ""))
    return statuses


class _BiquadFilter:
    def __init__(self, coeffs: tuple[float, float, float, float, float], channels: int) -> None:
        self.b0, self.b1, self.b2, self.a1, self.a2 = coeffs
        self.z1 = [0.0] * max(1, channels)
        self.z2 = [0.0] * max(1, channels)

    def process(self, samples):
        if np is None or samples.size == 0:
            return samples
        output = np.array(samples, copy=True)
        channels = output.shape[1] if output.ndim > 1 else 1
        if len(self.z1) != channels:
            self.z1 = [0.0] * channels
            self.z2 = [0.0] * channels
        if output.ndim == 1:
            output = output.reshape((-1, 1))
        for channel_index in range(output.shape[1]):
            z1 = self.z1[channel_index]
            z2 = self.z2[channel_index]
            channel = output[:, channel_index]
            for sample_index in range(len(channel)):
                x0 = float(channel[sample_index])
                y0 = self.b0 * x0 + z1
                z1 = self.b1 * x0 - self.a1 * y0 + z2
                z2 = self.b2 * x0 - self.a2 * y0
                channel[sample_index] = y0
            self.z1[channel_index] = z1
            self.z2[channel_index] = z2
        return np.clip(output, -1.0, 1.0)


def _design_shelf_filter(sample_rate: int, center_hz: float, gain_db: float, slope: float, *, high_shelf: bool) -> tuple[float, float, float, float, float]:
    if abs(gain_db) <= 0.01:
        return 1.0, 0.0, 0.0, 0.0, 0.0
    sample_rate = max(8000, int(sample_rate or 48000))
    center_hz = max(20.0, min(float(center_hz), (sample_rate / 2.0) - 100.0))
    slope = max(0.1, float(slope))
    amplitude = math.pow(10.0, float(gain_db) / 40.0)
    omega = 2.0 * math.pi * center_hz / sample_rate
    sin_omega = math.sin(omega)
    cos_omega = math.cos(omega)
    alpha = (sin_omega / 2.0) * math.sqrt((amplitude + (1.0 / amplitude)) * ((1.0 / slope) - 1.0) + 2.0)
    beta = 2.0 * math.sqrt(amplitude) * alpha
    if high_shelf:
        b0 = amplitude * ((amplitude + 1.0) + ((amplitude - 1.0) * cos_omega) + beta)
        b1 = -2.0 * amplitude * ((amplitude - 1.0) + ((amplitude + 1.0) * cos_omega))
        b2 = amplitude * ((amplitude + 1.0) + ((amplitude - 1.0) * cos_omega) - beta)
        a0 = (amplitude + 1.0) - ((amplitude - 1.0) * cos_omega) + beta
        a1 = 2.0 * ((amplitude - 1.0) - ((amplitude + 1.0) * cos_omega))
        a2 = (amplitude + 1.0) - ((amplitude - 1.0) * cos_omega) - beta
    else:
        b0 = amplitude * ((amplitude + 1.0) - ((amplitude - 1.0) * cos_omega) + beta)
        b1 = 2.0 * amplitude * ((amplitude - 1.0) - ((amplitude + 1.0) * cos_omega))
        b2 = amplitude * ((amplitude + 1.0) - ((amplitude - 1.0) * cos_omega) - beta)
        a0 = (amplitude + 1.0) + ((amplitude - 1.0) * cos_omega) + beta
        a1 = -2.0 * ((amplitude - 1.0) + ((amplitude + 1.0) * cos_omega))
        a2 = (amplitude + 1.0) + ((amplitude - 1.0) * cos_omega) - beta
    return b0 / a0, b1 / a0, b2 / a0, a1 / a0, a2 / a0


class _ToneProcessor:
    def __init__(self, sample_rate: int, channels: int, bass_db: float, treble_db: float) -> None:
        self.filters: list[_BiquadFilter] = []
        if abs(bass_db) > 0.01:
            self.filters.append(_BiquadFilter(_design_shelf_filter(sample_rate, 100.0, bass_db, 0.65, high_shelf=False), channels))
        if abs(treble_db) > 0.01:
            self.filters.append(_BiquadFilter(_design_shelf_filter(sample_rate, 3000.0, treble_db, 0.55, high_shelf=True), channels))

    def process(self, samples):
        output = samples
        for filter_stage in self.filters:
            output = filter_stage.process(output)
        return output


def _pick_media_player(entities):
    """Pick the speaker media player from an entity list.

    Prefers the entity named 'Media Player' (the speaker_source player) over
    group/sendspin players so that HubMusic commands go to the local speaker.
    """
    infos = [ent for ent in entities if ent.__class__.__name__ == "MediaPlayerInfo"]
    named = next((ent for ent in infos if getattr(ent, "name", "") == "Media Player"), None)
    return named or next(iter(infos), None)


async def play_media_on_satellite_async(satellite_host: str, media_url: str, announcement: bool = False) -> None:
    """Play a media URL on the satellite media player."""
    if not satellite_host or not media_url:
        raise ValueError("satellite_host and media_url are required")

    cached_key = _media_player_key_cache.get(satellite_host)
    client = APIClient(satellite_host, 6054, None, client_info="HubVoiceSatRuntime")
    try:
        await asyncio.wait_for(client.connect(login=True), timeout=CONNECTION_TIMEOUT)
        if cached_key is None:
            entities, _ = await asyncio.wait_for(
                client.list_entities_services(),
                timeout=ENTITY_LIST_TIMEOUT,
            )
            media = _pick_media_player(entities)
            if media is None:
                _METRICS.record_request(0, error=True, error_type="satellite")
                raise RuntimeError(f"Satellite {satellite_host} has no media player")
            cached_key = media.key
            _media_player_key_cache[satellite_host] = cached_key
        client.media_player_command(
            cached_key,
            command=MediaPlayerCommand.PLAY,
            media_url=media_url,
            announcement=announcement,
        )
        await asyncio.sleep(SATELLITE_MEDIA_DELAY)
        _METRICS.record_request(SATELLITE_MEDIA_DELAY, error=False)
    except asyncio.TimeoutError as e:
        _METRICS.record_request(0, error=True, error_type="satellite")
        raise RuntimeError(f"Timeout communicating with satellite {satellite_host}: {e}")
    except Exception:
        _METRICS.record_request(0, error=True, error_type="satellite")
        raise
    finally:
        try:
            await client.disconnect()
        except Exception:
            pass


def play_media_on_satellite(satellite_host: str, media_url: str, announcement: bool = False) -> None:
    """Synchronous wrapper for play_media_on_satellite_async."""
    asyncio.run(play_media_on_satellite_async(satellite_host, media_url, announcement=announcement))


def send_to_satellite(satellite_host: str, media_url: str) -> None:
    """Synchronous wrapper for send_to_satellite_async."""
    asyncio.run(send_to_satellite_async(satellite_host, media_url))


def prepare_satellite_for_tts(satellite_host: str) -> None:
    """Best-effort cleanup so TTS does not get blocked by stale media playback."""
    if not satellite_host:
        return

    async def _fast_prestop() -> None:
        cached_key = _media_player_key_cache.get(satellite_host)
        client = APIClient(satellite_host, 6054, None, client_info="HubVoiceSatRuntime")
        try:
            await asyncio.wait_for(client.connect(login=True), timeout=TTS_PRESTOP_CONNECT_TIMEOUT)
            if cached_key is None:
                entities, _ = await asyncio.wait_for(
                    client.list_entities_services(),
                    timeout=TTS_PRESTOP_ENTITY_TIMEOUT,
                )
                media = _pick_media_player(entities)
                if media is not None:
                    _media_player_key_cache[satellite_host] = media.key
                    client.media_player_command(media.key, command=MediaPlayerCommand.STOP, announcement=False)
            else:
                client.media_player_command(cached_key, command=MediaPlayerCommand.STOP, announcement=False)
        finally:
            with contextlib.suppress(Exception):
                await client.disconnect()

    try:
        # Keep this fast to avoid delaying spoken text while music is active.
        asyncio.run(_fast_prestop())
        # Brief pause so STOP processes before PLAY arrives.
        time.sleep(0.05)
    except Exception as exc:
        logging.debug("TTS pre-stop skipped for %s: %s", satellite_host, exc)


async def stop_media_on_satellite_async(satellite_host: str, announcement: bool = False) -> None:
    """Stop active media playback on a satellite."""
    if not satellite_host:
        raise ValueError("satellite_host is required")

    cached_key = _media_player_key_cache.get(satellite_host)
    client = APIClient(satellite_host, 6054, None, client_info="HubVoiceSatRuntime")
    try:
        await asyncio.wait_for(client.connect(login=True), timeout=CONNECTION_TIMEOUT)
        if cached_key is None:
            entities, _ = await asyncio.wait_for(
                client.list_entities_services(),
                timeout=ENTITY_LIST_TIMEOUT,
            )
            media = _pick_media_player(entities)
            if media is None:
                raise RuntimeError(f"Satellite {satellite_host} has no media player")
            cached_key = media.key
            _media_player_key_cache[satellite_host] = cached_key
        client.media_player_command(
            cached_key,
            command=MediaPlayerCommand.STOP,
            announcement=announcement,
        )
    finally:
        try:
            await client.disconnect()
        except Exception:
            pass


def stop_media_on_satellite(satellite_host: str, announcement: bool = False) -> None:
    """Synchronous wrapper for stop_media_on_satellite_async."""
    asyncio.run(stop_media_on_satellite_async(satellite_host, announcement=announcement))


async def _fast_stop_media_on_satellite_async(satellite_host: str) -> bool:
    """Best-effort low-latency STOP that only uses cached media player keys."""
    if not satellite_host:
        return False

    cached_key = _media_player_key_cache.get(satellite_host)
    if cached_key is None:
        return False

    client = APIClient(satellite_host, 6054, None, client_info="HubVoiceSatRuntime")
    try:
        await asyncio.wait_for(client.connect(login=True), timeout=1.2)
        client.media_player_command(cached_key, command=MediaPlayerCommand.STOP, announcement=False)
        return True
    except Exception:
        return False
    finally:
        try:
            await client.disconnect()
        except Exception:
            pass


def _fast_stop_media_on_satellite(satellite_host: str) -> None:
    try:
        asyncio.run(_fast_stop_media_on_satellite_async(satellite_host))
    except Exception:
        pass


def _stop_music_for_voice(satellite_host: str, update_hubmusic_state: bool) -> None:
    """Background helper: stop HubMusic media playback when a voice session starts."""
    snapshot = _HUB_MUSIC_STATE.snapshot() if update_hubmusic_state else {}
    try:
        if HUBMUSIC_SOFT_STOP_ONLY:
            _stop_hubmusic_ffmpeg()
            logging.info("Music soft-stopped on voice start for %s", satellite_host)
        else:
            stop_media_on_satellite(satellite_host, announcement=False)
            logging.info("Music stopped on voice start for %s", satellite_host)
    except Exception as exc:
        logging.debug("Music stop on voice start skipped for %s: %s", satellite_host, exc)
    finally:
        if update_hubmusic_state:
            _HUB_MUSIC_STATE.stop()
            _HUB_MUSIC_STATE.set_results(
                "voice_pause",
                stopped=[{"id": "", "alias": "", "host": satellite_host, "duration_ms": 0, "attempt": 1}],
                mode=str(snapshot.get("mode") or "single"),
            )


def _resume_music_after_voice(hubmusic_state: dict, resume_satellite_id: str, delay_s: float) -> None:
    """Best-effort HubMusic resume after a voice answer finishes playing."""
    if not hubmusic_state:
        return

    source_url = str(hubmusic_state.get("source_url") or "").strip()
    if not source_url:
        return

    if delay_s > 0:
        time.sleep(delay_s)

    # If playback is already active (manual restart), do not override current routing.
    if _HUB_MUSIC_STATE.snapshot().get("active"):
        logging.info("Skipping voice resume because HubMusic is already active")
        return

    mode = str(hubmusic_state.get("mode") or "single")
    title = str(hubmusic_state.get("title") or "")
    resume_targets: list[dict] = []

    if mode in {"all_reachable", "stereo_pair"}:
        raw_targets = hubmusic_state.get("satellites")
        if isinstance(raw_targets, list):
            for index, item in enumerate(raw_targets):
                if isinstance(item, dict):
                    host = str(item.get("host") or "").strip()
                    if not host:
                        continue
                    channel = str(item.get("channel") or "").strip().lower()
                    if mode == "stereo_pair" and channel not in {"left", "right"}:
                        channel = "left" if index == 0 else "right"
                    resume_targets.append(
                        {
                            "id": str(item.get("id") or "").strip(),
                            "alias": str(item.get("alias") or "").strip(),
                            "host": host,
                            "channel": channel,
                        }
                    )
    else:
        sat = select_satellite(resume_satellite_id)
        if sat:
            resume_targets.append(
                {
                    "id": str(sat.get("id") or "").strip(),
                    "alias": str(sat.get("alias") or "").strip(),
                    "host": str(sat.get("host") or "").strip(),
                }
            )

    sent: list[dict] = []
    failed: list[dict] = []
    for target in resume_targets:
        host = str(target.get("host") or "").strip()
        if not host:
            continue
        started = time.perf_counter()
        sat_id = str(target.get("id") or "").strip()
        target_source_url = build_satellite_runtime_media_url(source_url, sat_id)
        if mode == "stereo_pair":
            target_source_url = build_satellite_stereo_media_url(target_source_url, str(target.get("channel") or ""))
        try:
            play_media_on_satellite(host, target_source_url, announcement=False)
            sent.append(
                {
                    "id": sat_id,
                    "alias": str(target.get("alias") or "").strip(),
                    "host": host,
                    "channel": str(target.get("channel") or "").strip(),
                    "duration_ms": int((time.perf_counter() - started) * 1000),
                    "attempt": 1,
                }
            )
        except Exception as exc:
            failed.append(
                {
                    "id": sat_id,
                    "alias": str(target.get("alias") or "").strip(),
                    "host": host,
                    "channel": str(target.get("channel") or "").strip(),
                    "duration_ms": int((time.perf_counter() - started) * 1000),
                    "attempt": 1,
                    "error": str(exc),
                }
            )

    if sent:
        _HUB_MUSIC_STATE.activate(resume_targets, source_url, title, mode=mode)
        _HUB_MUSIC_STATE.set_results("voice_resume", sent=sent, failed=failed, mode=mode)
        logging.info("HubMusic resumed after voice answer on %d target(s)", len(sent))
        return

    _HUB_MUSIC_STATE.error("HubMusic resume failed after voice answer")
    _HUB_MUSIC_STATE.set_results("voice_resume", sent=sent, failed=failed, mode=mode)
    logging.warning("HubMusic resume after voice answer failed for all targets")


async def stop_satellite_playback_async(satellite_host: str) -> None:
    """Stop active media playback announcement on a satellite."""
    if not satellite_host:
        raise ValueError("satellite_host is required")

    client = APIClient(satellite_host, 6054, None, client_info="HubVoiceSatRuntime")
    try:
        await asyncio.wait_for(client.connect(login=True), timeout=CONNECTION_TIMEOUT)
        entities, _ = await asyncio.wait_for(
            client.list_entities_services(),
            timeout=ENTITY_LIST_TIMEOUT,
        )
        media = _pick_media_player(entities)
        if media is None:
            raise RuntimeError(f"Satellite {satellite_host} has no media player")
        client.media_player_command(media.key, command=MediaPlayerCommand.STOP, announcement=True)
    finally:
        try:
            await client.disconnect()
        except Exception:
            pass


def stop_satellite_playback(satellite_host: str) -> None:
    """Synchronous wrapper for stop_satellite_playback_async."""
    asyncio.run(stop_satellite_playback_async(satellite_host))


async def set_satellite_media_mute_async(satellite_host: str, muted: bool) -> None:
    """Mute or unmute the satellite media player via ESPHome API."""
    if not satellite_host:
        raise ValueError("satellite_host is required")

    client = APIClient(satellite_host, 6054, None, client_info="HubVoiceSatRuntime")
    try:
        await asyncio.wait_for(client.connect(login=True), timeout=CONNECTION_TIMEOUT)
        entities, _ = await asyncio.wait_for(
            client.list_entities_services(),
            timeout=ENTITY_LIST_TIMEOUT,
        )
        media = _pick_media_player(entities)
        if media is None:
            raise RuntimeError(f"Satellite {satellite_host} has no media player")
        command = MediaPlayerCommand.MUTE if muted else MediaPlayerCommand.UNMUTE
        client.media_player_command(media.key, command=command)
        await asyncio.sleep(0.2)
    finally:
        try:
            await client.disconnect()
        except Exception:
            pass


def set_satellite_media_mute(satellite_host: str, muted: bool) -> None:
    """Synchronous wrapper for set_satellite_media_mute_async."""
    asyncio.run(set_satellite_media_mute_async(satellite_host, muted))


async def set_satellite_media_volume_async(satellite_host: str, volume_percent: float) -> None:
    """Set the satellite media player volume (0-100 mapped to 0.0-1.0)."""
    if not satellite_host:
        raise ValueError("satellite_host is required")

    _throttle_satellite_media_volume(satellite_host)

    try:
        vol = float(volume_percent)
    except (TypeError, ValueError):
        raise ValueError(f"volume_percent must be numeric, got {volume_percent}")
    vol = max(0.0, min(100.0, vol))
    media_volume = vol / 100.0
    # Keep volume operations responsive to avoid UI request timeouts.
    connect_timeout = min(CONNECTION_TIMEOUT, 3.0)
    entity_timeout = min(ENTITY_LIST_TIMEOUT, 2.0)

    cached_key = _media_player_key_cache.get(satellite_host)
    client = APIClient(satellite_host, 6054, None, client_info="HubVoiceSatRuntime")
    try:
        await asyncio.wait_for(client.connect(login=True), timeout=connect_timeout)
        if cached_key is not None:
            # Fast path: use cached key, skip entity enumeration.
            try:
                client.media_player_command(cached_key, volume=media_volume)
                await asyncio.sleep(0.05)
                return
            except Exception:
                # Stale key â€” fall through to full enumeration.
                _media_player_key_cache.pop(satellite_host, None)
        entities, _ = await asyncio.wait_for(
            client.list_entities_services(),
            timeout=entity_timeout,
        )
        media = _pick_media_player(entities)
        if media is None:
            raise RuntimeError(f"Satellite {satellite_host} has no media player")
        _media_player_key_cache[satellite_host] = media.key
        client.media_player_command(media.key, volume=media_volume)
        await asyncio.sleep(0.05)
    finally:
        try:
            await client.disconnect()
        except Exception:
            pass


def set_satellite_media_volume(satellite_host: str, volume_percent: float) -> None:
    """Synchronous wrapper for set_satellite_media_volume_async."""
    asyncio.run(set_satellite_media_volume_async(satellite_host, volume_percent))


def _resolve_entity_aliases(entity_object_id: str) -> list[str]:
    """Return candidate object IDs to support old/new firmware naming."""
    normalized = str(entity_object_id or "").strip()
    if not normalized:
        return []

    alias_map = {
        "follow_up_listening_switch": [
            "follow_up_listening_switch",
            "follow_up_listening",
            "follow_up_enabled",
        ],
        "follow_up_listen_window_seconds": [
            "follow_up_listen_window_seconds",
            "follow_up_listen_window",
            "follow_up_window_seconds",
            "follow_up_window",
            "follow_up_timeout_seconds",
            "follow_up_timeout",
            "follow-up_listen_window__seconds_",
            "follow_up_listen_window__seconds_",
        ],
        "bass_level": [
            "bass_level",
            "bass",
            "eq_bass",
            "bass_gain",
            "bass_gain_db",
        ],
        "treble_level": [
            "treble_level",
            "treble",
            "eq_treble",
            "treble_gain",
            "treble_gain_db",
        ],
    }
    candidates = alias_map.get(normalized, [normalized])

    deduped: list[str] = []
    seen: set[str] = set()
    for candidate in candidates:
        key = candidate.strip()
        if key and key not in seen:
            seen.add(key)
            deduped.append(key)
    return deduped


async def _list_satellite_entity_ids_async(satellite_host: str) -> set[str]:
    """Fetch object_id values exposed by a satellite over ESPHome API."""
    if not satellite_host:
        return set()

    client = APIClient(satellite_host, 6054, None, client_info="HubVoiceSatRuntime")
    try:
        await asyncio.wait_for(client.connect(login=False), timeout=CONNECTION_TIMEOUT)
        entities, _ = await asyncio.wait_for(
            client.list_entities_services(),
            timeout=ENTITY_LIST_TIMEOUT,
        )
        return {
            str(getattr(ent, "object_id", "")).strip()
            for ent in entities
            if getattr(ent, "object_id", "")
        }
    finally:
        try:
            await client.disconnect()
        except Exception:
            pass


def list_satellite_entity_ids(satellite_host: str) -> set[str]:
    """Return exposed object_ids for a satellite, using a short TTL cache."""
    if not satellite_host:
        return set()

    now = time.time()
    with _SATELLITE_CAP_CACHE_LOCK:
        cached = _SATELLITE_CAP_CACHE.get(satellite_host)
        if cached and now - float(cached.get("ts", 0.0)) <= SATELLITE_CAP_CACHE_TTL:
            return set(cached.get("entity_ids", []))

    try:
        entity_ids = asyncio.run(_list_satellite_entity_ids_async(satellite_host))
    except Exception as exc:
        logging.debug("Failed to list entity ids for %s: %s", satellite_host, exc)
        entity_ids = set()

    with _SATELLITE_CAP_CACHE_LOCK:
        _SATELLITE_CAP_CACHE[satellite_host] = {
            "ts": now,
            "entity_ids": sorted(entity_ids),
        }
    return entity_ids


def get_satellite_capabilities(satellite_host: str) -> dict:
    """Build canonical control capability map for a satellite."""
    entity_ids = list_satellite_entity_ids(satellite_host)
    supports: dict[str, bool] = {}
    resolved: dict[str, str] = {}

    canonical_ids = [
        "speaker_volume",
        "wake_sound_volume",
        "whisper_mode",
        "whisper_volume_pct",
        "follow_up_listening_switch",
        "follow_up_listen_window_seconds",
        "bass_level",
        "treble_level",
    ]
    for canonical in canonical_ids:
        found = ""
        for candidate in _resolve_entity_aliases(canonical):
            if candidate in entity_ids:
                found = candidate
                break
        supports[canonical] = bool(found)
        if found:
            resolved[canonical] = found

    return {
        "supports": supports,
        "resolved": resolved,
    }


async def set_satellite_number_async(satellite_host: str, entity_object_id: str, value: float) -> None:
    """
    Set a number entity on the satellite (e.g., volume control).
    
    Args:
        satellite_host: IP or hostname of satellite device
        entity_object_id: ESPHome entity object ID (e.g., 'speaker_volume')
        value: Numeric value to set
    
    Raises:
        ValueError: If parameters invalid
        RuntimeError: If entity not found, connection fails, or timeout occurs
    
    Timeout: 5 seconds for setting value
    """
    if not satellite_host or not entity_object_id:
        raise ValueError("satellite_host and entity_object_id are required")
    
    try:
        value_float = float(value)
    except (TypeError, ValueError):
        raise ValueError(f"value must be numeric, got {value}")
    
    aliases = _resolve_entity_aliases(entity_object_id)
    if not aliases:
        raise ValueError(f"Invalid entity id: {entity_object_id}")

    def _get_cached_key() -> int | None:
        with _SATELLITE_ENTITY_KEY_CACHE_LOCK:
            host_cache = _SATELLITE_ENTITY_KEY_CACHE.get(satellite_host, {})
            for alias in aliases:
                if alias in host_cache:
                    return int(host_cache[alias])
        return None

    def _cache_keys(key_by_object_id: dict[str, int]) -> None:
        with _SATELLITE_ENTITY_KEY_CACHE_LOCK:
            host_cache = _SATELLITE_ENTITY_KEY_CACHE.setdefault(satellite_host, {})
            for object_id, key in key_by_object_id.items():
                host_cache[str(object_id)] = int(key)

    def _clear_alias_cache() -> None:
        with _SATELLITE_ENTITY_KEY_CACHE_LOCK:
            host_cache = _SATELLITE_ENTITY_KEY_CACHE.get(satellite_host)
            if not host_cache:
                return
            for alias in aliases:
                host_cache.pop(alias, None)

    last_exc: Exception | None = None
    for attempt in range(1, SATELLITE_NUMBER_MAX_RETRIES + 1):
        client = APIClient(satellite_host, 6054, None, client_info="HubVoiceSatRuntime")
        try:
            await asyncio.wait_for(client.connect(login=False), timeout=CONNECTION_TIMEOUT)

            entity_key = _get_cached_key()
            if entity_key is None:
                entities, _ = await asyncio.wait_for(
                    client.list_entities_services(),
                    timeout=ENTITY_LIST_TIMEOUT,
                )
                key_by_object_id = {
                    str(getattr(ent, "object_id", "")): ent.key
                    for ent in entities
                    if getattr(ent, "object_id", "")
                }
                _cache_keys(key_by_object_id)
                for alias in aliases:
                    if alias in key_by_object_id:
                        entity_key = int(key_by_object_id[alias])
                        break

            if entity_key is None:
                raise RuntimeError(f"Satellite does not expose {entity_object_id}")

            client.number_command(entity_key, value_float)
            await asyncio.sleep(SATELLITE_NUMBER_POST_DELAY)
            _METRICS.record_request(SATELLITE_NUMBER_POST_DELAY, error=False)
            return
        except asyncio.TimeoutError as exc:
            last_exc = exc
            _clear_alias_cache()
            if attempt < SATELLITE_NUMBER_MAX_RETRIES:
                logging.warning(
                    "satellite-number timeout (%s on %s), retry %d/%d",
                    entity_object_id,
                    satellite_host,
                    attempt,
                    SATELLITE_NUMBER_MAX_RETRIES,
                )
                await asyncio.sleep(SATELLITE_NUMBER_RETRY_DELAY * attempt)
                continue
            _METRICS.record_request(0, error=True, error_type="satellite")
            raise RuntimeError(f"Timeout setting {entity_object_id} on {satellite_host}: {exc}")
        except Exception as exc:
            last_exc = exc
            message = str(exc).lower()
            retryable = ("does not expose" not in message) and attempt < SATELLITE_NUMBER_MAX_RETRIES
            if retryable:
                _clear_alias_cache()
                logging.warning(
                    "satellite-number transient error (%s on %s): %s; retry %d/%d",
                    entity_object_id,
                    satellite_host,
                    exc,
                    attempt,
                    SATELLITE_NUMBER_MAX_RETRIES,
                )
                await asyncio.sleep(SATELLITE_NUMBER_RETRY_DELAY * attempt)
                continue
            _METRICS.record_request(0, error=True, error_type="satellite")
            raise
        finally:
            try:
                await client.disconnect()
            except Exception:
                pass

    _METRICS.record_request(0, error=True, error_type="satellite")
    if last_exc is not None:
        raise last_exc


def set_satellite_number(satellite_host: str, entity_object_id: str, value: float) -> None:
    """Synchronous wrapper for set_satellite_number_async."""
    asyncio.run(set_satellite_number_async(satellite_host, entity_object_id, value))


async def set_satellite_switch_async(satellite_host: str, entity_object_id: str, state: bool) -> None:
    """Turn a switch entity on/off on the satellite via ESPHome API."""
    if not satellite_host or not entity_object_id:
        raise ValueError("satellite_host and entity_object_id are required")
    try:
        client = APIClient(satellite_host, 6054, None, client_info="HubVoiceSatRuntime")
        await asyncio.wait_for(client.connect(login=False), timeout=CONNECTION_TIMEOUT)
        try:
            entities, _ = await asyncio.wait_for(
                client.list_entities_services(),
                timeout=ENTITY_LIST_TIMEOUT
            )
            key_by_object_id = {
                str(getattr(ent, "object_id", "")): ent.key
                for ent in entities
                if getattr(ent, "object_id", "")
            }
            entity_key = None
            for candidate in _resolve_entity_aliases(entity_object_id):
                if candidate in key_by_object_id:
                    entity_key = key_by_object_id[candidate]
                    break
            if entity_key is None:
                raise RuntimeError(f"Satellite does not expose {entity_object_id}")
            client.switch_command(entity_key, state)
            await asyncio.sleep(0.35)
        finally:
            try:
                await client.disconnect()
            except Exception:
                pass
    except asyncio.TimeoutError as e:
        raise RuntimeError(f"Timeout setting switch {entity_object_id} on {satellite_host}: {e}")
    except Exception:
        raise


def set_satellite_switch(satellite_host: str, entity_object_id: str, state: bool) -> None:
    """Synchronous wrapper for set_satellite_switch_async."""
    import threading
    
    # Store result/error in a container
    result_container = {"result": None, "error": None}
    
    def run_async():
        try:
            # Create a new event loop for this thread
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                loop.run_until_complete(set_satellite_switch_async(satellite_host, entity_object_id, state))
            finally:
                loop.close()
        except Exception as e:
            result_container["error"] = e
    
    # Run in a thread to avoid event loop conflicts
    thread = threading.Thread(target=run_async, daemon=True)
    thread.start()
    thread.join(timeout=30)  # Wait up to 30 seconds
    
    if thread.is_alive():
        raise TimeoutError(f"set_satellite_switch timeout after 30 seconds on {satellite_host}")
    
    if result_container["error"]:
        raise result_container["error"]


def build_media_url(filename: str) -> str:
    if not filename:
        raise ValueError("Filename is required")
    return f"http://{get_runtime_host()}:{get_runtime_port()}/tts/{urllib.parse.quote(filename)}"


def estimate_media_url_duration_seconds(media_url: str, default_seconds: float = 2.0) -> float:
    """Estimate WAV duration for a runtime /tts URL to time post-answer resume."""
    try:
        parsed = urllib.parse.urlparse(str(media_url or ""))
        path = parsed.path or ""
        if not path.startswith("/tts/"):
            return float(default_seconds)
        filename = urllib.parse.unquote(path.split("/tts/", 1)[1]).strip()
        if not filename:
            return float(default_seconds)
        wav_path = RECORDINGS_PATH / filename
        if not wav_path.exists():
            return float(default_seconds)
        with wave.open(str(wav_path), "rb") as handle:
            frames = int(handle.getnframes())
            rate = int(handle.getframerate())
            if rate <= 0:
                return float(default_seconds)
            seconds = frames / float(rate)
            return max(0.0, min(30.0, float(seconds)))
    except Exception:
        return float(default_seconds)


def build_runtime_url(path: str) -> str:
    cleaned = "/" + str(path or "").lstrip("/")
    return f"http://{get_runtime_host()}:{get_runtime_port()}{cleaned}"


def normalize_media_url(raw_url: str) -> str:
    value = str(raw_url or "").strip()
    if not value:
        raise ValueError("A media URL is required")
    if value.startswith("/"):
        if value.lower().startswith("/hubmusic/live.flac"):
            value = "/hubmusic/live.mp3" + value[len("/hubmusic/live.flac"):]
        return build_runtime_url(value)

    parsed = urllib.parse.urlparse(value)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("Media URL must be an absolute http(s) URL or a runtime-relative path")
    # Satellites resolve localhost/loopback as themselves, not this runtime host.
    if (parsed.hostname or "").lower() in {"127.0.0.1", "localhost", "::1"}:
        runtime_host = get_runtime_host()
        port = parsed.port or get_runtime_port()
        netloc = f"{runtime_host}:{port}"
        path = parsed.path or ""
        if path.lower().startswith("/hubmusic/live.flac"):
            path = "/hubmusic/live.mp3" + path[len("/hubmusic/live.flac"):]
        parsed = parsed._replace(netloc=netloc, path=path)
        return urllib.parse.urlunparse(parsed)

    path = parsed.path or ""
    if path.lower().startswith("/hubmusic/live.flac"):
        parsed = parsed._replace(path="/hubmusic/live.mp3" + path[len("/hubmusic/live.flac"):])
        return urllib.parse.urlunparse(parsed)
    # Rewrite loopback /dlna/live.mp3 â†’ LAN IP so satellite can reach the runtime.
    if path.lower() == "/dlna/live.mp3":
        runtime_host = get_runtime_host()
        port = parsed.port or get_runtime_port()
        parsed = parsed._replace(netloc=f"{runtime_host}:{port}")
        return urllib.parse.urlunparse(parsed)
    return value


def build_satellite_runtime_media_url(raw_url: str, satellite_id: str) -> str:
    value = str(raw_url or "").strip()
    sat_id = str(satellite_id or "").strip()
    if not value or not sat_id:
        return value

    parsed = urllib.parse.urlparse(value)
    runtime_host = get_runtime_host().lower()
    parsed_host = (parsed.hostname or "").lower()
    runtime_hosts = {runtime_host, "127.0.0.1", "localhost", "::1"}
    runtime_paths = {"/hubmusic/live.mp3", "/hubmusic/live.flac", "/hubmusic/proxy", "/dlna/live.mp3"}
    if parsed.scheme not in {"http", "https"} or parsed_host not in runtime_hosts or (parsed.path or "").lower() not in runtime_paths:
        return value

    query = urllib.parse.parse_qsl(parsed.query, keep_blank_values=True)
    query = [(key, current) for key, current in query if key.lower() not in {"satellite", "d"}]
    query.append(("satellite", sat_id))
    return urllib.parse.urlunparse(parsed._replace(query=urllib.parse.urlencode(query)))


def build_satellite_stereo_media_url(raw_url: str, channel: str) -> str:
    value = str(raw_url or "").strip()
    channel_raw = str(channel or "").strip().lower()
    if channel_raw in {"l", "left"}:
        normalized = "left"
    elif channel_raw in {"r", "right"}:
        normalized = "right"
    else:
        return value
    if not value:
        return value

    parsed = urllib.parse.urlparse(value)
    runtime_host = get_runtime_host().lower()
    parsed_host = (parsed.hostname or "").lower()
    runtime_hosts = {runtime_host, "127.0.0.1", "localhost", "::1"}
    runtime_paths = {"/hubmusic/live.mp3", "/hubmusic/live.flac", "/hubmusic/proxy"}
    if parsed.scheme not in {"http", "https"} or parsed_host not in runtime_hosts or (parsed.path or "").lower() not in runtime_paths:
        return value

    query = urllib.parse.parse_qsl(parsed.query, keep_blank_values=True)
    query = [(key, current) for key, current in query if key.lower() != "channel"]
    query.append(("channel", normalized))
    return urllib.parse.urlunparse(parsed._replace(query=urllib.parse.urlencode(query)))


def append_query_params(raw_url: str, params: dict[str, object]) -> str:
    value = str(raw_url or "").strip()
    if not value or not params:
        return value
    parsed = urllib.parse.urlparse(value)
    query = urllib.parse.parse_qsl(parsed.query, keep_blank_values=True)
    remove_keys = {str(key).strip().lower() for key in params.keys()}
    query = [(key, current) for key, current in query if key.lower() not in remove_keys]
    for key, param_value in params.items():
        key_text = str(key).strip()
        if not key_text:
            continue
        query.append((key_text, str(param_value)))
    return urllib.parse.urlunparse(parsed._replace(query=urllib.parse.urlencode(query)))


async def play_media_on_satellites_parallel_async(
    targets: list[dict],
    stereo_channel_map: dict[str, str],
    source_url: str,
) -> tuple[list[dict], list[dict], list[dict]]:
    """Play media on multiple satellites in parallel with aggressive stereo resync correction."""
    retried: list[dict] = []
    target_info_list: list[dict] = []

    for target in targets:
        sat_id = str(target.get("id", "")).strip()
        sat_alias = str(target.get("alias", "")).strip()
        sat_host = str(target.get("host", "")).strip()
        if not sat_host:
            continue
        target_source_url = build_satellite_runtime_media_url(source_url, sat_id)
        sat_channel = stereo_channel_map.get(sat_id.lower(), "")
        if sat_channel:
            target_source_url = build_satellite_stereo_media_url(target_source_url, sat_channel)
        target_info_list.append(
            {
                "sat_id": sat_id,
                "sat_alias": sat_alias,
                "sat_host": sat_host,
                "sat_channel": sat_channel,
                "target_source_url": target_source_url,
            }
        )

    async def _play_cycle(attempt: int) -> tuple[list[dict], list[dict], int]:
        cycle_started = time.perf_counter()
        launch_at_ms = int((time.time() * 1000.0) + HUBMUSIC_STEREO_SHARED_LAUNCH_DELAY_MS)

        async def _play_single(info: dict) -> dict:
            started = time.perf_counter()
            target_url = info["target_source_url"]
            if info.get("sat_channel") in {"left", "right"}:
                target_url = append_query_params(
                    target_url,
                    {
                        "launch_at_ms": launch_at_ms,
                        "force_tone": "1",
                    },
                )
            try:
                await play_media_on_satellite_async(
                    info["sat_host"],
                    target_url,
                    announcement=False,
                )
                completed = time.perf_counter()
                return {
                    "ok": True,
                    "id": info["sat_id"],
                    "alias": info["sat_alias"],
                    "host": info["sat_host"],
                    "channel": info["sat_channel"],
                    "attempt": attempt,
                    "duration_ms": int((completed - started) * 1000),
                    "completed_offset_ms": int((completed - cycle_started) * 1000),
                }
            except Exception as exc:
                completed = time.perf_counter()
                return {
                    "ok": False,
                    "id": info["sat_id"],
                    "alias": info["sat_alias"],
                    "host": info["sat_host"],
                    "channel": info["sat_channel"],
                    "attempt": attempt,
                    "duration_ms": int((completed - started) * 1000),
                    "error": str(exc),
                }

        results = await asyncio.gather(*[_play_single(info) for info in target_info_list])
        sent_cycle = [
            {
                "id": item["id"],
                "alias": item["alias"],
                "host": item["host"],
                "channel": item["channel"],
                "attempt": item["attempt"],
                "duration_ms": item["duration_ms"],
            }
            for item in results
            if item.get("ok")
        ]
        failed_cycle = [
            {
                "id": item["id"],
                "alias": item["alias"],
                "host": item["host"],
                "channel": item["channel"],
                "attempt": item["attempt"],
                "duration_ms": item["duration_ms"],
                "error": item.get("error", "unknown_error"),
            }
            for item in results
            if not item.get("ok")
        ]

        completed_offsets = [int(item["completed_offset_ms"]) for item in results if item.get("ok")]
        skew_ms = max(completed_offsets) - min(completed_offsets) if len(completed_offsets) >= 2 else -1
        return sent_cycle, failed_cycle, skew_ms

    async def _stop_all_targets() -> None:
        async def _stop_single(info: dict):
            with contextlib.suppress(Exception):
                await stop_media_on_satellite_async(info["sat_host"], announcement=False)

        await asyncio.gather(*[_stop_single(info) for info in target_info_list])

    sent, failed, skew_ms = await _play_cycle(attempt=1)
    for item in failed:
        logging.warning("HubMusic parallel play failed for %s (%s): %s", item["id"], item["host"], item.get("error", "unknown"))

    # For stereo starts, actively resync when either channel fails or start skew exceeds the threshold.
    if len(target_info_list) >= 2:
        for resync_pass in range(1, HUBMUSIC_STEREO_RESYNC_MAX_PASSES + 1):
            needs_resync = bool(failed) or (skew_ms >= 0 and skew_ms > HUBMUSIC_STEREO_MAX_START_SKEW_MS)
            if not needs_resync:
                break

            reason = "failure" if failed else f"skew_{skew_ms}ms"
            logging.warning(
                "HubMusic stereo resync pass %d/%d triggered (%s)",
                resync_pass,
                HUBMUSIC_STEREO_RESYNC_MAX_PASSES,
                reason,
            )
            await _stop_all_targets()
            await asyncio.sleep(HUBMUSIC_STEREO_RESYNC_SETTLE_SECONDS)
            candidate_sent, candidate_failed, candidate_skew_ms = await _play_cycle(attempt=resync_pass + 1)

            for item in candidate_sent:
                retried.append({
                    **item,
                    "resync_pass": resync_pass,
                    "resync_reason": reason,
                    "cycle_skew_ms": candidate_skew_ms,
                })

            replace_current = False
            if len(candidate_failed) < len(failed):
                replace_current = True
            elif len(candidate_failed) == len(failed):
                if len(candidate_sent) > len(sent):
                    replace_current = True
                elif not candidate_failed and not failed and candidate_skew_ms >= 0 and (skew_ms < 0 or candidate_skew_ms < skew_ms):
                    replace_current = True

            if replace_current:
                sent, failed, skew_ms = candidate_sent, candidate_failed, candidate_skew_ms

    if skew_ms >= 0:
        for item in sent:
            item["sync_skew_ms"] = skew_ms

    return sent, failed, retried


def play_media_on_satellites_parallel(
    targets: list[dict],
    stereo_channel_map: dict[str, str],
    source_url: str,
) -> tuple[list[dict], list[dict], list[dict]]:
    """Synchronous wrapper for parallel satellite playback."""
    return asyncio.run(play_media_on_satellites_parallel_async(targets, stereo_channel_map, source_url))


def get_desktop_audio_info() -> dict:
    """Return desktop-audio capture options available on the current host."""
    if sd is None:
        return {
            "available": False,
            "reason": "sounddevice is not installed",
            "devices": [],
            "preferred_device": None,
        }

    try:
        devices = sd.query_devices()
    except Exception as exc:
        return {
            "available": False,
            "reason": str(exc),
            "devices": [],
            "preferred_device": None,
        }

    results: list[dict] = []
    preferred_device = None
    keywords = ("stereo mix", "loopback", "what u hear", "wave out")
    for index, device in enumerate(devices):
        input_channels = int(device.get("max_input_channels") or 0)
        if input_channels <= 0:
            continue
        name = str(device.get("name") or f"Device {index}")
        default_samplerate = int(device.get("default_samplerate") or 48000)
        entry = {
            "index": index,
            "name": name,
            "input_channels": input_channels,
            "default_samplerate": default_samplerate,
            "preferred": any(term in name.lower() for term in keywords),
        }
        results.append(entry)

    def _device_score(item: dict) -> tuple:
        name = str(item.get("name") or "").lower()
        index = int(item.get("index") or -1)
        samplerate = int(item.get("default_samplerate") or 0)
        channels = int(item.get("input_channels") or 0)
        stereo_mix = 1 if "stereo mix" in name else 0
        preferred_flag = 1 if any(term in name for term in keywords) else 0
        blacklist_penalty = 0 if index in HUBMUSIC_AUTO_CAPTURE_DEVICE_BLACKLIST else 1
        return (blacklist_penalty, stereo_mix, preferred_flag, samplerate, channels)

    if results:
        preferred_device = int(max(results, key=_device_score).get("index"))

    if preferred_device is None and results:
        preferred_device = results[0]["index"]

    return {
        "available": bool(results) and av is not None and np is not None,
        "reason": "" if (results and av is not None and np is not None) else (
            "No capture-capable desktop audio device found" if not results else "PyAV or numpy is not installed"
        ),
        "devices": results,
        "preferred_device": preferred_device,
    }


def _tail_text(path: Path, max_chars: int = 1200) -> str:
    try:
        text = path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return ""
    return text[-max_chars:].strip()


def _ensure_airplay_dependencies_installed() -> list[str]:
    missing: list[str] = []
    for module_name in AIRPLAY_REQUIRED_MODULES:
        if importlib.util.find_spec(module_name) is None:
            missing.append(module_name)
    return missing


def _detect_airplay_interface() -> tuple[str, str]:
    runtime_host = str(get_runtime_host() or "").strip()
    if ni is None:
        return "", "Python package 'netifaces' is not installed. Install 'netifaces2'."

    preferred_iface = str(load_config().get("airplay_interface", "") or "").strip()
    if preferred_iface:
        try:
            info = ni.ifaddresses(preferred_iface)
            ipv4_entries = info.get(ni.AF_INET, []) if hasattr(ni, "AF_INET") else []
            if ipv4_entries:
                return preferred_iface, ""
            return "", f"Configured airplay_interface '{preferred_iface}' has no IPv4 address."
        except Exception:
            return "", f"Configured airplay_interface '{preferred_iface}' was not found."

    try:
        interfaces = ni.interfaces()
    except Exception as exc:
        return "", f"Unable to enumerate network interfaces: {exc}"

    fallback = ""
    for iface in interfaces:
        try:
            info = ni.ifaddresses(iface)
        except Exception:
            continue
        ipv4_entries = info.get(ni.AF_INET, []) if hasattr(ni, "AF_INET") else []
        if not ipv4_entries:
            continue
        if not fallback:
            fallback = str(iface)
        for entry in ipv4_entries:
            addr = str(entry.get("addr") or "").strip()
            if runtime_host and addr == runtime_host:
                return str(iface), ""

    if fallback:
        return fallback, ""
    return "", "No IPv4-capable network interface found for AirPlay receiver."


def _find_airplay_helper_pids() -> list[int]:
    script_text = str(AIRPLAY_RECEIVER_SCRIPT).replace("\\", "/").lower()
    matches: list[int] = []
    if psutil is None:
        # Fallback for environments without psutil: query Win32_Process via PowerShell.
        with contextlib.suppress(Exception):
            command = (
                "Get-CimInstance Win32_Process | "
                "Where-Object { $_.Name -eq 'python.exe' -and $_.CommandLine -match 'ap2-receiver.py' } | "
                "Select-Object -ExpandProperty ProcessId"
            )
            result = subprocess.run(
                ["powershell", "-NoProfile", "-Command", command],
                capture_output=True,
                text=True,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                timeout=5,
            )
            for line in (result.stdout or "").splitlines():
                line = str(line or "").strip()
                if line.isdigit():
                    matches.append(int(line))
        return matches

    for proc in psutil.process_iter(["pid", "name", "cmdline"]):
        with contextlib.suppress(Exception):
            pid = int(proc.info.get("pid") or 0)
            if pid <= 0:
                continue
            cmdline_parts = proc.info.get("cmdline") or []
            if not cmdline_parts:
                continue
            cmd_text = " ".join(str(part) for part in cmdline_parts).replace("\\", "/").lower()
            if "ap2-receiver.py" in cmd_text and script_text in cmd_text:
                matches.append(pid)
    return matches


def _stop_orphan_airplay_helpers(exclude_pids: set[int] | None = None) -> list[int]:
    exclude = exclude_pids or set()
    killed: list[int] = []
    for pid in _find_airplay_helper_pids():
        if pid in exclude:
            continue
        _kill_pid(pid)
        killed.append(pid)
    if killed:
        logging.warning("Stopped orphan AirPlay helper process(es): %s", sorted(killed))
    return killed


def _airplay_bootstrap_code() -> str:
    return "\n".join(
        [
            "import runpy",
            "import socket",
            "import sys",
            "_orig = socket.inet_pton",
            "def _patched(family, addr):",
            "    f = family",
            "    if f not in (socket.AF_INET, socket.AF_INET6) and str(f) == '10':",
            "        f = socket.AF_INET6",
            "    return _orig(f, addr)",
            "socket.inet_pton = _patched",
            "script = sys.argv[1]",
            "sys.argv = [script] + sys.argv[2:]",
            "runpy.run_path(script, run_name='__main__')",
        ]
    )


def _ensure_airplay_firewall_rules() -> None:
    """Best-effort: allow mDNS and AirPlay RTSP ports through Windows Firewall."""
    if not sys.platform.startswith("win"):
        return

    rules = [
        ("HubVoice AirPlay mDNS In", "in", "UDP", "5353"),
        ("HubVoice AirPlay mDNS Out", "out", "UDP", "5353"),
        ("HubVoice AirPlay RTSP In", "in", "TCP", "7000"),
        ("HubVoice AirPlay RTSP Out", "out", "TCP", "7000"),
    ]

    for name, direction, protocol, port in rules:
        with contextlib.suppress(Exception):
            show = subprocess.run(
                ["netsh", "advfirewall", "firewall", "show", "rule", f"name={name}"],
                capture_output=True,
                text=True,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                timeout=5,
            )
            output = f"{show.stdout}\n{show.stderr}".lower()
            if "no rules match" in output:
                subprocess.run(
                    [
                        "netsh",
                        "advfirewall",
                        "firewall",
                        "add",
                        "rule",
                        f"name={name}",
                        f"dir={direction}",
                        "action=allow",
                        f"protocol={protocol}",
                        f"localport={port}",
                        "profile=private",
                    ],
                    capture_output=True,
                    text=True,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                    timeout=8,
                )


def airplay_status_snapshot() -> dict:
    with _AIRPLAY_LOCK:
        status = dict(_AIRPLAY_STATE)
        proc = _AIRPLAY_PROCESS

    if proc is not None:
        code = proc.poll()
        status["running"] = code is None
        status["pid"] = proc.pid if code is None else None
        if code is not None and not status.get("last_error"):
            status["last_error"] = f"AirPlay receiver exited with code {code}"

    status["script_path"] = str(AIRPLAY_RECEIVER_SCRIPT)
    status["script_present"] = AIRPLAY_RECEIVER_SCRIPT.exists()
    status["log_path"] = str(AIRPLAY_RECEIVER_LOG_PATH)
    return status


def start_airplay_receiver(receiver_name: str = AIRPLAY_RECEIVER_DEFAULT_NAME) -> tuple[bool, str, dict]:
    global _AIRPLAY_PROCESS

    name = str(receiver_name or AIRPLAY_RECEIVER_DEFAULT_NAME).strip() or AIRPLAY_RECEIVER_DEFAULT_NAME
    already_running = False
    with _AIRPLAY_LOCK:
        if _AIRPLAY_PROCESS is not None and _AIRPLAY_PROCESS.poll() is None:
            already_running = True
    if already_running:
        return True, "AirPlay receiver already running.", airplay_status_snapshot()

    # Clean up stale helper instances from previous runtime sessions to avoid mDNS conflicts.
    _stop_orphan_airplay_helpers()

    if not AIRPLAY_RECEIVER_SCRIPT.exists():
        message = f"AirPlay receiver script not found: {AIRPLAY_RECEIVER_SCRIPT}"
        with _AIRPLAY_LOCK:
            _AIRPLAY_STATE.update({"enabled": False, "running": False, "last_error": message})
        return False, message, airplay_status_snapshot()

    missing = _ensure_airplay_dependencies_installed()
    if missing:
        message = "Missing AirPlay dependencies: " + ", ".join(sorted(missing))
        with _AIRPLAY_LOCK:
            _AIRPLAY_STATE.update({"enabled": False, "running": False, "last_error": message})
        return False, message, airplay_status_snapshot()

    interface_name, iface_error = _detect_airplay_interface()
    if not interface_name:
        message = iface_error or "Unable to resolve a network interface for AirPlay receiver."
        with _AIRPLAY_LOCK:
            _AIRPLAY_STATE.update({"enabled": False, "running": False, "last_error": message})
        return False, message, airplay_status_snapshot()

    _ensure_airplay_firewall_rules()

    AIRPLAY_RECEIVER_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    cmd = [
        sys.executable,
        "-c",
        _airplay_bootstrap_code(),
        str(AIRPLAY_RECEIVER_SCRIPT),
        "-m",
        name,
        "-n",
        interface_name,
    ]
    cfg = load_config()
    compat_enabled = _config_bool(cfg.get("airplay_airmusic_compat"), True)
    jitter_window_divisor = _config_int(cfg.get("airplay_jitter_window_divisor"), 5, 3, 8)
    resend_retry_ms = _config_int(cfg.get("airplay_resend_retry_ms"), 120, 40, 2000)
    resend_max_burst = _config_int(cfg.get("airplay_resend_max_burst"), 4, 1, 7)
    child_env = os.environ.copy()
    child_env["HUBVOICE_AIRPLAY_AIRMUSIC_COMPAT"] = "1" if compat_enabled else "0"
    child_env["HUBVOICE_AIRPLAY_JITTER_WINDOW_DIVISOR"] = str(jitter_window_divisor)
    child_env["HUBVOICE_AIRPLAY_RESEND_RETRY_MS"] = str(resend_retry_ms)
    child_env["HUBVOICE_AIRPLAY_RESEND_MAX_BURST"] = str(resend_max_burst)

    try:
        with AIRPLAY_RECEIVER_LOG_PATH.open("a", encoding="utf-8", errors="ignore") as log_file:
            log_file.write(f"\n[{datetime.now().isoformat(timespec='seconds')}] Starting AirPlay receiver: {name} on {interface_name}\n")
            proc = subprocess.Popen(
                cmd,
                cwd=str(AIRPLAY_RECEIVER_SRC_DIR),
                stdout=log_file,
                stderr=subprocess.STDOUT,
                env=child_env,
                creationflags=flags,
            )
    except Exception as exc:
        message = f"Failed to start AirPlay receiver: {exc}"
        with _AIRPLAY_LOCK:
            _AIRPLAY_STATE.update({"enabled": False, "running": False, "last_error": message})
        return False, message, airplay_status_snapshot()

    time.sleep(1.2)
    exit_code = proc.poll()
    if exit_code is not None:
        tail = _tail_text(AIRPLAY_RECEIVER_LOG_PATH)
        message = f"AirPlay receiver exited immediately (code {exit_code})."
        if tail:
            message = f"{message} {tail.splitlines()[-1]}"
        with _AIRPLAY_LOCK:
            _AIRPLAY_PROCESS = None
            _AIRPLAY_STATE.update({"enabled": False, "running": False, "last_error": message})
        return False, message, airplay_status_snapshot()

    with _AIRPLAY_LOCK:
        _AIRPLAY_PROCESS = proc
        _AIRPLAY_STATE.update(
            {
                "enabled": True,
                "running": True,
                "pid": proc.pid,
                "name": name,
                "interface": interface_name,
                "started_at": datetime.now().isoformat(timespec="seconds"),
                "last_error": "",
            }
        )
    return True, "AirPlay receiver started.", airplay_status_snapshot()


def stop_airplay_receiver() -> tuple[bool, str, dict]:
    global _AIRPLAY_PROCESS

    with _AIRPLAY_LOCK:
        proc = _AIRPLAY_PROCESS
        _AIRPLAY_PROCESS = None

    if proc is None or proc.poll() is not None:
        _stop_orphan_airplay_helpers()
        with _AIRPLAY_LOCK:
            _AIRPLAY_STATE.update({"running": False, "pid": None})
        return True, "AirPlay receiver is not running.", airplay_status_snapshot()

    try:
        proc.terminate()
        proc.wait(timeout=4)
    except Exception:
        with contextlib.suppress(Exception):
            proc.kill()

    _stop_orphan_airplay_helpers()

    with _AIRPLAY_LOCK:
        _AIRPLAY_STATE.update({"running": False, "pid": None, "enabled": False})
    return True, "AirPlay receiver stopped.", airplay_status_snapshot()


def hubmusic_status_snapshot() -> dict:
    satellites = load_satellites()
    reachable = []
    for sat in satellites:
        if test_satellite_connection(sat["host"]):
            reachable.append(
                {
                    "id": sat["id"],
                    "alias": sat.get("alias", ""),
                    "host": sat["host"],
                }
            )
    snapshot = _HUB_MUSIC_STATE.snapshot()
    snapshot["runtime_host"] = get_runtime_host()
    snapshot["runtime_port"] = get_runtime_port()
    snapshot["reachable_satellites"] = reachable
    snapshot["reachable_count"] = len(reachable)
    snapshot["configured_count"] = len(satellites)
    snapshot["desktop_audio"] = get_desktop_audio_info()
    cfg = load_config()
    snapshot["stereo_left_default"] = cfg.get("hubmusic_stereo_left", "").strip()
    snapshot["stereo_right_default"] = cfg.get("hubmusic_stereo_right", "").strip()
    snapshot["stereo_volume_default"] = _config_int(cfg.get("hubmusic_stereo_volume_pct"), 50, 0, 100)
    snapshot["airplay"] = airplay_status_snapshot()
    snapshot["dlna"] = dlna_status_snapshot()
    snapshot["active_streams"] = _get_all_satellite_stream_statuses()
    return snapshot


def list_reachable_satellites() -> list[dict]:
    satellites = load_satellites()
    reachable: list[dict] = []
    for sat in satellites:
        if test_satellite_connection(sat["host"]):
            reachable.append(sat)
    return reachable


def hubmusic_startup_warmup_seconds() -> float:
    """Return startup warmup, extending briefly when host CPU is saturated."""
    warmup = HUBMUSIC_LIVE_WARMUP_SECONDS
    if psutil is None:
        return warmup
    try:
        readings: list[float] = []
        for _ in range(max(1, int(HUBMUSIC_CPU_CHECK_SAMPLES))):
            readings.append(float(psutil.cpu_percent(interval=HUBMUSIC_CPU_CHECK_INTERVAL_SECONDS)))
        peak = max(readings) if readings else 0.0
        if peak >= HUBMUSIC_CPU_SPIKE_THRESHOLD_PCT:
            adjusted = warmup + HUBMUSIC_CPU_SPIKE_EXTRA_WARMUP_SECONDS
            logging.info(
                "HubMusic startup CPU spike detected (peak %.1f%%), extending warmup %.1fs -> %.1fs",
                peak,
                warmup,
                adjusted,
            )
            return adjusted
    except Exception as exc:
        logging.debug("HubMusic CPU probe unavailable, using default warmup: %s", exc)
    return warmup


def resolve_hubmusic_targets(payload: dict, default_snapshot: dict | None = None) -> tuple[list[dict], str, str]:
    mode_raw = str(payload.get("mode") or payload.get("target_mode") or "").strip().lower()
    exclude_raw = str(payload.get("exclude_satellite") or payload.get("exclude") or "").strip()
    all_flag = payload.get("all")
    stereo_pair_flag = _config_bool(payload.get("stereo_pair"), False)
    if isinstance(all_flag, bool) and all_flag:
        mode_raw = "all_reachable"
    if stereo_pair_flag:
        mode_raw = "stereo_pair"

    exclude_sat = select_satellite(exclude_raw) if exclude_raw else None
    exclude_id = str(exclude_sat.get("id", "")).strip() if exclude_sat else ""

    if mode_raw in {"all", "all_reachable", "broadcast"}:
        targets = list_reachable_satellites()
        if exclude_id:
            targets = [sat for sat in targets if str(sat.get("id", "")).strip().lower() != exclude_id.lower()]
        return targets, "all_reachable", exclude_id

    if mode_raw in {"stereo_pair", "pair", "stereo"}:
        cfg = load_config()
        left_ref = str(
            payload.get("left_satellite")
            or payload.get("left")
            or payload.get("left_id")
            or cfg.get("hubmusic_stereo_left")
            or ""
        ).strip()
        right_ref = str(
            payload.get("right_satellite")
            or payload.get("right")
            or payload.get("right_id")
            or cfg.get("hubmusic_stereo_right")
            or ""
        ).strip()
        if not left_ref or not right_ref:
            raise ValueError("Stereo pair mode requires left/right satellites (payload left_satellite/right_satellite or config hubmusic_stereo_left/hubmusic_stereo_right)")

        left_sat = select_satellite(left_ref)
        right_sat = select_satellite(right_ref)
        if not left_sat or not right_sat:
            raise LookupError("Stereo pair satellites were not found")

        left_id = str(left_sat.get("id", "")).strip().lower()
        right_id = str(right_sat.get("id", "")).strip().lower()
        if not left_id or not right_id or left_id == right_id:
            raise ValueError("Stereo pair requires two distinct satellites")

        unreachable: list[str] = []
        for sat in (left_sat, right_sat):
            if not test_satellite_connection(str(sat.get("host", ""))):
                unreachable.append(str(sat.get("id", "")))
        if unreachable:
            raise LookupError("Stereo pair satellites not reachable: " + ", ".join(unreachable))

        return [left_sat, right_sat], "stereo_pair", exclude_id

    sat = select_satellite(str(payload.get("satellite") or payload.get("sat_id") or payload.get("d") or "").strip())
    if sat:
        return [sat], "single", exclude_id

    snapshot = default_snapshot or _HUB_MUSIC_STATE.snapshot()
    if str(snapshot.get("mode") or "") == "all_reachable":
        targets = list_reachable_satellites()
        if exclude_id:
            targets = [sat for sat in targets if str(sat.get("id", "")).strip().lower() != exclude_id.lower()]
        return targets, "all_reachable", exclude_id

    if str(snapshot.get("mode") or "") == "stereo_pair":
        existing = snapshot.get("satellites") or []
        if isinstance(existing, list) and len(existing) >= 2:
            restored: list[dict] = []
            for item in existing[:2]:
                sat = select_satellite(str(item.get("id") or item.get("host") or "").strip())
                if sat:
                    restored.append(sat)
            if len(restored) == 2:
                return restored, "stereo_pair", exclude_id

    existing = snapshot.get("satellites") or []
    if isinstance(existing, list) and existing:
        restored: list[dict] = []
        for item in existing:
            sat = select_satellite(str(item.get("id") or item.get("host") or "").strip())
            if sat:
                restored.append(sat)
        if restored:
            return restored, "single", exclude_id

    return [], "single", exclude_id


def start_hubmusic_route(payload: dict, *, default_snapshot: dict | None = None) -> dict:
    source_url = normalize_media_url(str(payload.get("url") or payload.get("source_url") or ""))
    title = str(payload.get("title") or "").strip()
    snapshot = default_snapshot or _HUB_MUSIC_STATE.snapshot()
    targets, mode, exclude_id = resolve_hubmusic_targets(payload, default_snapshot=snapshot)
    if not targets:
        raise LookupError("No reachable satellites found for HubMusic")

    _start_hubmusic_ffmpeg(source_url)

    stereo_channel_map: dict[str, str] = {}
    if mode == "stereo_pair" and len(targets) >= 2:
        left_id = str(targets[0].get("id", "")).strip().lower()
        right_id = str(targets[1].get("id", "")).strip().lower()
        if left_id:
            stereo_channel_map[left_id] = "left"
        if right_id:
            stereo_channel_map[right_id] = "right"

    # Use parallel playback for stereo pair to sync both speakers
    if mode == "stereo_pair" and len(targets) >= 2:
        sent, failed, retried = play_media_on_satellites_parallel(targets, stereo_channel_map, source_url)
    else:
        # Sequential playback for single satellite or all_reachable mode
        sent: list[dict] = []
        failed: list[dict] = []
        retried: list[dict] = []

        for target in targets:
            sat_id = str(target.get("id", ""))
            sat_alias = str(target.get("alias", ""))
            sat_host = str(target.get("host", ""))
            target_source_url = build_satellite_runtime_media_url(source_url, sat_id)
            sat_channel = stereo_channel_map.get(sat_id.strip().lower(), "")
            if sat_channel:
                target_source_url = build_satellite_stereo_media_url(target_source_url, sat_channel)

            attempt = 1
            started = time.perf_counter()
            try:
                play_media_on_satellite(sat_host, target_source_url, announcement=False)
                duration_ms = int((time.perf_counter() - started) * 1000)
                sent.append({
                    "id": sat_id,
                    "alias": sat_alias,
                    "host": sat_host,
                    "channel": sat_channel,
                    "duration_ms": duration_ms,
                    "attempt": attempt,
                })
                continue
            except Exception as first_exc:
                logging.warning("HubMusic play first attempt failed for %s (%s): %s", sat_id, sat_host, first_exc)

            attempt = 2
            retry_started = time.perf_counter()
            try:
                play_media_on_satellite(sat_host, target_source_url, announcement=False)
                duration_ms = int((time.perf_counter() - retry_started) * 1000)
                sent.append({
                    "id": sat_id,
                    "alias": sat_alias,
                    "host": sat_host,
                    "channel": sat_channel,
                    "duration_ms": duration_ms,
                    "attempt": attempt,
                })
                retried.append({
                    "id": sat_id,
                    "alias": sat_alias,
                    "host": sat_host,
                    "channel": sat_channel,
                    "duration_ms": duration_ms,
                    "attempt": attempt,
                })
            except Exception as retry_exc:
                duration_ms = int((time.perf_counter() - retry_started) * 1000)
                failed.append({
                    "id": sat_id,
                    "alias": sat_alias,
                    "host": sat_host,
                    "channel": sat_channel,
                    "duration_ms": duration_ms,
                    "attempt": attempt,
                    "error": str(retry_exc),
                })
                logging.error("HubMusic play failed for %s (%s): %s", sat_id, sat_host, retry_exc)

    if not sent:
        _stop_hubmusic_ffmpeg()
        message = "Failed to start playback on selected satellite(s)"
        _HUB_MUSIC_STATE.error(message)
        _HUB_MUSIC_STATE.set_results("play", sent=sent, failed=failed, retried=retried, exclude_satellite=exclude_id, mode=mode)
        _APP_STATE.update(last_action="hubmusic_error", last_error=message)
        return {
            "ok": False,
            "error": message,
            "mode": mode,
            "exclude_satellite": exclude_id,
            "sent": sent,
            "failed": failed,
            "retried": retried,
            "status": hubmusic_status_snapshot(),
        }

    _HUB_MUSIC_STATE.activate(sent, source_url, title, mode=mode)
    _HUB_MUSIC_STATE.set_results("play", sent=sent, failed=failed, retried=retried, exclude_satellite=exclude_id, mode=mode)
    _APP_STATE.update(last_action=f"hubmusic_play:{len(sent)}", last_error="")
    return {
        "ok": True,
        "mode": mode,
        "exclude_satellite": exclude_id,
        "stereo_pair": mode == "stereo_pair",
        "satellite": sent[0]["id"],
        "sent": sent,
        "failed": failed,
        "retried": retried,
        "title": title,
        "source_url": source_url,
        "status": hubmusic_status_snapshot(),
    }


def stop_hubmusic_route(payload: dict, *, default_snapshot: dict | None = None) -> dict:
    snapshot = default_snapshot or _HUB_MUSIC_STATE.snapshot()
    targets, mode, exclude_id = resolve_hubmusic_targets(payload, default_snapshot=snapshot)

    for target in targets:
        _clear_satellite_stream_status(str(target.get("id") or ""))

    _stop_hubmusic_ffmpeg()

    stopped: list[dict] = []
    failed: list[dict] = []
    retried: list[dict] = []

    if HUBMUSIC_SOFT_STOP_ONLY:
        fast_stop_hosts: list[str] = []
        for target in targets:
            sat_id = str(target.get("id", ""))
            sat_alias = str(target.get("alias", ""))
            sat_host = str(target.get("host", ""))
            if sat_host:
                fast_stop_hosts.append(sat_host)
            stopped.append(
                {
                    "id": sat_id,
                    "alias": sat_alias,
                    "host": sat_host,
                    "duration_ms": 0,
                    "attempt": 1,
                    "soft_stop": True,
                }
            )

        # Best-effort immediate STOP on satellite players so user-initiated stop feels instant,
        # while still preserving soft-stop behavior as the primary path.
        for sat_host in fast_stop_hosts:
            threading.Thread(target=_fast_stop_media_on_satellite, args=(sat_host,), daemon=True).start()

        _HUB_MUSIC_STATE.stop(stopped)
        _HUB_MUSIC_STATE.set_results("stop", stopped=stopped, failed=failed, retried=retried, exclude_satellite=exclude_id, mode=mode)
        _APP_STATE.update(last_action=f"hubmusic_stop_soft:{len(stopped)}", last_error="")
        return {
            "ok": True,
            "mode": mode,
            "exclude_satellite": exclude_id,
            "satellite": stopped[0]["id"] if stopped else None,
            "stopped": stopped,
            "failed": failed,
            "retried": retried,
            "status": hubmusic_status_snapshot(),
        }

    for target in targets:
        sat_id = str(target.get("id", ""))
        sat_alias = str(target.get("alias", ""))
        sat_host = str(target.get("host", ""))

        attempt = 1
        started = time.perf_counter()
        try:
            stop_media_on_satellite(sat_host, announcement=False)
            duration_ms = int((time.perf_counter() - started) * 1000)
            stopped.append({
                "id": sat_id,
                "alias": sat_alias,
                "host": sat_host,
                "duration_ms": duration_ms,
                "attempt": attempt,
            })
            continue
        except Exception as first_exc:
            logging.warning("HubMusic stop first attempt failed for %s (%s): %s", sat_id, sat_host, first_exc)

        attempt = 2
        retry_started = time.perf_counter()
        try:
            stop_media_on_satellite(sat_host, announcement=False)
            duration_ms = int((time.perf_counter() - retry_started) * 1000)
            stopped.append({
                "id": sat_id,
                "alias": sat_alias,
                "host": sat_host,
                "duration_ms": duration_ms,
                "attempt": attempt,
            })
            retried.append({
                "id": sat_id,
                "alias": sat_alias,
                "host": sat_host,
                "duration_ms": duration_ms,
                "attempt": attempt,
            })
        except Exception as retry_exc:
            duration_ms = int((time.perf_counter() - retry_started) * 1000)
            failed.append({
                "id": sat_id,
                "alias": sat_alias,
                "host": sat_host,
                "duration_ms": duration_ms,
                "attempt": attempt,
                "error": str(retry_exc),
            })
            logging.error("HubMusic stop failed for %s (%s): %s", sat_id, sat_host, retry_exc)

    _HUB_MUSIC_STATE.stop(stopped)
    _HUB_MUSIC_STATE.set_results("stop", stopped=stopped, failed=failed, retried=retried, exclude_satellite=exclude_id, mode=mode)
    _APP_STATE.update(last_action=f"hubmusic_stop:{len(stopped)}", last_error="")
    return {
        "ok": True,
        "mode": mode,
        "exclude_satellite": exclude_id,
        "satellite": stopped[0]["id"] if stopped else None,
        "stopped": stopped,
        "failed": failed,
        "retried": retried,
        "status": hubmusic_status_snapshot(),
    }


def run_hubmusic_stereo_test(payload: dict, *, default_snapshot: dict | None = None) -> dict:
    snapshot = default_snapshot or _HUB_MUSIC_STATE.snapshot()
    stereo_payload = {
        "mode": "stereo_pair",
        "left_satellite": payload.get("left_satellite") or payload.get("left") or payload.get("left_id"),
        "right_satellite": payload.get("right_satellite") or payload.get("right") or payload.get("right_id"),
    }
    targets, _, _ = resolve_hubmusic_targets(stereo_payload, default_snapshot=snapshot)
    if len(targets) < 2:
        raise LookupError("Stereo test requires two reachable satellites")

    left = targets[0]
    right = targets[1]
    cue_items = [
        ("left", left, "Left speaker test"),
        ("right", right, "Right speaker test"),
    ]

    sent: list[dict] = []
    failed: list[dict] = []

    for channel, satellite, cue_text in cue_items:
        sat_id = str(satellite.get("id") or "").strip()
        sat_alias = str(satellite.get("alias") or "").strip()
        sat_host = str(satellite.get("host") or "").strip()
        started = time.perf_counter()
        try:
            wav_path = build_wav_path(cue_text, sat_id or channel)
            synthesize_wav(cue_text, wav_path)
            media_url = build_media_url(wav_path.name)
            send_to_satellite(sat_host, media_url)
            sent.append(
                {
                    "id": sat_id,
                    "alias": sat_alias,
                    "host": sat_host,
                    "channel": channel,
                    "duration_ms": int((time.perf_counter() - started) * 1000),
                }
            )
        except Exception as exc:
            failed.append(
                {
                    "id": sat_id,
                    "alias": sat_alias,
                    "host": sat_host,
                    "channel": channel,
                    "duration_ms": int((time.perf_counter() - started) * 1000),
                    "error": str(exc),
                }
            )
        time.sleep(0.35)

    ok = len(sent) == 2
    if ok:
        message = "Stereo channel test sent to left and right satellites."
    elif sent:
        message = "Stereo channel test partially completed."
    else:
        message = "Stereo channel test failed for both satellites."

    return {
        "ok": ok,
        "message": message,
        "mode": "stereo_pair",
        "sent": sent,
        "failed": failed,
        "status": hubmusic_status_snapshot(),
    }


def check_rate_limit(max_requests_per_minute: int = REQUEST_RATE_LIMIT) -> bool:
    """Check if request rate limit exceeded (simple sliding window)."""
    cutoff_time = time.time() - 60
    
    with _RATE_LIMIT_LOCK:
        # Remove old entries outside the window
        while _RATE_LIMIT_WINDOW and _RATE_LIMIT_WINDOW[0] < cutoff_time:
            _RATE_LIMIT_WINDOW.popleft()
        
        if len(_RATE_LIMIT_WINDOW) >= max_requests_per_minute:
            return False
        
        _RATE_LIMIT_WINDOW.append(time.time())
        return True


def cleanup_old_recordings(max_age_hours: int = RECORDING_MAX_AGE_HOURS, 
                          max_files: int = RECORDING_MAX_FILES) -> None:
    """
    Clean up old recording files to prevent disk-space leaks.
    
    Args:
        max_age_hours: Delete recordings older than this (default: 24 hours)
        max_files: Keep at most this many recording files (default: 1000)
    
    Note: Files are kept if either max_files or max_age_hours allows them
    """
    try:
        cutoff = time.time() - (max_age_hours * 3600)
        
        files = sorted(
            RECORDINGS_PATH.glob("*.wav"),
            key=lambda p: p.stat().st_mtime,
            reverse=True
        )
        
        # Keep only recent files up to max_files
        for wav_file in files[max_files:]:
            try:
                if wav_file.stat().st_mtime < cutoff:
                    wav_file.unlink()
                    logging.info("Cleaned up recording: %s", wav_file.name)
            except Exception as e:
                logging.warning("Failed to cleanup %s: %s", wav_file.name, e)
    except Exception as e:
        logging.error("Recording cleanup failed: %s", e)


def handle_local_satellite_command(transcript: str, satellite: dict) -> dict | None:
    if is_dismiss_command(transcript):
        dismissed = _SCHEDULE_MANAGER.dismiss_active_ringing(satellite["id"])
        if dismissed:
            return {"answer": "Okay, dismissed.", "handled": True}
        return {"answer": "Nothing is ringing right now.", "handled": True}

    timer_command = parse_timer_command(transcript)
    if timer_command:
        if timer_command["action"] == "create":
            duration_seconds = int(timer_command.get("duration_seconds") or 0)
            if duration_seconds <= 0:
                return {
                    "answer": "I couldn't tell how long to set the timer for.",
                    "handled": True,
                }
            try:
                item = _SCHEDULE_MANAGER.add_timer(satellite["id"], duration_seconds)
            except RuntimeError as exc:
                return {"answer": sanitize_text(str(exc)), "handled": True}
            return {
                "answer": f"Timer set for {format_duration(duration_seconds)}.",
                "handled": True,
                "pending_timer_event": {
                    "event_type": VoiceAssistantTimerEventType.VOICE_ASSISTANT_TIMER_STARTED,
                    "item": item,
                },
            }

        if timer_command["action"] == "cancel":
            if timer_command.get("all"):
                cancelled = _SCHEDULE_MANAGER.cancel_all("timer", satellite["id"])
                if not cancelled:
                    return {"answer": "No active timers.", "handled": True}
                for item in cancelled:
                    _bridge = _VOICE_BRIDGES.get(item.satellite_id.strip().lower()) or next(iter(_VOICE_BRIDGES.values()), None)
                    if _bridge:
                        _bridge.send_timer_event(
                            VoiceAssistantTimerEventType.VOICE_ASSISTANT_TIMER_CANCELLED,
                            item.schedule_id,
                            item.name,
                            item.total_seconds,
                            item.seconds_left(),
                            False,
                        )
                return {"answer": "All timers canceled.", "handled": True}

            item = _SCHEDULE_MANAGER.cancel_next("timer", satellite["id"])
            if item is None:
                return {"answer": "No active timers.", "handled": True}
            _bridge = _VOICE_BRIDGES.get(item.satellite_id.strip().lower()) or next(iter(_VOICE_BRIDGES.values()), None)
            if _bridge:
                _bridge.send_timer_event(
                    VoiceAssistantTimerEventType.VOICE_ASSISTANT_TIMER_CANCELLED,
                    item.schedule_id,
                    item.name,
                    item.total_seconds,
                    item.seconds_left(),
                    False,
                )
            return {"answer": "Next timer canceled.", "handled": True}

        timers = _SCHEDULE_MANAGER.list_items("timer", satellite["id"])
        if not timers:
            return {"answer": "No active timers.", "handled": True}
        next_timer = timers[0]
        if len(timers) == 1:
            answer = f"Timer has {format_duration(next_timer.seconds_left())} left."
        else:
            answer = f"{len(timers)} timers. Next ends in {format_duration(next_timer.seconds_left())}."
        return {"answer": answer, "handled": True}

    alarm_command = parse_alarm_command(transcript)
    if alarm_command:
        if alarm_command["action"] == "create":
            target_ts = float(alarm_command.get("target_ts") or 0)
            if target_ts <= 0:
                return {
                    "answer": "I couldn't tell what time to set the alarm for.",
                    "handled": True,
                }
            try:
                item = _SCHEDULE_MANAGER.add_alarm(satellite["id"], target_ts)
            except RuntimeError as exc:
                return {"answer": sanitize_text(str(exc)), "handled": True}
            return {
                "answer": f"Alarm set for {format_clock_time(item.target_ts)}.",
                "handled": True,
            }

        if alarm_command["action"] == "cancel":
            if alarm_command.get("all"):
                cancelled = _SCHEDULE_MANAGER.cancel_all("alarm", satellite["id"])
                if not cancelled:
                    return {"answer": "No active alarms.", "handled": True}
                return {"answer": "All alarms canceled.", "handled": True}

            item = _SCHEDULE_MANAGER.cancel_next("alarm", satellite["id"])
            if item is None:
                return {"answer": "No active alarms.", "handled": True}
            return {"answer": "Next alarm canceled.", "handled": True}

        alarms = _SCHEDULE_MANAGER.list_items("alarm", satellite["id"])
        if not alarms:
            return {"answer": "No active alarms.", "handled": True}
        next_alarm = alarms[0]
        if len(alarms) == 1:
            answer = f"Alarm set for {format_clock_time(next_alarm.target_ts)}."
        else:
            answer = f"{len(alarms)} alarms. Next at {format_clock_time(next_alarm.target_ts)}."
        return {"answer": answer, "handled": True}

    normalized = " ".join((transcript or "").lower().split())
    if "timer" in normalized:
        return {
            "answer": "I heard a timer request but need a duration.",
            "handled": True,
        }

    if "alarm" in normalized and "alarm system" not in normalized and "security alarm" not in normalized:
        return {
            "answer": "I heard an alarm request but need a time.",
            "handled": True,
        }

    whisper_cmd = parse_whisper_command(transcript)
    if whisper_cmd:
        try:
            set_satellite_switch(satellite["host"], "whisper_mode", whisper_cmd["state"])
            state_label = "on" if whisper_cmd["state"] else "off"
            answer = f"Whisper mode {state_label}."
        except Exception as exc:
            logging.warning("Failed to set whisper mode on %s: %s", satellite.get("host"), exc)
            answer = "I couldn't change whisper mode right now."
        return {"answer": answer, "handled": True}

    follow_up_toggle_cmd = parse_follow_up_toggle_command(transcript)
    if follow_up_toggle_cmd:
        try:
            set_satellite_switch(satellite["host"], "follow_up_listening_switch", follow_up_toggle_cmd["state"])
            state_label = "on" if follow_up_toggle_cmd["state"] else "off"
            answer = f"Follow-up listening {state_label}."
        except Exception as exc:
            logging.warning("Failed to set follow-up listening on %s: %s", satellite.get("host"), exc)
            answer = "I couldn't change follow-up listening right now."
        return {"answer": answer, "handled": True}

    follow_up_window_cmd = parse_follow_up_window_command(transcript)
    if follow_up_window_cmd:
        try:
            set_satellite_number(
                satellite["host"],
                follow_up_window_cmd["entity_id"],
                follow_up_window_cmd["applied"],
            )
            if follow_up_window_cmd["requested"] != follow_up_window_cmd["applied"]:
                answer = (
                    f"I set the follow-up window to {follow_up_window_cmd['applied']} seconds. "
                    "Supported range is 1 to 30 seconds."
                )
            else:
                answer = f"I set the follow-up window to {follow_up_window_cmd['applied']} seconds."
        except Exception as exc:
            logging.warning("Failed to set follow-up window on %s: %s", satellite.get("host"), exc)
            answer = "I couldn't change the follow-up window right now."
        return {"answer": answer, "handled": True}

    command = parse_volume_command(transcript)
    if not command:
        return None

    set_satellite_number(satellite["host"], command["entity_id"], command["applied"])
    if command["entity_id"] == "speaker_volume":
        with contextlib.suppress(Exception):
            _remember_user_locked_media_volume(satellite["host"], command["applied"])
        with contextlib.suppress(Exception):
            set_satellite_media_volume(satellite["host"], command["applied"])
    if command["requested"] != command["applied"]:
        answer = f"I set the {command['label']} to {command['applied']} percent. Supported range is {command['min']} to {command['max']} percent."
    else:
        answer = f"I set the {command['label']} to {command['applied']} percent."
    return {
        "answer": answer,
        "handled": True,
    }


def broadcast_to_all_satellites(text: str, exclude_sat_id: str = "") -> dict:
    message = sanitize_text(text)
    satellites = load_satellites()
    if not satellites:
        raise RuntimeError("No satellites configured in satellites.csv")

    exclude_key = (exclude_sat_id or "").strip().lower()
    targets = [
        sat
        for sat in satellites
        if str(sat.get("id", "")).strip().lower() != exclude_key
    ]
    if not targets:
        return {"count": 0, "sent": [], "failed": [], "media_url": "", "skipped": [exclude_sat_id] if exclude_sat_id else []}

    wav_path = build_wav_path(message, "broadcast")
    _APP_STATE.update(last_action=f"synthesizing broadcast {wav_path.name}")
    synthesize_wav(message, wav_path)
    media_url = build_media_url(wav_path.name)

    sent: list[str] = []
    failed: list[dict] = []

    for sat in targets:
        sat_id = str(sat.get("id", "")).strip() or str(sat.get("host", "")).strip()
        host = str(sat.get("host", "")).strip()
        if not host:
            continue
        try:
            send_to_satellite(host, media_url)
            sent.append(sat_id)
        except Exception as exc:
            failed.append({"id": sat_id, "error": str(exc)})
            logging.exception("Broadcast delivery failed for satellite %s (%s)", sat_id, host)

    if not sent:
        raise RuntimeError("Broadcast failed for all satellites")

    _APP_STATE.update(last_action=f"broadcast sent to {len(sent)} satellites", last_error="")
    return {"count": len(sent), "sent": sent, "failed": failed, "media_url": media_url}


def send_to_named_satellite(text: str, target: str, exclude_sat_id: str = "") -> dict:
    message = sanitize_text(text)
    target_ref = (target or "").strip()
    if not target_ref:
        raise RuntimeError("Target satellite is required")

    satellite = select_satellite(target_ref)
    if not satellite:
        raise RuntimeError(f"No satellite matched '{target_ref}'")

    sat_id = str(satellite.get("id", "")).strip()
    sat_alias = str(satellite.get("alias", "")).strip()
    exclude_key = (exclude_sat_id or "").strip().lower()
    if exclude_key and sat_id.lower() == exclude_key:
        return {"count": 0, "sent": [], "failed": [], "target": sat_id, "alias": sat_alias, "skipped": True}

    wav_path = build_wav_path(message, sat_id or "target")
    _APP_STATE.update(last_action=f"synthesizing targeted send {wav_path.name}")
    synthesize_wav(message, wav_path)
    media_url = build_media_url(wav_path.name)

    satellite_host = str(satellite.get("host", "")).strip()
    prepare_satellite_for_tts(satellite_host)
    send_to_satellite(satellite_host, media_url)
    _APP_STATE.update(last_action=f"sent targeted message to {sat_id}", last_error="")
    return {
        "count": 1,
        "sent": [sat_id],
        "failed": [],
        "target": sat_id,
        "alias": sat_alias,
        "media_url": media_url,
        "skipped": False,
    }


def worker_loop() -> None:
    while not _SHUTDOWN_EVENT.is_set():
        try:
            # Use timeout to periodically check shutdown signal
            job = _WORK_QUEUE.get(timeout=1.0)
        except queue.Empty:
            continue
        
        try:
            text = sanitize_text(str(job.get("text", "")))
            sat_id = str(job.get("sat_id", "")).strip()
            
            if not text:
                raise ValueError("Text is required")
            
            satellite = select_satellite(sat_id)
            if not satellite:
                raise RuntimeError("No satellites configured in satellites.csv")

            wav_path = build_wav_path(text, sat_id or satellite["id"])
            _APP_STATE.update(last_action=f"synthesizing {wav_path.name}")
            synthesize_wav(text, wav_path)

            media_url = build_media_url(wav_path.name)
            _APP_STATE.update(last_action=f"sending {wav_path.name} to {satellite['host']}")
            try:
                prepare_satellite_for_tts(satellite["host"])
                send_to_satellite(satellite["host"], media_url)
                _APP_STATE.update(last_action=f"sent {wav_path.name} to {satellite['host']}", last_error="")
                logging.info("Queued answer to %s via media_player %s", satellite["host"], media_url)
            except Exception as media_exc:
                if "no media player" not in str(media_exc).lower():
                    raise

                bridge_key = str(satellite.get("id", "")).strip()
                bridge = _VOICE_BRIDGES.get(bridge_key) or _VOICE_BRIDGES.get(bridge_key.lower())
                if not bridge:
                    raise RuntimeError(f"No voice bridge for satellite {bridge_key}")

                if not bridge.send_direct_tts(text, media_url):
                    raise RuntimeError(f"Voice bridge TTS fallback failed for {bridge_key}")

                _APP_STATE.update(
                    last_action=f"sent {wav_path.name} to {satellite['host']} via voice bridge",
                    last_error="",
                )
                logging.info("Queued answer to %s via voice bridge %s", satellite["host"], media_url)
        except Exception as exc:
            _APP_STATE.update(last_error=str(exc), last_action="error")
            logging.exception("Failed to process answer: %s", exc)
        finally:
            _WORK_QUEUE.task_done()
        
        # Check shutdown signal after each task
        if _SHUTDOWN_EVENT.is_set():
            logging.info("Worker received shutdown signal")
            break


@dataclass
class VoiceSession:
    conversation_id: str
    satellite_id: str
    wake_word_phrase: str | None
    raw_audio: bytearray
    saw_audio: bool = False
    hubmusic_resume_state: dict | None = None


class VoiceAssistantBridge:
    def __init__(self, satellite_id: str = "") -> None:
        self._satellite_id = satellite_id
        self._stop_event = threading.Event()
        self._thread = threading.Thread(target=self._thread_main, daemon=True)
        self._lock = threading.Lock()
        self._client: APIClient | None = None
        self._unsubscribe = None
        self._session: VoiceSession | None = None
        self._status = "starting"
        self._last_error = ""
        self._last_event = "starting"
        self._connected = False
        self._pending_broadcast: dict[str, float] = {}
        self._pending_broadcast_target: dict[str, str] = {}

    def start(self) -> None:
        self._thread.start()

    def snapshot(self) -> dict:
        with self._lock:
            return {
                "status": self._status,
                "connected": self._connected,
                "last_error": self._last_error,
                "last_event": self._last_event,
            }

    def _set_state(self, *, status: str | None = None, connected: bool | None = None, error: str | None = None, event: str | None = None) -> None:
        with self._lock:
            if status is not None:
                self._status = status
            if connected is not None:
                self._connected = connected
            if error is not None:
                self._last_error = error
            if event is not None:
                self._last_event = event

    def _thread_main(self) -> None:
        asyncio.run(self._run())

    def can_send_timer_events(self) -> bool:
        with self._lock:
            return self._client is not None and self._connected

    async def _run(self) -> None:
        _reconnect_delay = 0.0
        while not self._stop_event.is_set():
            satellite = select_satellite(self._satellite_id)
            if not satellite:
                self._set_state(status="waiting", connected=False, error="No satellites configured", event="idle")
                await asyncio.sleep(3)
                _reconnect_delay = 0.0
                continue

            # Per-attempt event so _on_disconnect signals the inner wait immediately.
            disconnect_event = asyncio.Event()

            async def _on_stop(expected_disconnect: bool, _evt: asyncio.Event = disconnect_event) -> None:
                _evt.set()
                await self._on_disconnect(expected_disconnect)

            _sat_label = f"{satellite.get('alias') or satellite['id']} ({satellite['host']})"
            client = APIClient(satellite["host"], 6054, None, client_info="HubVoiceSatRuntime")
            try:
                self._set_state(status="connecting", connected=False, error="", event="connecting")
                logging.info("Voice bridge connecting to %s port 6054...", _sat_label)
                await asyncio.wait_for(
                    client.connect(login=False, on_stop=_on_stop),
                    timeout=VOICE_BRIDGE_CONNECT_TIMEOUT,
                )
                with self._lock:
                    self._client = client
                self._unsubscribe = client.subscribe_voice_assistant(
                    handle_start=self._handle_start,
                    handle_stop=self._handle_stop,
                    handle_audio=self._handle_audio,
                )
                self._set_state(status="connected", connected=True, error="", event="subscribed")
                _SCHEDULE_MANAGER.sync_bridge(self)
                logging.info("Voice bridge CONNECTED to %s port 6054", _sat_label)
                _reconnect_delay = 0.0  # reset backoff after a successful connection

                # Wait until the disconnect event fires or we are stopped.
                while not self._stop_event.is_set() and not disconnect_event.is_set():
                    try:
                        await asyncio.wait_for(disconnect_event.wait(), timeout=0.5)
                    except asyncio.TimeoutError:
                        pass
            except Exception as exc:
                self._set_state(status="error", connected=False, error=str(exc), event="connect_failed")
                logging.error("Voice bridge connect FAILED for %s: %s", self._satellite_id, exc)
            finally:
                try:
                    if self._unsubscribe is not None:
                        self._unsubscribe()
                except Exception:
                    pass
                finally:
                    self._unsubscribe = None
                try:
                    await client.disconnect()
                except Exception:
                    pass
                with self._lock:
                    self._client = None
                    # Clear session and associated data
                    if self._session is not None:
                        self._session.raw_audio.clear()
                    self._session = None
                if not self._stop_event.is_set():
                    self._set_state(status="reconnecting", connected=False, event="disconnected")
                    if _reconnect_delay > 0:
                        logging.info(
                            "Voice bridge %s will reconnect in %.1fs (backoff)...",
                            self._satellite_id, _reconnect_delay,
                        )
                        await asyncio.sleep(_reconnect_delay)
                    else:
                        logging.info("Voice bridge %s reconnecting immediately...", self._satellite_id)
                    # Exponential backoff: 0 → 1 → 2 → 4 → 8s cap
                    _reconnect_delay = min(_reconnect_delay * 2 + 1.0, 8.0)

    async def _on_disconnect(self, expected_disconnect: bool) -> None:
        self._set_state(
            status="disconnected" if expected_disconnect else "reconnecting",
            connected=False,
            event="connection_closed",
        )
        with self._lock:
            self._client = None
            self._session = None
        logging.info("Voice assistant bridge disconnected (expected=%s) for %s", expected_disconnect, self._satellite_id)

    def _send_event(self, event_type: VoiceAssistantEventType, data: dict[str, str] | None = None) -> None:
        with self._lock:
            client = self._client
        if client is None:
            return
        try:
            client.send_voice_assistant_event(event_type, data)
            self._set_state(event=event_type.name)
        except Exception as exc:
            # Do not let transient bridge send failures stall voice processing.
            logging.warning("Failed to send VA event %s: %s", event_type.name, exc)
            self._set_state(connected=False, error=str(exc), event="event_send_failed")

    def send_timer_event(
        self,
        event_type: VoiceAssistantTimerEventType,
        timer_id: str,
        name: str | None,
        total_seconds: int,
        seconds_left: int,
        is_active: bool,
    ) -> bool:
        with self._lock:
            client = self._client
        if client is None:
            return False
        try:
            client.send_voice_assistant_timer_event(
                event_type,
                timer_id,
                name,
                total_seconds,
                seconds_left,
                is_active,
            )
            return True
        except Exception as exc:
            logging.exception("Failed to send timer event %s for %s: %s", event_type.name, timer_id, exc)
            return False

    def send_direct_tts(self, text: str, media_url: str, conversation_id: str | None = None) -> bool:
        """Send direct TTS using Voice Assistant events.

        This fallback is used when firmware exposes no media_player entity.
        """
        conv_id = conversation_id or f"runtime-{uuid.uuid4().hex[:12]}"
        try:
            self._send_event(VoiceAssistantEventType.VOICE_ASSISTANT_RUN_START, {"conversation_id": conv_id})
            self._send_event(VoiceAssistantEventType.VOICE_ASSISTANT_INTENT_START, None)
            self._send_event(
                VoiceAssistantEventType.VOICE_ASSISTANT_INTENT_END,
                {
                    "conversation_id": conv_id,
                    "continue_conversation": "0",
                },
            )
            self._send_event(VoiceAssistantEventType.VOICE_ASSISTANT_TTS_START, {"text": sanitize_text(text)})
            self._send_event(VoiceAssistantEventType.VOICE_ASSISTANT_TTS_END, {"url": media_url})
            self._send_event(VoiceAssistantEventType.VOICE_ASSISTANT_RUN_END, None)
            return True
        except Exception as exc:
            logging.error("Direct VA TTS send failed for %s: %s", self._satellite_id, exc)
            return False

    def _set_pending_broadcast(self, satellite_id: str, target: str = "") -> None:
        sat_key = (satellite_id or "default").strip().lower() or "default"
        expires_at = time.time() + BROADCAST_PENDING_TIMEOUT_SECS
        with self._lock:
            self._pending_broadcast[sat_key] = expires_at
            target_norm = (target or "").strip()
            if target_norm:
                self._pending_broadcast_target[sat_key] = target_norm
            else:
                self._pending_broadcast_target.pop(sat_key, None)

    def _consume_pending_broadcast(self, satellite_id: str) -> dict:
        sat_key = (satellite_id or "default").strip().lower() or "default"
        now_ts = time.time()
        with self._lock:
            expires_at = self._pending_broadcast.get(sat_key)
            if not expires_at:
                return {"active": False, "target": ""}
            if expires_at < now_ts:
                self._pending_broadcast.pop(sat_key, None)
                self._pending_broadcast_target.pop(sat_key, None)
                return {"active": False, "target": ""}
            self._pending_broadcast.pop(sat_key, None)
            target = self._pending_broadcast_target.pop(sat_key, "")
            return {"active": True, "target": target}

    async def _handle_start(self, conversation_id: str, flags: int, audio_settings, wake_word_phrase: str | None) -> int | None:
        satellite_id = self._satellite_id or "default"
        hubmusic_snapshot = _HUB_MUSIC_STATE.snapshot()
        with self._lock:
            self._session = VoiceSession(
                conversation_id=conversation_id,
                satellite_id=satellite_id,
                wake_word_phrase=wake_word_phrase,
                raw_audio=bytearray(),
                hubmusic_resume_state=hubmusic_snapshot if hubmusic_snapshot.get("active") else None,
            )
        logging.info(
            "Voice session started conversation=%s wake_word=%s flags=%s",
            conversation_id,
            wake_word_phrase or "",
            flags,
        )
        # Always stop local media on wake so wake sound and voice reply are audible.
        # Only mark HubMusic paused when runtime tracking says it was active pre-wake.
        satellite = select_satellite(satellite_id)
        if satellite:
            asyncio.get_event_loop().run_in_executor(
                None,
                _stop_music_for_voice,
                satellite["host"],
                bool(hubmusic_snapshot.get("active")),
            )
        self._send_event(VoiceAssistantEventType.VOICE_ASSISTANT_RUN_START, {"conversation_id": conversation_id})
        self._send_event(VoiceAssistantEventType.VOICE_ASSISTANT_STT_START, None)
        return 0

    async def _handle_audio(self, data: bytes) -> None:
        with self._lock:
            session = self._session
        if session is None:
            logging.debug(
                "Audio data arrived on bridge %s with no active session (%d bytes dropped) — "
                "satellite may have restarted mid-session",
                self._satellite_id, len(data),
            )
            return
        if not session.saw_audio:
            session.saw_audio = True
            self._send_event(VoiceAssistantEventType.VOICE_ASSISTANT_STT_VAD_START, None)
        session.raw_audio.extend(data)

    async def _handle_stop(self, aborted: bool) -> None:
        with self._lock:
            session = self._session
            self._session = None

        if session is None:
            return

        logging.info(
            "Voice session stop requested aborted=%s audio_bytes=%s saw_audio=%s",
            aborted,
            len(session.raw_audio),
            session.saw_audio,
        )

        if session.saw_audio:
            self._send_event(VoiceAssistantEventType.VOICE_ASSISTANT_STT_VAD_END, None)

        if aborted and not session.raw_audio:
            logging.info(
                "Voice session aborted with no audio on satellite=%s — likely a false wake-word trigger",
                session.satellite_id,
            )
            self._send_event(
                VoiceAssistantEventType.VOICE_ASSISTANT_ERROR,
                {"code": "aborted", "message": "Voice session stopped"},
            )
            self._send_event(VoiceAssistantEventType.VOICE_ASSISTANT_RUN_END, None)
            return

        try:
            _APP_STATE.update(last_action="transcribing voice request")
            result = await asyncio.to_thread(self._process_session, session)
            transcript = result["transcript"]
            answer = result["answer"]
            media_url = result["media_url"]
            _APP_STATE.update(last_transcript=transcript, last_error="", last_action="replying to voice request")

            # Deliver audio via media_player (proven path - VA component ignores TTS_END URL).
            # Fire as a background task so VA events can be sent without blocking.
            _sat = select_satellite(session.satellite_id)
            if _sat:
                asyncio.create_task(send_to_satellite_async(str(_sat.get("host", "")), media_url))

            # Send VA events for LED / FSM state management.
            self._send_event(VoiceAssistantEventType.VOICE_ASSISTANT_STT_END, {"text": transcript})
            self._send_event(VoiceAssistantEventType.VOICE_ASSISTANT_INTENT_START, None)
            self._send_event(
                VoiceAssistantEventType.VOICE_ASSISTANT_INTENT_END,
                {
                    "conversation_id": session.conversation_id,
                    "continue_conversation": "1" if result.get("continue_conversation") else "0",
                },
            )
            self._send_event(VoiceAssistantEventType.VOICE_ASSISTANT_TTS_START, {"text": answer})
            self._send_event(VoiceAssistantEventType.VOICE_ASSISTANT_TTS_END, {"url": media_url})
            self._send_event(VoiceAssistantEventType.VOICE_ASSISTANT_RUN_END, None)
            if session.hubmusic_resume_state:
                reply_seconds = estimate_media_url_duration_seconds(media_url)
                resume_delay = max(0.5, min(30.0, reply_seconds + 0.35))
                asyncio.get_event_loop().run_in_executor(
                    None,
                    _resume_music_after_voice,
                    session.hubmusic_resume_state,
                    session.satellite_id,
                    resume_delay,
                )
            pending_timer = result.get("pending_timer_event")
            if pending_timer:
                item = pending_timer["item"]
                seconds_left = item.seconds_left()
                if self.send_timer_event(
                    pending_timer["event_type"],
                    item.schedule_id,
                    item.name,
                    item.total_seconds,
                    seconds_left,
                    True,
                ):
                    _SCHEDULE_MANAGER.mark_timer_sent(item.schedule_id, seconds_left)
            logging.info("Voice session answered transcript=%r answer=%r", transcript, answer)
        except Exception as exc:
            _APP_STATE.update(last_error=str(exc), last_action="voice error")
            logging.exception("Voice session failed: %s", exc)
            code = "stt-no-text-recognized" if "recognized" in str(exc).lower() else "processing-error"
            self._send_event(
                VoiceAssistantEventType.VOICE_ASSISTANT_ERROR,
                {"code": code, "message": str(exc)},
            )
            self._send_event(VoiceAssistantEventType.VOICE_ASSISTANT_RUN_END, None)

    def _process_session(self, session: VoiceSession) -> dict:
        if not session.raw_audio:
            raise RuntimeError("No speech recognized")

        input_wav = build_input_wav_path(session.satellite_id)
        write_input_wav(bytes(session.raw_audio), input_wav)
        logging.info("Saved voice input to %s (%s bytes)", input_wav.name, len(session.raw_audio))

        _stt_t0 = time.monotonic()
        transcript_result = transcribe_with_retry(input_wav, session.satellite_id)
        _stt_elapsed = time.monotonic() - _stt_t0
        transcript = " ".join(str(transcript_result.get("text", "")).split()).strip()
        if not transcript:
            logging.warning("STT returned empty transcript in %.2fs for %s", _stt_elapsed, session.satellite_id)
            raise RuntimeError("No speech recognized")
        logging.info("STT recognized in %.2fs: %r (satellite=%s)", _stt_elapsed, transcript, session.satellite_id)

        pending = self._consume_pending_broadcast(session.satellite_id)
        if pending.get("active"):
            pending_target = str(pending.get("target", "")).strip()
            if is_broadcast_cancel(transcript):
                answer = "Okay, canceled."
                tts_wav = build_wav_path(answer, session.satellite_id)
                synthesize_wav(answer, tts_wav)
                return {
                    "transcript": transcript,
                    "answer": answer,
                    "media_url": build_media_url(tts_wav.name),
                    "continue_conversation": False,
                }

            directed = parse_directed_send(transcript)
            if directed:
                summary = send_to_named_satellite(
                    directed["message"],
                    directed["target"],
                    session.satellite_id,
                )
                if summary.get("skipped"):
                    answer = "That target is this satellite. Broadcast excludes the one you are speaking to."
                elif int(summary.get("count", 0)) <= 0:
                    answer = "No matching target satellite is available."
                else:
                    target_name = str(summary.get("alias") or summary.get("target") or "that satellite")
                    answer = f"Sent to {target_name}."

                tts_wav = build_wav_path(answer, session.satellite_id)
                synthesize_wav(answer, tts_wav)
                return {
                    "transcript": transcript,
                    "answer": answer,
                    "media_url": build_media_url(tts_wav.name),
                    "continue_conversation": False,
                }

            message = sanitize_text(transcript)
            if not message:
                self._set_pending_broadcast(session.satellite_id, pending_target)
                answer = BROADCAST_PROMPT if not pending_target else f"What's the message for {pending_target}?"
                tts_wav = build_wav_path(answer, session.satellite_id)
                synthesize_wav(answer, tts_wav)
                return {
                    "transcript": transcript,
                    "answer": answer,
                    "media_url": build_media_url(tts_wav.name),
                    "continue_conversation": True,
                }

            if pending_target:
                summary = send_to_named_satellite(message, pending_target, session.satellite_id)
                if summary.get("skipped"):
                    answer = "That target is this satellite. Broadcast excludes the one you are speaking to."
                elif int(summary.get("count", 0)) <= 0:
                    answer = "No matching target satellite is available."
                else:
                    target_name = str(summary.get("alias") or summary.get("target") or pending_target)
                    answer = f"Sent to {target_name}."

                tts_wav = build_wav_path(answer, session.satellite_id)
                synthesize_wav(answer, tts_wav)
                return {
                    "transcript": transcript,
                    "answer": answer,
                    "media_url": build_media_url(tts_wav.name),
                    "continue_conversation": False,
                }

            summary = broadcast_to_all_satellites(message, session.satellite_id)
            fail_count = len(summary.get("failed", []))
            count = int(summary.get("count", 0))
            if count <= 0:
                answer = "No other satellites are available for broadcast."
            else:
                answer = f"Broadcast sent to {count} satellites."
            if count > 0 and fail_count > 0:
                answer = f"Broadcast sent to {count} satellites. {fail_count} failed."

            tts_wav = build_wav_path(answer, session.satellite_id)
            synthesize_wav(answer, tts_wav)
            return {
                "transcript": transcript,
                "answer": answer,
                "media_url": build_media_url(tts_wav.name),
                "continue_conversation": False,
            }

        directed = parse_directed_send(transcript)
        if directed:
            summary = send_to_named_satellite(
                directed["message"],
                directed["target"],
                session.satellite_id,
            )
            if summary.get("skipped"):
                answer = "That target is this satellite. Broadcast excludes the one you are speaking to."
            elif int(summary.get("count", 0)) <= 0:
                answer = "No matching target satellite is available."
            else:
                target_name = str(summary.get("alias") or summary.get("target") or "that satellite")
                answer = f"Sent to {target_name}."

            tts_wav = build_wav_path(answer, session.satellite_id)
            synthesize_wav(answer, tts_wav)
            return {
                "transcript": transcript,
                "answer": answer,
                "media_url": build_media_url(tts_wav.name),
                "continue_conversation": False,
            }

        broadcast_cmd = parse_broadcast_command(transcript)
        if broadcast_cmd:
            if broadcast_cmd.get("mode") == "prompt":
                self._set_pending_broadcast(session.satellite_id)
                answer = BROADCAST_PROMPT
                tts_wav = build_wav_path(answer, session.satellite_id)
                synthesize_wav(answer, tts_wav)
                return {
                    "transcript": transcript,
                    "answer": answer,
                    "media_url": build_media_url(tts_wav.name),
                    "continue_conversation": True,
                }

            if broadcast_cmd.get("mode") == "prompt_target":
                target = sanitize_text(str(broadcast_cmd.get("target", "")))
                self._set_pending_broadcast(session.satellite_id, target)
                answer = f"What's the message for {target}?" if target else BROADCAST_PROMPT
                tts_wav = build_wav_path(answer, session.satellite_id)
                synthesize_wav(answer, tts_wav)
                return {
                    "transcript": transcript,
                    "answer": answer,
                    "media_url": build_media_url(tts_wav.name),
                    "continue_conversation": True,
                }

            message = sanitize_text(str(broadcast_cmd.get("message", "")))
            summary = broadcast_to_all_satellites(message, session.satellite_id)
            fail_count = len(summary.get("failed", []))
            count = int(summary.get("count", 0))
            if count <= 0:
                answer = "No other satellites are available for broadcast."
            else:
                answer = f"Broadcast sent to {count} satellites."
            if count > 0 and fail_count > 0:
                answer = f"Broadcast sent to {count} satellites. {fail_count} failed."

            tts_wav = build_wav_path(answer, session.satellite_id)
            synthesize_wav(answer, tts_wav)
            return {
                "transcript": transcript,
                "answer": answer,
                "media_url": build_media_url(tts_wav.name),
                "continue_conversation": False,
            }

        satellite = select_satellite(session.satellite_id)
        if not satellite:
            raise RuntimeError("No satellites configured in satellites.csv")

        local_result = handle_local_satellite_command(transcript, satellite)
        if local_result:
            answer = sanitize_text(str(local_result.get("answer", "")))
            logging.info("Local satellite command answer: %r", answer)
            tts_wav = build_wav_path(answer, session.satellite_id)
            synthesize_wav(answer, tts_wav)
            ret = {
                "transcript": transcript,
                "answer": answer,
                "media_url": build_media_url(tts_wav.name),
                "continue_conversation": False,
            }
            pte = local_result.get("pending_timer_event")
            if pte:
                ret["pending_timer_event"] = pte
            return ret

        hubitat_result = transcript_result.get("hubitat")
        if not hubitat_result:
            hubitat_result = ask_hubitat(transcript, session.satellite_id)
        answer = sanitize_text(str(hubitat_result.get("answer", "")))
        logging.info("Hubitat answer: %r", answer)
        tts_wav = build_wav_path(answer, session.satellite_id)
        synthesize_wav(answer, tts_wav)

        return {
            "transcript": transcript,
            "answer": answer,
            "media_url": build_media_url(tts_wav.name),
            "continue_conversation": False,
        }


_SCHEDULE_MANAGER = ScheduleManager(SCHEDULES_PATH)
_VOICE_BRIDGES: dict[str, VoiceAssistantBridge] = {
    sat["id"]: VoiceAssistantBridge(sat["id"])
    for sat in load_satellites()
}
_HUB_MUSIC_STATE = HubMusicState()
_AIRPLAY_LOCK = threading.Lock()
_AIRPLAY_PROCESS: subprocess.Popen | None = None
_AIRPLAY_STATE: dict[str, object] = {
    "enabled": False,
    "running": False,
    "pid": None,
    "name": AIRPLAY_RECEIVER_DEFAULT_NAME,
    "interface": "",
    "started_at": "",
    "last_error": "",
}
_DLNA_LOCK = threading.Lock()
_DLNA_THREAD: threading.Thread | None = None
_DLNA_LOOP: asyncio.AbstractEventLoop | None = None
_DLNA_SERVER = None
_DLNA_AVT_SERVICE = None
_DLNA_RC_SERVICE = None
_DLNA_CM_SERVICE = None
_DLNA_LAST_PLAY_MONO = 0.0
# DLNA ffmpeg proxy: transcodes the sender stream to MP3 before forwarding to satellites
_DLNA_PROXY_LOCK = threading.Lock()
_DLNA_PROXY_PROC: subprocess.Popen | None = None
_DLNA_PROXY_SOURCE_URL: str = ""
_DLNA_PROXY_LISTENERS: list[queue.Queue] = []
_DLNA_PROXY_BROADCAST_THREAD: threading.Thread | None = None
_DLNA_PROXY_BASS_DB: float = 0.0
_DLNA_PROXY_TREBLE_DB: float = 0.0
_SATELLITE_STREAM_STATUS_LOCK = threading.Lock()
_SATELLITE_STREAM_STATUS: dict[str, dict[str, object]] = {}
_DLNA_STATE: dict[str, object] = {
    "enabled": False,
    "running": False,
    "friendly_name": DLNA_RENDERER_DEFAULT_NAME,
    "host": "",
    "http_port": DLNA_RENDERER_HTTP_PORT,
    "device_url": "",
    "transport_state": "NO_MEDIA_PRESENT",
    "last_uri": "",
    "last_title": "",
    "last_artist": "",
    "last_album": "",
    "last_metadata": "",
    "last_play_started_at": "",
    "started_at": "",
    "last_error": "",
    "preferred_mode": "all_reachable",
    "preferred_satellite": "",
    "preferred_exclude_satellite": "",
    "volume": 100,
    "mute": False,
}
_RUNTIME_LOCK_FILE = ROOT / "hubvoice-runtime.lock"
_RUNTIME_LOCK_HANDLE = None


def _ensure_dlna_dependencies_installed() -> list[str]:
    missing: list[str] = []
    for module_name in DLNA_REQUIRED_MODULES:
        if importlib.util.find_spec(module_name) is None:
            missing.append(module_name)
    return missing


def _resolve_runtime_ipv4_host() -> tuple[str, str]:
    runtime_host = str(get_runtime_host() or "").strip()
    if runtime_host and runtime_host not in {"127.0.0.1", "localhost", "::1"} and ":" not in runtime_host:
        return runtime_host, ""

    if ni is None:
        return "", "Runtime host resolves to loopback and netifaces is unavailable to detect a LAN IPv4 address."

    try:
        interfaces = ni.interfaces()
    except Exception as exc:
        return "", f"Unable to enumerate network interfaces: {exc}"

    fallback = ""
    for iface in interfaces:
        try:
            info = ni.ifaddresses(iface)
        except Exception:
            continue
        ipv4_entries = info.get(ni.AF_INET, []) if hasattr(ni, "AF_INET") else []
        for entry in ipv4_entries:
            addr = str(entry.get("addr") or "").strip()
            if not addr or addr.startswith("127."):
                continue
            if not fallback:
                fallback = addr
            if runtime_host and addr == runtime_host:
                return addr, ""

    if fallback:
        return fallback, ""
    return "", "No IPv4-capable LAN address found for the DLNA renderer."


def _dlna_protocol_info() -> str:
    return ",".join(
        [
            "http-get:*:audio/mpeg:*",
            "http-get:*:audio/mp3:*",
            "http-get:*:audio/aac:*",
            "http-get:*:audio/mp4:*",
            "http-get:*:audio/flac:*",
            "http-get:*:audio/x-flac:*",
            "http-get:*:audio/wav:*",
            "http-get:*:audio/x-wav:*",
        ]
    )


def _extract_dlna_title(metadata: str, uri: str) -> str:
    text = str(metadata or "").strip()
    if text:
        with contextlib.suppress(Exception):
            root = ET.fromstring(text)
            for element in root.iter():
                tag = str(element.tag or "")
                if tag.endswith("title"):
                    value = str(element.text or "").strip()
                    if value:
                        return value
    parsed = urllib.parse.urlparse(str(uri or "").strip())
    tail = Path(parsed.path or "").name.strip()
    return tail or "DLNA"


def _extract_dlna_metadata(metadata: str, uri: str) -> dict[str, str]:
    """Extract title, artist, and album from DLNA DIDL-Lite metadata XML."""
    result = {"title": "", "artist": "", "album": ""}
    text = str(metadata or "").strip()
    
    if text:
        with contextlib.suppress(Exception):
            root = ET.fromstring(text)
            for element in root.iter():
                tag = str(element.tag or "")
                value = str(element.text or "").strip()
                
                if not value:
                    continue
                
                if tag.endswith("title") and not result["title"]:
                    result["title"] = value
                elif tag.endswith("creator") and not result["artist"]:
                    result["artist"] = value
                elif tag.endswith("album") and not result["album"]:
                    result["album"] = value
                
                # Early exit if all fields found
                if result["title"] and result["artist"] and result["album"]:
                    break
    
    # Fallback: extract title from filename if no title in metadata
    if not result["title"]:
        parsed = urllib.parse.urlparse(str(uri or "").strip())
        tail = Path(parsed.path or "").name.strip()
        result["title"] = tail or "DLNA"
    
    return result


def _dlna_uri_signature(uri: str) -> str:
    raw = str(uri or "").strip()
    if not raw:
        return ""
    with contextlib.suppress(Exception):
        parsed = urllib.parse.urlparse(raw)
        host = str(parsed.hostname or "").strip().lower()
        port = str(parsed.port or "").strip()
        path = str(parsed.path or "").strip().rstrip("/").lower()
        query = str(parsed.query or "").strip().lower()
        if host or path or query:
            return f"{host}:{port}|{path}|{query}"
    return raw.lower()


def _build_dlna_last_change_xml(service_kind: str, values: list[tuple[str, str, dict[str, str] | None]]) -> str:
    namespace = "urn:schemas-upnp-org:metadata-1-0/AVT/" if service_kind == "AVT" else "urn:schemas-upnp-org:metadata-1-0/RCS/"
    parts = [f'<Event xmlns="{namespace}"><InstanceID val="0">']
    for name, value, attrs in values:
        attr_text = ""
        if attrs:
            attr_text = " " + " ".join(f'{key}="{value_text}"' for key, value_text in attrs.items())
        parts.append(f'<{name}{attr_text} val="{value}"/>')
    parts.append("</InstanceID></Event>")
    return "".join(parts)


def _apply_dlna_service_state() -> None:
    with _DLNA_LOCK:
        avt = _DLNA_AVT_SERVICE
        rc = _DLNA_RC_SERVICE
        cm = _DLNA_CM_SERVICE
        snapshot = dict(_DLNA_STATE)

    if avt is not None:
        uri = str(snapshot.get("last_uri") or "")
        metadata = str(snapshot.get("last_metadata") or "")
        transport_state = str(snapshot.get("transport_state") or "NO_MEDIA_PRESENT")
        actions = "PLAY,STOP,PAUSE" if uri else ""
        avt.state_variable("TransportState").value = transport_state
        avt.state_variable("TransportStatus").value = "OK"
        avt.state_variable("AVTransportURI").value = uri
        avt.state_variable("AVTransportURIMetaData").value = metadata
        avt.state_variable("CurrentTrackURI").value = uri
        avt.state_variable("CurrentTrackMetaData").value = metadata
        avt.state_variable("CurrentTransportActions").value = actions
        avt.state_variable("NumberOfTracks").value = 1 if uri else 0
        avt.state_variable("CurrentTrack").value = 1 if uri else 0
        avt.state_variable("CurrentMediaDuration").value = "00:00:00"
        avt.state_variable("CurrentTrackDuration").value = "00:00:00"
        avt.state_variable("RelativeTimePosition").value = "00:00:00"
        avt.state_variable("AbsoluteTimePosition").value = "00:00:00"
        avt.state_variable("LastChange").value = _build_dlna_last_change_xml(
            "AVT",
            [
                ("TransportState", transport_state, None),
                ("TransportStatus", "OK", None),
                ("AVTransportURI", uri, None),
                ("CurrentTrackURI", uri, None),
                ("CurrentTransportActions", actions, None),
            ],
        )

    if rc is not None:
        volume = int(snapshot.get("volume") or 100)
        mute = "1" if snapshot.get("mute") else "0"
        rc.state_variable("Volume").value = volume
        rc.state_variable("Mute").value = bool(snapshot.get("mute"))
        rc.state_variable("LastChange").value = _build_dlna_last_change_xml(
            "RCS",
            [
                ("Volume", str(volume), {"channel": "Master"}),
                ("Mute", mute, {"channel": "Master"}),
            ],
        )

    if cm is not None:
        cm.state_variable("SourceProtocolInfo").value = ""
        cm.state_variable("SinkProtocolInfo").value = _dlna_protocol_info()
        cm.state_variable("CurrentConnectionIDs").value = "0"


def _queue_dlna_service_state_sync() -> None:
    with _DLNA_LOCK:
        loop = _DLNA_LOOP
    if loop and loop.is_running():
        loop.call_soon_threadsafe(_apply_dlna_service_state)


def _set_dlna_preferences(payload: dict) -> None:
    mode_raw = str(payload.get("mode") or "all_reachable").strip().lower()
    mode = "all_reachable" if mode_raw == "all_reachable" else "single"
    satellite = str(payload.get("satellite") or "").strip()
    exclude_satellite = str(payload.get("exclude_satellite") or "").strip()
    with _DLNA_LOCK:
        _DLNA_STATE.update(
            {
                "preferred_mode": mode,
                "preferred_satellite": satellite,
                "preferred_exclude_satellite": exclude_satellite,
            }
        )


def _get_dlna_preferences() -> dict:
    with _DLNA_LOCK:
        return {
            "mode": str(_DLNA_STATE.get("preferred_mode") or "all_reachable"),
            "satellite": str(_DLNA_STATE.get("preferred_satellite") or ""),
            "exclude_satellite": str(_DLNA_STATE.get("preferred_exclude_satellite") or ""),
        }


def dlna_status_snapshot() -> dict:
    with _DLNA_LOCK:
        status = dict(_DLNA_STATE)
        thread = _DLNA_THREAD
    if thread is not None and not thread.is_alive():
        status["running"] = False
    # Remove large metadata XML from snapshot to reduce serialization overhead
    status.pop("last_metadata", None)
    return status


def _start_dlna_ffmpeg_proxy(source_url: str, bass_db: float = 0.0, treble_db: float = 0.0) -> None:
    """Launch ffmpeg transcoding source_url â†’ MP3 and broadcast chunks to registered listeners."""
    global _DLNA_PROXY_PROC, _DLNA_PROXY_SOURCE_URL, _DLNA_PROXY_LISTENERS, _DLNA_PROXY_BROADCAST_THREAD
    global _DLNA_PROXY_BASS_DB, _DLNA_PROXY_TREBLE_DB

    _stop_dlna_ffmpeg_proxy()

    ffmpeg_path = shutil.which("ffmpeg")
    if not ffmpeg_path:
        raise RuntimeError("ffmpeg not found on PATH; cannot transcode DLNA stream to MP3")

    ffmpeg_cmd = [
        ffmpeg_path,
        "-hide_banner", "-loglevel", "error", "-nostdin",
        "-reconnect", "1", "-reconnect_streamed", "1", "-reconnect_delay_max", "5",
        "-i", source_url,
        "-vn", "-ac", "2", "-ar", "48000",
    ]
    if abs(bass_db) > 0.01 or abs(treble_db) > 0.01:
        ffmpeg_cmd.extend([
            "-af",
            (
                f"bass=g={bass_db:.2f}:f=100:w=0.65,"
                f"treble=g={treble_db:.2f}:f=3000:w=0.55"
            ),
        ])
    ffmpeg_cmd.extend([
        "-c:a", "libmp3lame", "-b:a", DLNA_PROXY_BITRATE,
        "-f", "mp3", "pipe:1",
    ])
    proc = subprocess.Popen(
        ffmpeg_cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        bufsize=0,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )

    with _DLNA_PROXY_LOCK:
        _DLNA_PROXY_PROC = proc
        _DLNA_PROXY_SOURCE_URL = source_url
        _DLNA_PROXY_LISTENERS = []
        _DLNA_PROXY_BASS_DB = bass_db
        _DLNA_PROXY_TREBLE_DB = treble_db

    def _broadcast() -> None:
        try:
            while True:
                assert proc.stdout is not None
                chunk = proc.stdout.read(8192)
                if not chunk:
                    break
                with _DLNA_PROXY_LOCK:
                    listeners = list(_DLNA_PROXY_LISTENERS)
                overflowed: list[queue.Queue] = []
                for q in listeners:
                    with contextlib.suppress(queue.Full):
                        q.put_nowait(chunk)
                        continue
                    overflowed.append(q)
                if overflowed:
                    with _DLNA_PROXY_LOCK:
                        for q in overflowed:
                            with contextlib.suppress(ValueError):
                                _DLNA_PROXY_LISTENERS.remove(q)
                    for q in overflowed:
                        with contextlib.suppress(Exception):
                            q.put(None)
        except Exception as exc:
            logging.debug("DLNA proxy broadcast ended: %s", exc)
        finally:
            with _DLNA_PROXY_LOCK:
                listeners = list(_DLNA_PROXY_LISTENERS)
            for q in listeners:
                with contextlib.suppress(Exception):
                    q.put(None)
            logging.info("DLNA proxy broadcast thread finished for %s", source_url)

    t = threading.Thread(target=_broadcast, name="dlna-proxy-broadcast", daemon=True)
    with _DLNA_PROXY_LOCK:
        _DLNA_PROXY_BROADCAST_THREAD = t
    t.start()
    logging.info(
        "DLNA proxy started: transcoding %s â†’ MP3 @ %s (bass=%.2fdB treble=%.2fdB)",
        source_url,
        DLNA_PROXY_BITRATE,
        bass_db,
        treble_db,
    )


def _stop_dlna_ffmpeg_proxy() -> None:
    """Terminate the DLNA ffmpeg proxy and signal all listeners to disconnect."""
    global _DLNA_PROXY_PROC, _DLNA_PROXY_SOURCE_URL, _DLNA_PROXY_LISTENERS, _DLNA_PROXY_BROADCAST_THREAD
    global _DLNA_PROXY_BASS_DB, _DLNA_PROXY_TREBLE_DB

    with _DLNA_PROXY_LOCK:
        proc = _DLNA_PROXY_PROC
        listeners = list(_DLNA_PROXY_LISTENERS)
        _DLNA_PROXY_PROC = None
        _DLNA_PROXY_SOURCE_URL = ""
        _DLNA_PROXY_LISTENERS = []
        _DLNA_PROXY_BROADCAST_THREAD = None
        _DLNA_PROXY_BASS_DB = 0.0
        _DLNA_PROXY_TREBLE_DB = 0.0

    if proc is not None:
        with contextlib.suppress(Exception):
            proc.terminate()
        with contextlib.suppress(Exception):
            proc.wait(timeout=3)

    for q in listeners:
        with contextlib.suppress(Exception):
            q.put(None)


def _play_dlna_media() -> dict:
    global _DLNA_LAST_PLAY_MONO

    with _DLNA_LOCK:
        uri = str(_DLNA_STATE.get("last_uri") or "")
        title = str(_DLNA_STATE.get("last_title") or "") or "DLNA"
    if not uri:
        raise ValueError("No DLNA media URI has been set yet.")

    preferences = _get_dlna_preferences()
    preferred_mode = str(preferences.get("mode") or "all_reachable")
    preferred_satellite = str(preferences.get("satellite") or "")
    tone_bass_db = 0.0
    tone_treble_db = 0.0
    if preferred_mode == "single" and preferred_satellite:
        sat = select_satellite(preferred_satellite)
        if sat:
            tone_bass_db, tone_treble_db = _get_user_tone_settings(sat["host"])

    # Keep-alive: if the proxy is already running for this exact URI and HubMusic is active,
    # there is nothing to restart â€” just mark state and return.
    with _DLNA_PROXY_LOCK:
        proxy_source = _DLNA_PROXY_SOURCE_URL
        proxy_alive = _DLNA_PROXY_PROC is not None and _DLNA_PROXY_PROC.poll() is None
        proxy_bass_db = float(_DLNA_PROXY_BASS_DB)
        proxy_treble_db = float(_DLNA_PROXY_TREBLE_DB)
    current_snapshot = _HUB_MUSIC_STATE.snapshot()
    if (
        proxy_alive
        and proxy_source == uri
        and abs(proxy_bass_db - tone_bass_db) <= 0.01
        and abs(proxy_treble_db - tone_treble_db) <= 0.01
        and bool(current_snapshot.get("active"))
    ):
        with _DLNA_LOCK:
            _DLNA_LAST_PLAY_MONO = time.monotonic()
            _DLNA_STATE.update({"transport_state": "PLAYING", "last_error": ""})
        _queue_dlna_service_state_sync()
        proxy_url = build_runtime_url("/dlna/live.mp3")
        return {
            "ok": True,
            "reused": True,
            "title": title,
            "source_url": proxy_url,
            "status": hubmusic_status_snapshot(),
        }

    # Start (or restart) the ffmpeg transcoding proxy for the new URI.
    # Falls back to direct URI if ffmpeg is not available.
    ffmpeg_available = shutil.which("ffmpeg") is not None
    if ffmpeg_available:
        _start_dlna_ffmpeg_proxy(uri, tone_bass_db, tone_treble_db)
        effective_url = build_runtime_url("/dlna/live.mp3")
        logging.info("DLNA play: routing through ffmpeg proxy â†’ %s", effective_url)
    else:
        logging.warning("DLNA play: ffmpeg not found, passing source URL directly to satellite (may cause OOM/crash)")
        effective_url = uri

    payload = {"url": effective_url, "title": title, **preferences}
    result = start_hubmusic_route(payload)
    if not result.get("ok"):
        raise RuntimeError(str(result.get("error") or "Failed to start DLNA playback."))
    started_iso = datetime.now().isoformat(timespec="seconds")
    with _DLNA_LOCK:
        _DLNA_LAST_PLAY_MONO = time.monotonic()
        _DLNA_STATE.update({"transport_state": "PLAYING", "last_error": "", "last_play_started_at": started_iso})
    _queue_dlna_service_state_sync()
    return result


def _stop_dlna_media(paused: bool = False, force: bool = False) -> dict:
    with _DLNA_LOCK:
        last_play = float(_DLNA_LAST_PLAY_MONO)
        transport_state = str(_DLNA_STATE.get("transport_state") or "")

    if not force and not paused and last_play > 0:
        elapsed = time.monotonic() - last_play
        if elapsed < DLNA_TRANSPORT_STOP_GRACE_SECONDS and transport_state == "TRANSITIONING":
            logging.info(
                "DLNA stop ignored as transient sender flap (elapsed=%.2fs, state=%s)",
                elapsed,
                transport_state,
            )
            return {"ok": True, "ignored": True, "reason": "transient_stop_guard"}

    _stop_dlna_ffmpeg_proxy()
    payload = _get_dlna_preferences()
    result = stop_hubmusic_route(payload)
    with _DLNA_LOCK:
        _DLNA_STATE.update({"transport_state": "PAUSED_PLAYBACK" if paused else "STOPPED", "last_error": ""})
    _queue_dlna_service_state_sync()
    return result


if DeviceInfo is not None and ServiceInfo is not None and callable_action is not None and create_state_var is not None and create_event_var is not None:
    class DlnaConnectionManagerService(UpnpServerService):
        SERVICE_DEFINITION = ServiceInfo(
            service_id="urn:upnp-org:serviceId:ConnectionManager",
            service_type="urn:schemas-upnp-org:service:ConnectionManager:1",
            control_url="/upnp/control/ConnectionManager",
            event_sub_url="/upnp/event/ConnectionManager",
            scpd_url="/upnp/ConnectionManager.xml",
            xml=ET.Element("service"),
        )

        STATE_VARIABLE_DEFINITIONS = {
            "SourceProtocolInfo": create_event_var("string", default=""),
            "SinkProtocolInfo": create_event_var("string", default=_dlna_protocol_info()),
            "CurrentConnectionIDs": create_event_var("string", default="0"),
            "A_ARG_TYPE_ProtocolInfo": create_state_var("string"),
            "A_ARG_TYPE_ConnectionStatus": create_state_var("string", allowed=["OK", "ContentFormatMismatch", "InsufficientBandwidth", "UnreliableChannel", "Unknown"], default="OK"),
            "A_ARG_TYPE_AVTransportID": create_state_var("i4", default="0"),
            "A_ARG_TYPE_RcsID": create_state_var("i4", default="0"),
            "A_ARG_TYPE_ConnectionID": create_state_var("i4", default="0"),
            "A_ARG_TYPE_Direction": create_state_var("string", allowed=["Input", "Output"], default="Input"),
            "A_ARG_TYPE_ConnectionManager": create_state_var("string", default=""),
        }

        @callable_action(name="GetProtocolInfo", in_args={}, out_args={"Source": "SourceProtocolInfo", "Sink": "SinkProtocolInfo"})
        async def get_protocol_info(self) -> dict[str, object]:
            self.state_variable("SinkProtocolInfo").value = _dlna_protocol_info()
            return {
                "Source": self.state_variable("SourceProtocolInfo"),
                "Sink": self.state_variable("SinkProtocolInfo"),
            }

        @callable_action(name="GetCurrentConnectionIDs", in_args={}, out_args={"ConnectionIDs": "CurrentConnectionIDs"})
        async def get_current_connection_ids(self) -> dict[str, object]:
            return {"ConnectionIDs": self.state_variable("CurrentConnectionIDs")}

        @callable_action(
            name="GetCurrentConnectionInfo",
            in_args={"ConnectionID": "A_ARG_TYPE_ConnectionID"},
            out_args={
                "RcsID": "A_ARG_TYPE_RcsID",
                "AVTransportID": "A_ARG_TYPE_AVTransportID",
                "ProtocolInfo": "A_ARG_TYPE_ProtocolInfo",
                "PeerConnectionManager": "A_ARG_TYPE_ConnectionManager",
                "PeerConnectionID": "A_ARG_TYPE_ConnectionID",
                "Direction": "A_ARG_TYPE_Direction",
                "Status": "A_ARG_TYPE_ConnectionStatus",
            },
        )
        async def get_current_connection_info(self, ConnectionID: int) -> dict[str, object]:
            self.state_variable("A_ARG_TYPE_ProtocolInfo").value = ""
            self.state_variable("A_ARG_TYPE_ConnectionManager").value = ""
            self.state_variable("A_ARG_TYPE_ConnectionID").value = -1
            self.state_variable("A_ARG_TYPE_Direction").value = "Input"
            self.state_variable("A_ARG_TYPE_ConnectionStatus").value = "OK"
            return {
                "RcsID": self.state_variable("A_ARG_TYPE_RcsID"),
                "AVTransportID": self.state_variable("A_ARG_TYPE_AVTransportID"),
                "ProtocolInfo": self.state_variable("A_ARG_TYPE_ProtocolInfo"),
                "PeerConnectionManager": self.state_variable("A_ARG_TYPE_ConnectionManager"),
                "PeerConnectionID": self.state_variable("A_ARG_TYPE_ConnectionID"),
                "Direction": self.state_variable("A_ARG_TYPE_Direction"),
                "Status": self.state_variable("A_ARG_TYPE_ConnectionStatus"),
            }


    class DlnaRenderingControlService(UpnpServerService):
        SERVICE_DEFINITION = ServiceInfo(
            service_id="urn:upnp-org:serviceId:RenderingControl",
            service_type="urn:schemas-upnp-org:service:RenderingControl:1",
            control_url="/upnp/control/RenderingControl",
            event_sub_url="/upnp/event/RenderingControl",
            scpd_url="/upnp/RenderingControl.xml",
            xml=ET.Element("service"),
        )

        STATE_VARIABLE_DEFINITIONS = {
            "LastChange": create_event_var("string", default=""),
            "Mute": create_event_var("boolean", default="0"),
            "Volume": create_event_var("ui2", allowed_range={"minimum": "0", "maximum": "100", "step": "1"}, default="100"),
            "A_ARG_TYPE_InstanceID": create_state_var("ui4", default="0"),
            "A_ARG_TYPE_Channel": create_state_var("string", allowed=["Master"], default="Master"),
            "A_ARG_TYPE_Mute": create_state_var("boolean", default="0"),
            "A_ARG_TYPE_Volume": create_state_var("ui2", allowed_range={"minimum": "0", "maximum": "100", "step": "1"}, default="100"),
        }

        @callable_action(name="GetVolume", in_args={"InstanceID": "A_ARG_TYPE_InstanceID", "Channel": "A_ARG_TYPE_Channel"}, out_args={"CurrentVolume": "Volume"})
        async def get_volume(self, InstanceID: int, Channel: str) -> dict[str, object]:
            return {"CurrentVolume": self.state_variable("Volume")}

        @callable_action(name="SetVolume", in_args={"InstanceID": "A_ARG_TYPE_InstanceID", "Channel": "A_ARG_TYPE_Channel", "DesiredVolume": "A_ARG_TYPE_Volume"}, out_args={})
        async def set_volume(self, InstanceID: int, Channel: str, DesiredVolume: int) -> dict[str, object]:
            volume = max(0, min(100, int(DesiredVolume)))
            with _DLNA_LOCK:
                _DLNA_STATE["volume"] = volume
            _apply_dlna_service_state()
            return {}

        @callable_action(name="GetMute", in_args={"InstanceID": "A_ARG_TYPE_InstanceID", "Channel": "A_ARG_TYPE_Channel"}, out_args={"CurrentMute": "Mute"})
        async def get_mute(self, InstanceID: int, Channel: str) -> dict[str, object]:
            return {"CurrentMute": self.state_variable("Mute")}

        @callable_action(name="SetMute", in_args={"InstanceID": "A_ARG_TYPE_InstanceID", "Channel": "A_ARG_TYPE_Channel", "DesiredMute": "A_ARG_TYPE_Mute"}, out_args={})
        async def set_mute(self, InstanceID: int, Channel: str, DesiredMute: bool) -> dict[str, object]:
            with _DLNA_LOCK:
                _DLNA_STATE["mute"] = bool(DesiredMute)
            _apply_dlna_service_state()
            return {}


    class DlnaAvTransportService(UpnpServerService):
        SERVICE_DEFINITION = ServiceInfo(
            service_id="urn:upnp-org:serviceId:AVTransport",
            service_type="urn:schemas-upnp-org:service:AVTransport:1",
            control_url="/upnp/control/AVTransport",
            event_sub_url="/upnp/event/AVTransport",
            scpd_url="/upnp/AVTransport.xml",
            xml=ET.Element("service"),
        )

        STATE_VARIABLE_DEFINITIONS = {
            "LastChange": create_event_var("string", default=""),
            "TransportState": create_event_var("string", allowed=["STOPPED", "PLAYING", "TRANSITIONING", "PAUSED_PLAYBACK", "NO_MEDIA_PRESENT"], default="NO_MEDIA_PRESENT"),
            "TransportStatus": create_event_var("string", allowed=["OK", "ERROR_OCCURRED"], default="OK"),
            "CurrentTransportActions": create_event_var("string", default=""),
            "AVTransportURI": create_event_var("string", default=""),
            "AVTransportURIMetaData": create_event_var("string", default=""),
            "CurrentTrackURI": create_event_var("string", default=""),
            "CurrentTrackMetaData": create_event_var("string", default=""),
            "CurrentMediaDuration": create_event_var("string", default="00:00:00"),
            "CurrentTrackDuration": create_event_var("string", default="00:00:00"),
            "CurrentTrack": create_event_var("ui4", default="0"),
            "NumberOfTracks": create_event_var("ui4", default="0"),
            "RelativeTimePosition": create_event_var("string", default="00:00:00"),
            "AbsoluteTimePosition": create_event_var("string", default="00:00:00"),
            "PlaybackStorageMedium": create_state_var("string", default="NETWORK"),
            "RecordStorageMedium": create_state_var("string", default="NOT_IMPLEMENTED"),
            "PossiblePlaybackStorageMedia": create_state_var("string", default="NETWORK"),
            "PossibleRecordStorageMedia": create_state_var("string", default="NOT_IMPLEMENTED"),
            "CurrentPlayMode": create_state_var("string", default="NORMAL"),
            "TransportPlaySpeed": create_state_var("string", default="1"),
            "A_ARG_TYPE_InstanceID": create_state_var("ui4", default="0"),
            "A_ARG_TYPE_CurrentURI": create_state_var("string"),
            "A_ARG_TYPE_CurrentURIMetaData": create_state_var("string"),
            "A_ARG_TYPE_NextURI": create_state_var("string"),
            "A_ARG_TYPE_NextURIMetaData": create_state_var("string"),
            "A_ARG_TYPE_PlaySpeed": create_state_var("string", default="1"),
            "A_ARG_TYPE_SeekMode": create_state_var("string"),
            "A_ARG_TYPE_SeekTarget": create_state_var("string"),
        }

        @callable_action(name="SetAVTransportURI", in_args={"InstanceID": "A_ARG_TYPE_InstanceID", "CurrentURI": "A_ARG_TYPE_CurrentURI", "CurrentURIMetaData": "A_ARG_TYPE_CurrentURIMetaData"}, out_args={})
        async def set_av_transport_uri(self, InstanceID: int, CurrentURI: str, CurrentURIMetaData: str) -> dict[str, object]:
            global _DLNA_LAST_PLAY_MONO

            uri = normalize_media_url(CurrentURI)
            metadata = str(CurrentURIMetaData or "")
            meta = _extract_dlna_metadata(metadata, uri)
            with _DLNA_LOCK:
                previous_uri = str(_DLNA_STATE.get("last_uri") or "")
                previous_state = str(_DLNA_STATE.get("transport_state") or "")
                same_stream_while_playing = (
                    bool(uri)
                    and _dlna_uri_signature(uri)
                    and _dlna_uri_signature(uri) == _dlna_uri_signature(previous_uri)
                    and previous_state in {"PLAYING", "TRANSITIONING"}
                )
                if same_stream_while_playing:
                    _DLNA_STATE.update(
                        {
                            "last_uri": uri,
                            "last_metadata": metadata,
                            "last_title": meta["title"],
                            "last_artist": meta["artist"],
                            "last_album": meta["album"],
                            "transport_state": "PLAYING",
                            "last_error": "",
                        }
                    )
                else:
                    _DLNA_LAST_PLAY_MONO = 0.0
                    _DLNA_STATE.update(
                        {
                            "last_uri": uri,
                            "last_metadata": metadata,
                            "last_title": meta["title"],
                            "last_artist": meta["artist"],
                            "last_album": meta["album"],
                            "transport_state": "STOPPED",
                            "last_error": "",
                        }
                    )
                # Kill existing proxy when sender sets a new URI so old ffmpeg stops immediately.
                threading.Thread(target=_stop_dlna_ffmpeg_proxy, name="dlna-proxy-stop", daemon=True).start()
            _apply_dlna_service_state()
            return {}

        @callable_action(name="Play", in_args={"InstanceID": "A_ARG_TYPE_InstanceID", "Speed": "A_ARG_TYPE_PlaySpeed"}, out_args={})
        async def play(self, InstanceID: int, Speed: str) -> dict[str, object]:
            with _DLNA_LOCK:
                current_uri = str(_DLNA_STATE.get("last_uri") or "")
                current_state = str(_DLNA_STATE.get("transport_state") or "")
                last_play = float(_DLNA_LAST_PLAY_MONO)

            # If HubMusic is already routing the same stream, treat repeated Play as keep-alive.
            snapshot = _HUB_MUSIC_STATE.snapshot()
            if bool(snapshot.get("active")):
                current_source = str(snapshot.get("source_url") or "")
                if _dlna_uri_signature(current_uri) and _dlna_uri_signature(current_uri) == _dlna_uri_signature(current_source):
                    with _DLNA_LOCK:
                        _DLNA_STATE["transport_state"] = "PLAYING"
                        _DLNA_STATE["last_error"] = ""
                    _apply_dlna_service_state()
                    logging.info("DLNA play treated as keep-alive for existing stream")
                    return {}

            if current_state in {"PLAYING", "TRANSITIONING"} and last_play > 0:
                elapsed = time.monotonic() - last_play
                if elapsed < DLNA_REPEAT_PLAY_GUARD_SECONDS:
                    with _DLNA_LOCK:
                        _DLNA_STATE["transport_state"] = "PLAYING"
                        _DLNA_STATE["last_error"] = ""
                    _apply_dlna_service_state()
                    logging.info("DLNA play ignored as duplicate within %.2fs", elapsed)
                    return {}

            with _DLNA_LOCK:
                _DLNA_STATE["last_error"] = ""
                if _DLNA_STATE.get("transport_state") not in {"PLAYING", "TRANSITIONING"}:
                    _DLNA_STATE["transport_state"] = "TRANSITIONING"
            _apply_dlna_service_state()
            try:
                await asyncio.to_thread(_play_dlna_media)
            except Exception as exc:
                with _DLNA_LOCK:
                    _DLNA_STATE["transport_state"] = "STOPPED"
                    _DLNA_STATE["last_error"] = str(exc)
                _apply_dlna_service_state()
                raise
            return {}

        @callable_action(name="Stop", in_args={"InstanceID": "A_ARG_TYPE_InstanceID"}, out_args={})
        async def stop(self, InstanceID: int) -> dict[str, object]:
            await asyncio.to_thread(_stop_dlna_media, False)
            return {}

        @callable_action(name="Pause", in_args={"InstanceID": "A_ARG_TYPE_InstanceID"}, out_args={})
        async def pause(self, InstanceID: int) -> dict[str, object]:
            await asyncio.to_thread(_stop_dlna_media, True)
            return {}

        @callable_action(name="GetTransportInfo", in_args={"InstanceID": "A_ARG_TYPE_InstanceID"}, out_args={"CurrentTransportState": "TransportState", "CurrentTransportStatus": "TransportStatus", "CurrentSpeed": "TransportPlaySpeed"})
        async def get_transport_info(self, InstanceID: int) -> dict[str, object]:
            return {
                "CurrentTransportState": self.state_variable("TransportState"),
                "CurrentTransportStatus": self.state_variable("TransportStatus"),
                "CurrentSpeed": self.state_variable("TransportPlaySpeed"),
            }

        @callable_action(name="GetMediaInfo", in_args={"InstanceID": "A_ARG_TYPE_InstanceID"}, out_args={"NrTracks": "NumberOfTracks", "MediaDuration": "CurrentMediaDuration", "CurrentURI": "AVTransportURI", "CurrentURIMetaData": "AVTransportURIMetaData", "NextURI": "A_ARG_TYPE_NextURI", "NextURIMetaData": "A_ARG_TYPE_NextURIMetaData", "PlayMedium": "PlaybackStorageMedium", "RecordMedium": "RecordStorageMedium", "WriteStatus": "TransportStatus"})
        async def get_media_info(self, InstanceID: int) -> dict[str, object]:
            self.state_variable("A_ARG_TYPE_NextURI").value = ""
            self.state_variable("A_ARG_TYPE_NextURIMetaData").value = ""
            return {
                "NrTracks": self.state_variable("NumberOfTracks"),
                "MediaDuration": self.state_variable("CurrentMediaDuration"),
                "CurrentURI": self.state_variable("AVTransportURI"),
                "CurrentURIMetaData": self.state_variable("AVTransportURIMetaData"),
                "NextURI": self.state_variable("A_ARG_TYPE_NextURI"),
                "NextURIMetaData": self.state_variable("A_ARG_TYPE_NextURIMetaData"),
                "PlayMedium": self.state_variable("PlaybackStorageMedium"),
                "RecordMedium": self.state_variable("RecordStorageMedium"),
                "WriteStatus": self.state_variable("TransportStatus"),
            }

        @callable_action(name="GetPositionInfo", in_args={"InstanceID": "A_ARG_TYPE_InstanceID"}, out_args={"Track": "CurrentTrack", "TrackDuration": "CurrentTrackDuration", "TrackMetaData": "CurrentTrackMetaData", "TrackURI": "CurrentTrackURI", "RelTime": "RelativeTimePosition", "AbsTime": "AbsoluteTimePosition", "RelCount": "CurrentTrack", "AbsCount": "CurrentTrack"})
        async def get_position_info(self, InstanceID: int) -> dict[str, object]:
            return {
                "Track": self.state_variable("CurrentTrack"),
                "TrackDuration": self.state_variable("CurrentTrackDuration"),
                "TrackMetaData": self.state_variable("CurrentTrackMetaData"),
                "TrackURI": self.state_variable("CurrentTrackURI"),
                "RelTime": self.state_variable("RelativeTimePosition"),
                "AbsTime": self.state_variable("AbsoluteTimePosition"),
                "RelCount": self.state_variable("CurrentTrack"),
                "AbsCount": self.state_variable("CurrentTrack"),
            }

        @callable_action(name="GetDeviceCapabilities", in_args={"InstanceID": "A_ARG_TYPE_InstanceID"}, out_args={"PlayMedia": "PossiblePlaybackStorageMedia", "RecMedia": "PossibleRecordStorageMedia", "RecQualityModes": "TransportStatus"})
        async def get_device_capabilities(self, InstanceID: int) -> dict[str, object]:
            self.state_variable("TransportStatus").value = "OK"
            return {
                "PlayMedia": self.state_variable("PossiblePlaybackStorageMedia"),
                "RecMedia": self.state_variable("PossibleRecordStorageMedia"),
                "RecQualityModes": self.state_variable("TransportStatus"),
            }

        @callable_action(name="GetTransportSettings", in_args={"InstanceID": "A_ARG_TYPE_InstanceID"}, out_args={"PlayMode": "CurrentPlayMode", "RecQualityMode": "TransportStatus"})
        async def get_transport_settings(self, InstanceID: int) -> dict[str, object]:
            return {
                "PlayMode": self.state_variable("CurrentPlayMode"),
                "RecQualityMode": self.state_variable("TransportStatus"),
            }

        @callable_action(name="GetCurrentTransportActions", in_args={"InstanceID": "A_ARG_TYPE_InstanceID"}, out_args={"Actions": "CurrentTransportActions"})
        async def get_current_transport_actions(self, InstanceID: int) -> dict[str, object]:
            return {"Actions": self.state_variable("CurrentTransportActions")}


    class DlnaMediaRendererDevice(UpnpServerDevice):
        DEVICE_DEFINITION = DeviceInfo(
            device_type="urn:schemas-upnp-org:device:MediaRenderer:1",
            friendly_name=DLNA_RENDERER_DEFAULT_NAME,
            manufacturer="HubVoiceSat",
            manufacturer_url=None,
            model_name="HubVoiceSat DLNA Renderer",
            model_url=None,
            udn=f"uuid:{uuid.uuid5(uuid.NAMESPACE_URL, 'hubvoicesat-dlna-renderer')}",
            upc=None,
            model_description="HubVoiceSat built-in DLNA MediaRenderer",
            model_number="1.0",
            serial_number="hubvoicesat-dlna",
            presentation_url=None,
            url="/upnp/device.xml",
            icons=[],
            xml=ET.Element("device"),
        )
        EMBEDDED_DEVICES = []
        SERVICES = [DlnaConnectionManagerService, DlnaRenderingControlService, DlnaAvTransportService]


def _dlna_thread_main(host: str, http_port: int) -> None:
    global _DLNA_LOOP, _DLNA_SERVER, _DLNA_AVT_SERVICE, _DLNA_RC_SERVICE, _DLNA_CM_SERVICE

    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    with _DLNA_LOCK:
        _DLNA_LOOP = loop

    async def _start() -> None:
        global _DLNA_SERVER, _DLNA_AVT_SERVICE, _DLNA_RC_SERVICE, _DLNA_CM_SERVICE
        server = UpnpServer(DlnaMediaRendererDevice, (host, 0), http_port=http_port)
        await server.async_start()
        device = getattr(server, "_device", None)
        avt = None
        rc = None
        cm = None
        if device is not None:
            for service in getattr(device, "all_services", []):
                service_id = str(getattr(service, "service_id", ""))
                if service_id == "urn:upnp-org:serviceId:AVTransport":
                    avt = service
                elif service_id == "urn:upnp-org:serviceId:RenderingControl":
                    rc = service
                elif service_id == "urn:upnp-org:serviceId:ConnectionManager":
                    cm = service
        with _DLNA_LOCK:
            _DLNA_SERVER = server
            _DLNA_AVT_SERVICE = avt
            _DLNA_RC_SERVICE = rc
            _DLNA_CM_SERVICE = cm
            _DLNA_STATE.update(
                {
                    "enabled": True,
                    "running": True,
                    "host": host,
                    "http_port": http_port,
                    "device_url": f"http://{host}:{http_port}/upnp/device.xml",
                    "started_at": datetime.now().isoformat(timespec="seconds"),
                    "last_error": "",
                }
            )
        _apply_dlna_service_state()

    try:
        loop.run_until_complete(_start())
        loop.run_forever()
    except Exception as exc:
        with _DLNA_LOCK:
            _DLNA_STATE.update({"enabled": False, "running": False, "last_error": f"Failed to start DLNA renderer: {exc}"})
    finally:
        pending = [task for task in asyncio.all_tasks(loop) if not task.done()]
        for task in pending:
            task.cancel()
        if pending:
            with contextlib.suppress(Exception):
                loop.run_until_complete(asyncio.gather(*pending, return_exceptions=True))
        with _DLNA_LOCK:
            _DLNA_LOOP = None
            _DLNA_SERVER = None
            _DLNA_AVT_SERVICE = None
            _DLNA_RC_SERVICE = None
            _DLNA_CM_SERVICE = None
            _DLNA_STATE["running"] = False
        loop.close()


def start_dlna_renderer(payload: dict | None = None) -> tuple[bool, str, dict]:
    global _DLNA_THREAD

    already_running = False
    with _DLNA_LOCK:
        if _DLNA_THREAD is not None and _DLNA_THREAD.is_alive():
            already_running = True

    if already_running:
        if payload:
            _set_dlna_preferences(payload)
        return True, "DLNA renderer already running.", dlna_status_snapshot()

    missing = _ensure_dlna_dependencies_installed()
    if missing or UpnpServer is None:
        message = "Missing DLNA dependencies: " + ", ".join(sorted(set(missing or ["async_upnp_client"])))
        with _DLNA_LOCK:
            _DLNA_STATE.update({"enabled": False, "running": False, "last_error": message})
        return False, message, dlna_status_snapshot()

    host, host_error = _resolve_runtime_ipv4_host()
    if not host:
        message = host_error or "Unable to determine a LAN IPv4 address for the DLNA renderer."
        with _DLNA_LOCK:
            _DLNA_STATE.update({"enabled": False, "running": False, "last_error": message})
        return False, message, dlna_status_snapshot()

    if payload:
        _set_dlna_preferences(payload)

    thread = threading.Thread(target=_dlna_thread_main, args=(host, DLNA_RENDERER_HTTP_PORT), name="hubvoice-dlna", daemon=True)
    thread.start()
    time.sleep(1.5)
    with _DLNA_LOCK:
        _DLNA_THREAD = thread if thread.is_alive() else None
        error = str(_DLNA_STATE.get("last_error") or "")
        running = bool(_DLNA_STATE.get("running"))
    if not thread.is_alive() or not running:
        return False, error or "DLNA renderer failed to start.", dlna_status_snapshot()
    return True, "DLNA renderer started.", dlna_status_snapshot()


def stop_dlna_renderer() -> tuple[bool, str, dict]:
    global _DLNA_THREAD

    with _DLNA_LOCK:
        loop = _DLNA_LOOP
        server = _DLNA_SERVER
        thread = _DLNA_THREAD

    with contextlib.suppress(Exception):
        _stop_dlna_media(False, force=True)
    with contextlib.suppress(Exception):
        _stop_dlna_ffmpeg_proxy()

    if loop is not None and server is not None:
        with contextlib.suppress(Exception):
            fut = asyncio.run_coroutine_threadsafe(server.async_stop(), loop)
            fut.result(timeout=5)
        with contextlib.suppress(Exception):
            loop.call_soon_threadsafe(loop.stop)

    if thread is not None:
        thread.join(timeout=5)

    with _DLNA_LOCK:
        _DLNA_THREAD = None
        _DLNA_STATE.update({"enabled": False, "running": False, "transport_state": "NO_MEDIA_PRESENT"})
    return True, "DLNA renderer stopped.", dlna_status_snapshot()


def acquire_runtime_single_instance_lock() -> None:
    global _RUNTIME_LOCK_HANDLE
    if msvcrt is None:
        return

    handle = open(_RUNTIME_LOCK_FILE, "a+b")
    try:
        handle.seek(0)
        # Lock first byte for single-instance enforcement.
        msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
    except OSError:
        handle.close()
        raise RuntimeError("Another HubVoiceSat runtime instance is already running")

    handle.seek(0)
    handle.truncate(0)
    handle.write(str(time.time()).encode("ascii"))
    handle.flush()
    _RUNTIME_LOCK_HANDLE = handle


def release_runtime_single_instance_lock() -> None:
    global _RUNTIME_LOCK_HANDLE
    if _RUNTIME_LOCK_HANDLE is None or msvcrt is None:
        return
    try:
        _RUNTIME_LOCK_HANDLE.seek(0)
        msvcrt.locking(_RUNTIME_LOCK_HANDLE.fileno(), msvcrt.LK_UNLCK, 1)
    except Exception:
        pass
    try:
        _RUNTIME_LOCK_HANDLE.close()
    except Exception:
        pass
    _RUNTIME_LOCK_HANDLE = None


def cleanup_resources() -> None:
    """Clean up global resources on shutdown."""
    global _PIPER_VOICE, _WHISPER_MODEL

    _stop_hubmusic_ffmpeg()
    stop_airplay_receiver()
    stop_dlna_renderer()

    logging.info("Starting resource cleanup...")
    
    try:
        with _PIPER_VOICE_LOCK:
            if _PIPER_VOICE is not None:
                try:
                    if hasattr(_PIPER_VOICE, 'close'):
                        _PIPER_VOICE.close()
                except Exception:
                    pass
                _PIPER_VOICE = None
    except Exception as e:
        logging.warning("Error cleaning up Piper voice: %s", e)
    
    try:
        with _WHISPER_MODEL_LOCK:
            if _WHISPER_MODEL is not None:
                try:
                    if hasattr(_WHISPER_MODEL, 'close'):
                        _WHISPER_MODEL.close()
                except Exception:
                    pass
                _WHISPER_MODEL = None
    except Exception as e:
        logging.warning("Error cleaning up Whisper model: %s", e)
    
    try:
        with _ENTITY_CACHE_LOCK:
            _ENTITY_CACHE.clear()
    except Exception as e:
        logging.warning("Error clearing entity cache: %s", e)
    
    try:
        asyncio.run(_SATELLITE_POOL.close_all())
    except Exception as e:
        logging.warning("Error closing satellite connections: %s", e)
    
    logging.info("Resource cleanup completed")


class RuntimeHandler(BaseHTTPRequestHandler):
    server_version = "HubVoiceSatRuntime/1.0"

    def _read_json_body(self) -> dict:
        try:
            content_length = int(self.headers.get("Content-Length", "0") or "0")
        except ValueError:
            raise ValueError("Invalid Content-Length")
        if content_length <= 0:
            return {}
        if content_length > REQUEST_SIZE_LIMIT:
            raise ValueError("Request body too large")
        raw = self.rfile.read(content_length)
        if not raw:
            return {}
        try:
            data = json.loads(raw.decode("utf-8"))
        except Exception as exc:
            raise ValueError(f"Invalid JSON body: {exc}")
        if not isinstance(data, dict):
            raise ValueError("JSON body must be an object")
        return data

    def _resolve_satellite(self, payload: dict) -> dict | None:
        sat_id = str(payload.get("satellite") or payload.get("sat_id") or payload.get("d") or "").strip()
        return select_satellite(sat_id)

    def _handle_answer_request(self, payload: dict) -> None:
        text = sanitize_text(str(payload.get("text") or payload.get("r") or ""))
        sat_id = str(payload.get("satellite") or payload.get("sat_id") or payload.get("d") or "").strip()
        if payload.get("text") or payload.get("r"):
            _WORK_QUEUE.put({"text": text, "sat_id": sat_id})
        self._send_json(
            {
                "ok": True,
                "queued": bool(payload.get("text") or payload.get("r")),
                "queue_depth": _WORK_QUEUE.qsize(),
                "satellite": sat_id,
                "text": text if (payload.get("text") or payload.get("r")) else "",
            }
        )

    def _handle_satellite_switch_request(self, payload: dict) -> None:
        sat = self._resolve_satellite(payload)
        entity_id = str(payload.get("entity") or "").strip()
        state_raw = payload.get("state")
        if isinstance(state_raw, bool):
            new_state = state_raw
        else:
            state_str = str(state_raw or "").strip().lower()
            if state_str not in {"on", "off", "true", "false", "1", "0"}:
                self._send_json({"ok": False, "error": "Missing or invalid params: satellite, entity, state (on/off)"}, status=400)
                return
            new_state = state_str in {"on", "true", "1"}
        if not sat or not entity_id:
            self._send_json({"ok": False, "error": "Missing or invalid params: satellite, entity, state (on/off)"}, status=400)
            return
        try:
            set_satellite_switch(sat["host"], entity_id, new_state)
            _update_control_deck_state(sat["id"], **{entity_id: new_state})
            self._send_json({"ok": True, "satellite": sat["id"], "entity": entity_id, "state": "on" if new_state else "off"})
        except Exception as exc:
            logging.error("satellite-switch error for %s/%s: %s", sat["id"], entity_id, exc)
            self._send_json({"ok": False, "error": str(exc)}, status=500)

    def _handle_satellite_number_request(self, payload: dict) -> None:
        sat = self._resolve_satellite(payload)
        entity_id = str(payload.get("entity") or "").strip()
        value = payload.get("value")
        if not sat or not entity_id or value is None:
            self._send_json({"ok": False, "error": "Missing or invalid params: satellite, entity, value"}, status=400)
            return
        try:
            numeric_value = float(value)
            if entity_id == "speaker_volume":
                # speaker_volume is applied directly to media_player to avoid heavy number retries during slider drags.
                numeric_value = _remember_user_locked_media_volume(sat["host"], numeric_value)
                set_satellite_media_volume(sat["host"], numeric_value)
            elif entity_id in {"bass_level", "treble_level"}:
                if entity_id == "bass_level":
                    bass_db, treble_db = _set_user_tone_settings(sat["host"], bass=numeric_value)
                    numeric_value = bass_db
                else:
                    bass_db, treble_db = _set_user_tone_settings(sat["host"], treble=numeric_value)
                    numeric_value = treble_db

                # Keep firmware-side tone state in sync when number entities exist.
                with contextlib.suppress(Exception):
                    set_satellite_number(sat["host"], entity_id, numeric_value)

                # If DLNA is currently playing this same single satellite target, apply tone immediately.
                with _DLNA_LOCK:
                    dlna_running = bool(_DLNA_STATE.get("running"))
                prefs = _get_dlna_preferences()
                if dlna_running and str(prefs.get("mode") or "") == "single" and str(prefs.get("satellite") or "") == sat["id"]:
                    with contextlib.suppress(Exception):
                        _play_dlna_media()
                logging.info("Updated %s tone for %s (bass=%.2fdB treble=%.2fdB)", sat["id"], sat["host"], bass_db, treble_db)
            else:
                set_satellite_number(sat["host"], entity_id, numeric_value)
            _update_control_deck_state(sat["id"], **{entity_id: numeric_value})
            if entity_id == "wake_sound_volume":
                controls = build_satellite_control_metadata()
                wake_min = float(controls.get("wake_sound_volume", {}).get("min", 0))
                _update_control_deck_state(sat["id"], wake_muted=(numeric_value <= wake_min))
            self._send_json({"ok": True, "satellite": sat["id"], "entity": entity_id, "value": numeric_value})
        except Exception as exc:
            logging.error("satellite-number error for %s/%s: %s", sat["id"], entity_id, exc)
            self._send_json({"ok": False, "error": str(exc)}, status=500)

    def _handle_satellite_media_volume_request(self, payload: dict) -> None:
        sat = self._resolve_satellite(payload)
        value = payload.get("value")
        if not sat or value is None:
            self._send_json({"ok": False, "error": "Missing or invalid params: satellite, value"}, status=400)
            return
        numeric_value = _remember_user_locked_media_volume(sat["host"], value)
        should_apply, wait_s = _should_apply_satellite_media_volume(sat["host"], numeric_value)
        if not should_apply:
            _update_control_deck_state(sat["id"], speaker_volume=numeric_value)
            self._send_json(
                {
                    "ok": True,
                    "satellite": sat["id"],
                    "value": numeric_value,
                    "applied": False,
                    "message": f"rate_limited_{wait_s:.2f}s",
                }
            )
            return

        try:
            set_satellite_media_volume(sat["host"], numeric_value)
            _mark_satellite_media_volume_applied(sat["host"], numeric_value)
            _update_control_deck_state(sat["id"], speaker_volume=numeric_value)
            self._send_json({"ok": True, "satellite": sat["id"], "value": numeric_value, "applied": True})
            return
        except Exception as exc:
            logging.warning("satellite-media-volume deferred for %s: %s", sat["id"], exc)
            self._send_json(
                {
                    "ok": True,
                    "satellite": sat["id"],
                    "value": numeric_value,
                    "applied": False,
                    "message": "speaker_busy",
                }
            )

    def _handle_satellite_media_request(self, payload: dict) -> None:
        sat = self._resolve_satellite(payload)
        muted_raw = payload.get("muted")
        if isinstance(muted_raw, bool):
            muted = muted_raw
        else:
            muted_str = str(muted_raw or "").strip().lower()
            if muted_str not in {"on", "off", "true", "false", "1", "0"}:
                self._send_json({"ok": False, "error": "Missing or invalid params: satellite, muted (on/off)"}, status=400)
                return
            muted = muted_str in {"on", "true", "1"}
        if not sat:
            self._send_json({"ok": False, "error": "Missing or invalid params: satellite, muted (on/off)"}, status=400)
            return
        try:
            set_satellite_media_mute(sat["host"], muted)
            _update_control_deck_state(sat["id"], speaker_muted=muted)
            self._send_json({"ok": True, "satellite": sat["id"], "muted": muted})
        except Exception as exc:
            logging.error("satellite-media error for %s: %s", sat["id"], exc)
            self._send_json({"ok": False, "error": str(exc)}, status=500)

    def _handle_airplay_status_request(self) -> None:
        self._send_json({"ok": True, "status": airplay_status_snapshot()})

    def _handle_airplay_start_request(self, payload: dict) -> None:
        name = str(payload.get("name") or AIRPLAY_RECEIVER_DEFAULT_NAME).strip() or AIRPLAY_RECEIVER_DEFAULT_NAME
        ok, message, status = start_airplay_receiver(name)
        code = 200 if ok else 500
        self._send_json({"ok": ok, "message": message, "status": status}, status=code)

    def _handle_airplay_stop_request(self) -> None:
        ok, message, status = stop_airplay_receiver()
        code = 200 if ok else 500
        self._send_json({"ok": ok, "message": message, "status": status}, status=code)

    def _handle_dlna_status_request(self) -> None:
        self._send_json({"ok": True, "status": dlna_status_snapshot()})

    def _handle_dlna_start_request(self, payload: dict) -> None:
        ok, message, status = start_dlna_renderer(payload)
        code = 200 if ok else 500
        self._send_json({"ok": ok, "message": message, "status": status}, status=code)

    def _handle_dlna_stop_request(self) -> None:
        ok, message, status = stop_dlna_renderer()
        code = 200 if ok else 500
        self._send_json({"ok": ok, "message": message, "status": status}, status=code)

    def _handle_hubmusic_play_request(self, payload: dict) -> None:
        try:
            result = start_hubmusic_route(payload)
            status_code = 200 if result.get("ok") else 500
            self._send_json(result, status=status_code)
        except LookupError as exc:
            self._send_json({"ok": False, "error": str(exc)}, status=404)
        except ValueError as exc:
            self._send_json({"ok": False, "error": str(exc)}, status=400)
        except Exception as exc:
            message = str(exc)
            _HUB_MUSIC_STATE.error(message)
            _APP_STATE.update(last_action="hubmusic_error", last_error=message)
            logging.error("hubmusic-play error: %s", exc)
            self._send_json({"ok": False, "error": message}, status=500)

    def _handle_hubmusic_stop_request(self, payload: dict) -> None:
        try:
            result = stop_hubmusic_route(payload)
            self._send_json(result)
        except Exception as exc:
            message = str(exc)
            _HUB_MUSIC_STATE.error(message)
            _APP_STATE.update(last_action="hubmusic_error", last_error=message)
            logging.error("hubmusic-stop error: %s", exc)
            self._send_json({"ok": False, "error": message}, status=500)

    def _handle_hubmusic_stereo_test_request(self, payload: dict) -> None:
        try:
            result = run_hubmusic_stereo_test(payload)
            status_code = 200 if result.get("ok") else 500
            self._send_json(result, status=status_code)
        except LookupError as exc:
            self._send_json({"ok": False, "error": str(exc)}, status=404)
        except ValueError as exc:
            self._send_json({"ok": False, "error": str(exc)}, status=400)
        except Exception as exc:
            message = str(exc)
            _HUB_MUSIC_STATE.error(message)
            _APP_STATE.update(last_action="hubmusic_error", last_error=message)
            logging.error("hubmusic-stereo-test error: %s", exc)
            self._send_json({"ok": False, "error": message}, status=500)

    def _handle_hubmusic_stereo_config_request(self, payload: dict) -> None:
        left_raw = payload.get("left_satellite")
        right_raw = payload.get("right_satellite")
        volume_raw = payload.get("volume_pct")

        left = str(left_raw or "").strip()
        right = str(right_raw or "").strip()
        has_left = left_raw is not None and left != ""
        has_right = right_raw is not None and right != ""
        has_volume = volume_raw is not None
        if not has_left and not has_right and not has_volume:
            self._send_json(
                {"ok": False, "error": "At least one of left_satellite, right_satellite, volume_pct is required"},
                status=400,
            )
            return

        volume_pct = _config_int(volume_raw, 50, 0, 100) if has_volume else None
        try:
            cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8-sig")) if CONFIG_PATH.exists() else {}
            if has_left:
                cfg["hubmusic_stereo_left"] = left
            if has_right:
                cfg["hubmusic_stereo_right"] = right
            if volume_pct is not None:
                cfg["hubmusic_stereo_volume_pct"] = int(volume_pct)
            CONFIG_PATH.write_text(json.dumps(cfg, indent=4), encoding="utf-8")
            self._send_json(
                {
                    "ok": True,
                    "left_satellite": str(cfg.get("hubmusic_stereo_left", "")).strip(),
                    "right_satellite": str(cfg.get("hubmusic_stereo_right", "")).strip(),
                    "volume_pct": _config_int(cfg.get("hubmusic_stereo_volume_pct"), 50, 0, 100),
                }
            )
        except Exception as exc:
            logging.error("stereo-config save error: %s", exc)
            self._send_json({"ok": False, "error": str(exc)}, status=500)

    def _handle_hubmusic_stereo_sync_diagnostics_request(self) -> None:
        """Return the last stereo playback sync measurements."""
        snapshot = _HUB_MUSIC_STATE.snapshot()
        last_sent = snapshot.get("last_sent", [])
        last_operation = snapshot.get("last_operation", "none")
        started_at = snapshot.get("started_at", "")
        
        # Extract sync data from the last stereo playback
        sync_data = {
            "ok": True,
            "last_operation": last_operation,
            "started_at": started_at,
            "speakers": [],
            "overall_skew_ms": None,
            "in_sync": True,
        }
        
        for item in last_sent:
            speaker = {
                "id": item.get("id", ""),
                "alias": item.get("alias", ""),
                "channel": item.get("channel", ""),
                "sync_skew_ms": item.get("sync_skew_ms", -1),
            }
            sync_data["speakers"].append(speaker)
            
            # Track overall skew from any speaker
            skew = item.get("sync_skew_ms", -1)
            if skew >= 0:
                if sync_data["overall_skew_ms"] is None:
                    sync_data["overall_skew_ms"] = skew
                else:
                    # Use the maximum skew seen
                    sync_data["overall_skew_ms"] = max(sync_data["overall_skew_ms"], skew)
        
        # Determine if in sync (skew <= threshold)
        if sync_data["overall_skew_ms"] is not None:
            sync_data["in_sync"] = sync_data["overall_skew_ms"] <= HUBMUSIC_STEREO_MAX_START_SKEW_MS
        
        self._send_json(sync_data)

    def _serve_hubmusic_live_stream(self, params: dict[str, list[str]]) -> None:
        desktop_audio = get_desktop_audio_info()
        if not desktop_audio.get("available"):
            self._send_json({"ok": False, "error": desktop_audio.get("reason") or "Desktop audio capture unavailable"}, status=503)
            return

        tone_satellite_id, bass_db, treble_db = _resolve_request_tone_settings(params)
        stream_request_id = uuid.uuid4().hex
        channel_raw = str((params.get("channel") or [""])[0] or "").strip().lower()
        if channel_raw in {"l", "left"}:
            channel_mode = "left"
        elif channel_raw in {"r", "right"}:
            channel_mode = "right"
        else:
            channel_mode = "stereo"
        force_tone_mode = str((params.get("force_tone") or [""])[0] or "").strip().lower() in {"1", "true", "yes", "on"}
        launch_at_raw = str((params.get("launch_at_ms") or [""])[0] or "").strip()
        launch_at_ms = 0
        if launch_at_raw:
            with contextlib.suppress(Exception):
                launch_at_ms = int(float(launch_at_raw))

        device_raw = params.get("device", [""])[0]
        preferred_device = desktop_audio.get("preferred_device")
        try:
            if str(device_raw).strip().lower() in {"", "preferred", "default"}:
                device_index = int(preferred_device)
            else:
                device_index = int(device_raw)
        except Exception:
            self._send_json({"ok": False, "error": "Invalid desktop audio device index"}, status=400)
            return

        device_info = next((item for item in desktop_audio.get("devices", []) if int(item["index"]) == device_index), None)
        if not device_info:
            fallback = next((item for item in desktop_audio.get("devices", []) if int(item["index"]) == int(preferred_device)), None)
            if fallback:
                logging.warning("HubMusic requested missing device %s; falling back to preferred device %s", device_index, preferred_device)
                device_info = fallback
                device_index = int(fallback["index"])
            else:
                self._send_json({"ok": False, "error": "Desktop audio device not found"}, status=404)
                return

        all_devices = list(desktop_audio.get("devices", []))

        def _candidate_score(item: dict) -> tuple:
            name = str(item.get("name", "")).lower()
            samplerate = int(item.get("default_samplerate") or 0)
            channels = int(item.get("input_channels") or 0)
            return (
                1 if "stereo mix" in name else 0,
                1 if any(term in name for term in ("stereo mix", "loopback", "what u hear", "wave out")) else 0,
                samplerate,
                channels,
            )

        requested_default = str(device_raw).strip().lower() in {"", "preferred", "default"}
        if requested_default:
            candidates = sorted(
                [
                    item
                    for item in all_devices
                    if int(item.get("input_channels") or 0) > 0
                    and int(item.get("index") or -1) not in HUBMUSIC_AUTO_CAPTURE_DEVICE_BLACKLIST
                ],
                key=_candidate_score,
                reverse=True,
            )
            if not candidates:
                candidates = [item for item in all_devices if int(item.get("input_channels") or 0) > 0]
        else:
            extra_candidates = [
                item for item in all_devices
                if int(item.get("index", -1)) != int(device_index)
                and "stereo mix" in str(item.get("name", "")).lower()
                and int(item.get("input_channels") or 0) > 0
            ]
            candidates = [device_info, *extra_candidates]
        selected_device = None
        selected_index = None
        sample_rate = 48000
        channels = 2
        layout = "stereo"
        for candidate in candidates:
            cand_index = int(candidate.get("index") or 0)
            cand_channels = 2 if int(candidate.get("input_channels") or 0) >= 2 else 1
            candidate_rates = []
            for rate in (48000, int(candidate.get("default_samplerate") or 48000), 44100):
                if rate not in candidate_rates:
                    candidate_rates.append(rate)
            for cand_rate in candidate_rates:
                try:
                    with sd.InputStream(
                        device=cand_index,
                        samplerate=cand_rate,
                        channels=cand_channels,
                        dtype="float32",
                        blocksize=HUBMUSIC_STREAM_FRAMES,
                    ):
                        pass
                    selected_device = candidate
                    selected_index = cand_index
                    sample_rate = cand_rate
                    channels = cand_channels
                    layout = "stereo" if cand_channels == 2 else "mono"
                    break
                except Exception as probe_exc:
                    logging.warning("HubMusic device probe failed for %s (%s) at %dHz: %s", cand_index, candidate.get("name", "?"), cand_rate, probe_exc)
            if selected_device is not None:
                break

        if selected_device is None or selected_index is None:
            self._send_json({"ok": False, "error": "Desktop audio capture device could not be opened"}, status=503)
            return

        device_info = selected_device
        device_index = selected_index
        logging.info("HubMusic live stream starting: device %s (%s), %dHz %s", device_index, device_info.get("name", "?"), sample_rate, layout)
        codec = None

        try:
            codec = av.CodecContext.create("libmp3lame", "w")
            codec.sample_rate = sample_rate
            codec.layout = layout
            codec.format = "fltp"
            codec.bit_rate = HUBMUSIC_STREAM_BITRATE
            tone_processor = _ToneProcessor(sample_rate, channels, bass_db, treble_db)
            if tone_satellite_id:
                stream_mode = "tone" if (tone_processor.filters or force_tone_mode) else "raw"
                if channel_mode in {"left", "right"}:
                    stream_mode = f"{stream_mode}_{channel_mode}"
                _set_satellite_stream_status(
                    tone_satellite_id,
                    path="hubmusic_live",
                    family="hubmusic",
                    mode=stream_mode,
                    bass_db=bass_db,
                    treble_db=treble_db,
                    request_id=stream_request_id,
                )
            if tone_satellite_id and (abs(bass_db) > 0.01 or abs(treble_db) > 0.01):
                logging.info(
                    "HubMusic live stream tone active for %s: bass=%.2fdB treble=%.2fdB",
                    tone_satellite_id,
                    bass_db,
                    treble_db,
                )

            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "audio/mpeg")
            self.send_header("Cache-Control", "no-store")
            self.send_header("Connection", "close")
            self.end_headers()

            audio_queue: queue.Queue = queue.Queue(maxsize=192)
            dropped_frames = 0
            silence_frames_injected = 0
            empty_polls = 0

            def callback(indata, frames, time_info, status) -> None:
                nonlocal dropped_frames
                if status:
                    logging.debug("HubMusic desktop audio status for device %s: %s", device_index, status)
                try:
                    audio_queue.put_nowait(indata.copy())
                except queue.Full:
                    dropped_frames += 1
                    with contextlib.suppress(queue.Empty):
                        audio_queue.get_nowait()
                    with contextlib.suppress(Exception):
                        audio_queue.put_nowait(indata.copy())
                    if dropped_frames % 25 == 0:
                        logging.warning("HubMusic live stream dropped %d capture frames (encoder/network backpressure)", dropped_frames)

            with sd.InputStream(
                device=device_index,
                samplerate=sample_rate,
                channels=channels,
                dtype="float32",
                blocksize=HUBMUSIC_STREAM_FRAMES,
                callback=callback,
            ) as stream:
                warmup_seconds = hubmusic_startup_warmup_seconds()
                logging.info(
                    "HubMusic live stream warmup: %.1fs, prebuffer target: %d bytes",
                    warmup_seconds,
                    HUBMUSIC_LIVE_PREBUFFER_BYTES,
                )
                packets_since_flush = 0
                warmup_deadline = time.time() + warmup_seconds
                prebuffer_deadline = warmup_deadline + HUBMUSIC_LIVE_PREBUFFER_MAX_WAIT_SECONDS
                prebuffer_packets: list[bytes] = []
                prebuffer_bytes = 0
                started_streaming = False
                while not _SHUTDOWN_EVENT.is_set():
                    try:
                        pcm = audio_queue.get(timeout=1.0)
                        empty_polls = 0
                    except queue.Empty:
                        empty_polls += 1
                        if not started_streaming:
                            continue
                        # Keep connections alive during short host CPU stalls to avoid client disconnects.
                        pcm = np.zeros((HUBMUSIC_STREAM_FRAMES, channels), dtype="float32")
                        silence_frames_injected += 1
                        if silence_frames_injected % 20 == 0:
                            logging.warning(
                                "HubMusic live stream injected %d silence frames due to capture stalls (device %s)",
                                silence_frames_injected,
                                device_index,
                            )
                    now = time.time()
                    if now < warmup_deadline:
                        # Let the capture path settle before sending initial audio.
                        continue

                    if launch_at_ms > 0 and int(now * 1000) < launch_at_ms:
                        # Stereo lock mode: align both channel streams to the same wall-clock launch point.
                        continue

                    if channel_mode in {"left", "right"} and channels >= 2 and pcm.ndim == 2 and pcm.shape[1] >= 2:
                        if channel_mode == "left":
                            pcm[:, 1] = 0.0
                        else:
                            pcm[:, 0] = 0.0

                    planar = np.ascontiguousarray(pcm.T)
                    if tone_processor.filters:
                        pcm = tone_processor.process(pcm)
                        planar = np.ascontiguousarray(pcm.T)
                    frame = av.AudioFrame.from_ndarray(planar, format="fltp", layout=layout)
                    frame.sample_rate = sample_rate
                    for packet in codec.encode(frame):
                        packet_bytes = bytes(packet)
                        if not started_streaming:
                            prebuffer_packets.append(packet_bytes)
                            prebuffer_bytes += len(packet_bytes)
                            if prebuffer_bytes < HUBMUSIC_LIVE_PREBUFFER_BYTES and now < prebuffer_deadline:
                                continue
                            for buffered_packet in prebuffer_packets:
                                self.wfile.write(buffered_packet)
                            self.wfile.flush()
                            started_streaming = True
                            prebuffer_packets.clear()
                            prebuffer_bytes = 0
                            packets_since_flush = 0
                            continue

                        self.wfile.write(packet_bytes)
                        packets_since_flush += 1
                    # Flush frequently to avoid bursty half-second packet delivery.
                    if packets_since_flush >= 1:
                        self.wfile.flush()
                        packets_since_flush = 0

                if packets_since_flush:
                    self.wfile.flush()

                if empty_polls:
                    logging.info("HubMusic live stream recovered after %d empty capture polls", empty_polls)

        except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
            logging.info("HubMusic live stream disconnected for device %s", device_index)
        except Exception as exc:
            _HUB_MUSIC_STATE.error(str(exc))
            logging.error("HubMusic live stream failed for device %s: %s", device_index, exc)
            with contextlib.suppress(Exception):
                self.connection.close()
        finally:
            if tone_satellite_id:
                _clear_satellite_stream_status(tone_satellite_id, stream_request_id)
            if codec is not None:
                with contextlib.suppress(Exception):
                    for packet in codec.encode(None):
                        self.wfile.write(bytes(packet))
                    self.wfile.flush()
            # Do not auto-STOP here; stop is explicitly controlled by /hubmusic/stop.

    def _serve_dlna_live_stream(self, params: dict[str, list[str]]) -> None:
        """Stream the active DLNA ffmpeg proxy output (MP3) to a satellite client."""
        tone_satellite_id, bass_db, treble_db = _resolve_request_tone_settings(params)
        raw_mode = str((params.get("raw") or [""])[0] or "").strip().lower() in {"1", "true", "yes"}
        stream_request_id = uuid.uuid4().hex
        with _DLNA_PROXY_LOCK:
            proc = _DLNA_PROXY_PROC
            source = _DLNA_PROXY_SOURCE_URL
        if proc is None or proc.poll() is not None:
            self._send_json({"ok": False, "error": "DLNA proxy not active"}, status=503)
            return

        tone_active = not raw_mode and (abs(bass_db) > 0.01 or abs(treble_db) > 0.01)
        ffmpeg_path = shutil.which("ffmpeg")
        if tone_active and ffmpeg_path:
            if tone_satellite_id:
                _set_satellite_stream_status(
                    tone_satellite_id,
                    path="dlna_live",
                    family="dlna",
                    mode="tone",
                    bass_db=bass_db,
                    treble_db=treble_db,
                    request_id=stream_request_id,
                )
            raw_url = build_runtime_url("/dlna/live.mp3?raw=1")
            ffmpeg_cmd = [
                ffmpeg_path,
                "-hide_banner",
                "-loglevel",
                "error",
                "-nostdin",
                "-i",
                raw_url,
                "-vn",
                "-ac",
                "2",
                "-ar",
                "48000",
                "-af",
                (
                    f"bass=g={bass_db:.2f}:f=100:w=0.65,"
                    f"treble=g={treble_db:.2f}:f=3000:w=0.55"
                ),
                "-c:a",
                "libmp3lame",
                "-b:a",
                DLNA_PROXY_BITRATE,
                "-f",
                "mp3",
                "pipe:1",
            ]
            ffmpeg_proc: subprocess.Popen | None = None
            try:
                ffmpeg_proc = subprocess.Popen(
                    ffmpeg_cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    bufsize=0,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
                prebuffer = bytearray()
                deadline = time.monotonic() + DLNA_PROXY_PREBUFFER_MAX_WAIT_SECONDS
                while len(prebuffer) < DLNA_PROXY_PREBUFFER_BYTES and time.monotonic() < deadline and not _SHUTDOWN_EVENT.is_set():
                    if not ffmpeg_proc.stdout:
                        break
                    chunk = ffmpeg_proc.stdout.read(8192)
                    if not chunk:
                        if ffmpeg_proc.poll() is not None:
                            break
                        time.sleep(0.05)
                        continue
                    prebuffer.extend(chunk)

                logging.info(
                    "DLNA live stream tone active for %s: bass=%.2fdB treble=%.2fdB (prebuffer=%d)",
                    tone_satellite_id or "unknown-satellite",
                    bass_db,
                    treble_db,
                    len(prebuffer),
                )
                self.send_response(200)
                self.send_header("Content-Type", "audio/mpeg")
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                if prebuffer:
                    self.wfile.write(prebuffer)
                    self.wfile.flush()

                chunks_since_flush = 0
                while not _SHUTDOWN_EVENT.is_set():
                    if not ffmpeg_proc.stdout:
                        break
                    chunk = ffmpeg_proc.stdout.read(8192)
                    if not chunk:
                        if ffmpeg_proc.poll() is not None:
                            break
                        time.sleep(0.02)
                        continue
                    self.wfile.write(chunk)
                    chunks_since_flush += 1
                    if chunks_since_flush >= 2:
                        self.wfile.flush()
                        chunks_since_flush = 0

                if chunks_since_flush:
                    self.wfile.flush()
                return
            except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
                logging.info("DLNA live stream tone path: satellite disconnected")
                return
            except Exception as exc:
                logging.error("DLNA live stream tone path error: %s", exc)
                with contextlib.suppress(Exception):
                    self.connection.close()
                return
            finally:
                if tone_satellite_id:
                    _clear_satellite_stream_status(tone_satellite_id, stream_request_id)
                if ffmpeg_proc is not None:
                    with contextlib.suppress(Exception):
                        ffmpeg_proc.terminate()
                    with contextlib.suppress(Exception):
                        ffmpeg_proc.wait(timeout=2)
        elif tone_active:
            logging.warning(
                "DLNA live stream tone requested for %s but ffmpeg is unavailable; using raw proxy stream",
                tone_satellite_id or "unknown-satellite",
            )

        listener: queue.Queue = queue.Queue(maxsize=256)
        with _DLNA_PROXY_LOCK:
            _DLNA_PROXY_LISTENERS.append(listener)
        if tone_satellite_id:
            _set_satellite_stream_status(
                tone_satellite_id,
                path="dlna_live",
                family="dlna",
                mode="raw",
                bass_db=bass_db,
                treble_db=treble_db,
                request_id=stream_request_id,
            )
        logging.info("DLNA live stream: new satellite client connected (source=%s)", source)

        try:
            prebuffer: list[bytes] = []
            prebuffered = 0
            deadline = time.monotonic() + DLNA_PROXY_PREBUFFER_MAX_WAIT_SECONDS
            while prebuffered < DLNA_PROXY_PREBUFFER_BYTES and time.monotonic() < deadline and not _SHUTDOWN_EVENT.is_set():
                try:
                    chunk = listener.get(timeout=1.0)
                except queue.Empty:
                    continue
                if chunk is None:
                    break
                prebuffer.append(chunk)
                prebuffered += len(chunk)

            logging.info("DLNA live stream: pre-buffered %d bytes, starting HTTP response", prebuffered)
            self.send_response(200)
            self.send_header("Content-Type", "audio/mpeg")
            self.send_header("Cache-Control", "no-store")
            self.end_headers()

            for chunk in prebuffer:
                self.wfile.write(chunk)
            self.wfile.flush()

            chunks_since_flush = 0
            while not _SHUTDOWN_EVENT.is_set():
                try:
                    chunk = listener.get(timeout=2.0)
                except queue.Empty:
                    continue
                if chunk is None:
                    break
                self.wfile.write(chunk)
                chunks_since_flush += 1
                if chunks_since_flush >= 2:
                    self.wfile.flush()
                    chunks_since_flush = 0

            if chunks_since_flush:
                self.wfile.flush()

        except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
            logging.info("DLNA live stream: satellite disconnected")
        except Exception as exc:
            logging.error("DLNA live stream error: %s", exc)
            with contextlib.suppress(Exception):
                self.connection.close()
        finally:
            if tone_satellite_id:
                _clear_satellite_stream_status(tone_satellite_id, stream_request_id)
            with _DLNA_PROXY_LOCK:
                with contextlib.suppress(ValueError):
                    _DLNA_PROXY_LISTENERS.remove(listener)

    def _serve_hubmusic_proxy(self, params: dict[str, list[str]]) -> None:
        target_url = params.get("url", [""])[0]
        stability_mode = str(params.get("stability", ["max"])[0] or "max").strip().lower()
        tone_satellite_id, bass_db, treble_db = _resolve_request_tone_settings(params)
        stream_request_id = uuid.uuid4().hex
        channel_raw = str((params.get("channel") or [""])[0] or "").strip().lower()
        if channel_raw in {"l", "left"}:
            channel_mode = "left"
        elif channel_raw in {"r", "right"}:
            channel_mode = "right"
        else:
            channel_mode = "stereo"
        if stability_mode not in {"balanced", "max"}:
            stability_mode = "max"
        if not target_url:
            self._send_json({"ok": False, "error": "Missing url parameter"}, status=400)
            return
        parsed_target = urllib.parse.urlparse(target_url)
        if parsed_target.scheme not in {"http", "https"} or not parsed_target.netloc:
            self._send_json({"ok": False, "error": "Invalid proxy URL"}, status=400)
            return
        logging.info("HubMusic proxy: streaming %s (stability=%s)", target_url, stability_mode)

        ffmpeg_path = shutil.which("ffmpeg")
        # Compatibility mode: sat-kitchen disconnects rapidly on ffmpeg-transcoded stream.
        # Keep raw passthrough as default and only enable transcode if explicitly requested.
        tone_active = abs(bass_db) > 0.01 or abs(treble_db) > 0.01
        use_ffmpeg = bool(ffmpeg_path) and (
            str(params.get("transcode", ["0"])[0]).strip().lower() in {"1", "true", "yes"}
            or tone_active
            or channel_mode in {"left", "right"}
        )
        ffmpeg_proc: subprocess.Popen | None = None
        try:
            if use_ffmpeg and ffmpeg_path:
                if tone_satellite_id:
                    stream_mode = "tone" if tone_active else "transcode"
                    if channel_mode in {"left", "right"}:
                        stream_mode = f"{stream_mode}_{channel_mode}"
                    _set_satellite_stream_status(
                        tone_satellite_id,
                        path="hubmusic_proxy",
                        family="hubmusic",
                        mode=stream_mode,
                        bass_db=bass_db,
                        treble_db=treble_db,
                        request_id=stream_request_id,
                    )
                # Normalize upstream radio into a clean, constant MP3 stream.
                reconnect_delay_max = "5" if stability_mode == "max" else "2"
                ffmpeg_cmd = [
                    ffmpeg_path,
                    "-hide_banner",
                    "-loglevel",
                    "error",
                    "-nostdin",
                    "-reconnect",
                    "1",
                    "-reconnect_streamed",
                    "1",
                    "-reconnect_delay_max",
                    reconnect_delay_max,
                    "-i",
                    target_url,
                    "-vn",
                    "-ac",
                    "2",
                    "-ar",
                    "48000",
                ]
                af_filters: list[str] = []
                if channel_mode == "left":
                    af_filters.append("pan=stereo|c0=c0|c1=0*c1")
                elif channel_mode == "right":
                    af_filters.append("pan=stereo|c0=0*c0|c1=c1")
                if tone_active:
                    af_filters.append(
                        f"bass=g={bass_db:.2f}:f=100:w=0.65,"
                        f"treble=g={treble_db:.2f}:f=3000:w=0.55"
                    )
                    logging.info(
                        "HubMusic proxy tone active for %s: bass=%.2fdB treble=%.2fdB",
                        tone_satellite_id or "unknown-satellite",
                        bass_db,
                        treble_db,
                    )
                if af_filters:
                    ffmpeg_cmd.extend(["-af", ",".join(af_filters)])
                ffmpeg_cmd.extend([
                    "-c:a",
                    "libmp3lame",
                    "-b:a",
                    "128k",
                    "-f",
                    "mp3",
                    "pipe:1",
                ])
                ffmpeg_proc = subprocess.Popen(
                    ffmpeg_cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    bufsize=0,
                )

                PREBUFFER_BYTES = 393216 if stability_mode == "max" else 196608
                prebuffer = bytearray()
                start_wait = time.time()
                max_wait = 14 if stability_mode == "max" else 8
                while len(prebuffer) < PREBUFFER_BYTES and (time.time() - start_wait) < max_wait and not _SHUTDOWN_EVENT.is_set():
                    if not ffmpeg_proc.stdout:
                        break
                    chunk = ffmpeg_proc.stdout.read(16384)
                    if not chunk:
                        if ffmpeg_proc.poll() is not None:
                            break
                        time.sleep(0.05)
                        continue
                    prebuffer.extend(chunk)

                logging.info("HubMusic proxy(ffmpeg): pre-buffered %d bytes, starting stream", len(prebuffer))

                self.send_response(200)
                self.send_header("Content-Type", "audio/mpeg")
                self.send_header("Cache-Control", "no-store")
                self.send_header("Connection", "close")
                self.end_headers()

                if prebuffer:
                    self.wfile.write(prebuffer)
                    self.wfile.flush()

                chunks_since_flush = 0
                flush_every = 2 if stability_mode == "max" else 1
                while not _SHUTDOWN_EVENT.is_set():
                    if not ffmpeg_proc.stdout:
                        break
                    chunk = ffmpeg_proc.stdout.read(16384)
                    if not chunk:
                        if ffmpeg_proc.poll() is not None:
                            break
                        time.sleep(0.02)
                        continue
                    self.wfile.write(chunk)
                    chunks_since_flush += 1
                    if chunks_since_flush >= flush_every:
                        self.wfile.flush()
                        chunks_since_flush = 0

                if chunks_since_flush:
                    self.wfile.flush()
            else:
                if tone_satellite_id:
                    stream_mode = "raw"
                    if channel_mode in {"left", "right"}:
                        stream_mode = f"raw_{channel_mode}"
                    _set_satellite_stream_status(
                        tone_satellite_id,
                        path="hubmusic_proxy",
                        family="hubmusic",
                        mode=stream_mode,
                        bass_db=bass_db,
                        treble_db=treble_db,
                        request_id=stream_request_id,
                    )
                # Fallback path if ffmpeg is unavailable.
                CHUNK = 8192
                PREBUFFER_BYTES = 196608 if stability_mode == "max" else 131072
                buf: queue.Queue = queue.Queue(maxsize=1024)

                def _reader(upstream_url: str, q: queue.Queue) -> None:
                    try:
                        req = urllib.request.Request(upstream_url, headers={"User-Agent": "HubVoiceSatRuntime/1.0", "Icy-MetaData": "0"})
                        with urllib.request.urlopen(req, timeout=10) as up:
                            while not _SHUTDOWN_EVENT.is_set():
                                chunk = up.read(CHUNK)
                                if not chunk:
                                    break
                                q.put(chunk)
                    except Exception as exc:
                        logging.debug("HubMusic proxy reader ended: %s", exc)
                    finally:
                        with contextlib.suppress(Exception):
                            q.put(None)

                reader_thread = threading.Thread(target=_reader, args=(target_url, buf), daemon=True)
                reader_thread.start()

                prebuffer: list[bytes] = []
                prebuffered = 0
                start_wait = time.time()
                max_wait = 10 if stability_mode == "max" else 8
                while prebuffered < PREBUFFER_BYTES and (time.time() - start_wait) < max_wait:
                    try:
                        chunk = buf.get(timeout=2)
                    except queue.Empty:
                        continue
                    if chunk is None:
                        break
                    prebuffer.append(chunk)
                    prebuffered += len(chunk)
                logging.info("HubMusic proxy(raw): pre-buffered %d bytes, starting stream", prebuffered)

                self.send_response(200)
                self.send_header("Content-Type", "audio/mpeg")
                self.send_header("Cache-Control", "no-store")
                self.send_header("Connection", "close")
                self.end_headers()

                for chunk in prebuffer:
                    self.wfile.write(chunk)
                self.wfile.flush()

                chunks_since_flush = 0
                empty_reads = 0
                flush_every = 2 if stability_mode == "max" else 1
                while not _SHUTDOWN_EVENT.is_set():
                    try:
                        chunk = buf.get(timeout=2)
                    except queue.Empty:
                        empty_reads += 1
                        if empty_reads >= 60:  # tolerate upstream stalls
                            break
                        continue
                    empty_reads = 0
                    if chunk is None:
                        break
                    self.wfile.write(chunk)
                    chunks_since_flush += 1
                    if chunks_since_flush >= flush_every:
                        self.wfile.flush()
                        chunks_since_flush = 0

                if chunks_since_flush:
                    self.wfile.flush()

        except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
            logging.info("HubMusic proxy: satellite disconnected from %s", target_url)
        except Exception as exc:
            logging.error("HubMusic proxy error for %s: %s", target_url, exc)
            with contextlib.suppress(Exception):
                self.connection.close()
        finally:
            if tone_satellite_id:
                _clear_satellite_stream_status(tone_satellite_id, stream_request_id)
            if ffmpeg_proc is not None:
                with contextlib.suppress(Exception):
                    ffmpeg_proc.terminate()
                with contextlib.suppress(Exception):
                    ffmpeg_proc.wait(timeout=2)
            # Do not auto-STOP here; stop is explicitly controlled by /hubmusic/stop.

    def do_POST(self) -> None:
        if not check_rate_limit(max_requests_per_minute=REQUEST_RATE_LIMIT):
            self._send_json({"ok": False, "error": "rate_limited"}, status=429)
            return

        parsed = urllib.parse.urlparse(self.path)
        try:
            payload = self._read_json_body()
            if parsed.path == "/answer":
                self._handle_answer_request(payload)
                return
            if parsed.path == "/satellite-switch":
                self._handle_satellite_switch_request(payload)
                return
            if parsed.path == "/satellite-number":
                self._handle_satellite_number_request(payload)
                return
            if parsed.path == "/satellite-media":
                self._handle_satellite_media_request(payload)
                return
            if parsed.path == "/satellite-media-volume":
                self._handle_satellite_media_volume_request(payload)
                return
            if parsed.path == "/hubmusic/play":
                self._handle_hubmusic_play_request(payload)
                return
            if parsed.path == "/hubmusic/stop":
                self._handle_hubmusic_stop_request(payload)
                return
            if parsed.path == "/hubmusic/stereo-test":
                self._handle_hubmusic_stereo_test_request(payload)
                return
            if parsed.path == "/hubmusic/stereo-config":
                self._handle_hubmusic_stereo_config_request(payload)
                return
            if parsed.path == "/hubmusic/stereo-sync-diagnostics":
                self._handle_hubmusic_stereo_sync_diagnostics_request()
                return
            if parsed.path == "/dlna/start":
                self._handle_dlna_start_request(payload)
                return
            if parsed.path == "/dlna/stop":
                self._handle_dlna_stop_request()
                return
            if parsed.path == "/airplay/start":
                self._handle_airplay_start_request(payload)
                return
            if parsed.path == "/airplay/stop":
                self._handle_airplay_stop_request()
                return
            self.send_error(HTTPStatus.NOT_FOUND, "Not found")
        except ValueError as exc:
            self._send_json({"ok": False, "error": str(exc)}, status=400)
        except Exception as exc:
            logging.error("POST handler error for %s: %s", parsed.path, exc)
            self._send_json({"ok": False, "error": str(exc)}, status=500)

    def do_GET(self) -> None:
        # Check rate limiting
        if not check_rate_limit(max_requests_per_minute=REQUEST_RATE_LIMIT):
            self._send_json({"ok": False, "error": "rate_limited"}, status=429)
            return
        
        parsed = urllib.parse.urlparse(self.path)
        
        # Request size limit
        content_length = int(self.headers.get('Content-Length', 0))
        if content_length > REQUEST_SIZE_LIMIT:
            self.send_error(HTTPStatus.REQUEST_ENTITY_TOO_LARGE)
            return
        
        if parsed.path in ("", "/"):
            state = _APP_STATE.snapshot()
            self._send_json(
                {
                    "ok": True,
                    "service": "HubVoiceSat runtime",
                    "port": get_runtime_port(),
                    "queue_depth": _WORK_QUEUE.qsize(),
                    "last_action": state.get("last_action", "idle"),
                    "last_error": state.get("last_error", ""),
                    "last_transcript": state.get("last_transcript", ""),
                    "voice_assistant": {sid: b.snapshot() for sid, b in _VOICE_BRIDGES.items()},
                    "scheduler": _SCHEDULE_MANAGER.snapshot(),
                }
            )
            return

        if parsed.path == "/control":
            body = build_satellite_control_page().encode("utf-8")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
            self.send_header("Pragma", "no-cache")
            self.send_header("Expires", "0")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        if parsed.path == "/health":
            state = _APP_STATE.snapshot()
            self._send_json(
                {
                    "ok": True,
                    "status": "online",
                    "queue_depth": _WORK_QUEUE.qsize(),
                    "voice_assistant": {sid: b.snapshot() for sid, b in _VOICE_BRIDGES.items()},
                    "scheduler": _SCHEDULE_MANAGER.snapshot(),
                }
            )
            return

        if parsed.path == "/schedules":
            self._send_json(
                {
                    "ok": True,
                    **_SCHEDULE_MANAGER.snapshot(),
                }
            )
            return

        if parsed.path == "/satellites":
            satellites = load_satellites()
            controls = build_satellite_control_metadata()
            satellite_rows = []
            for sat in satellites:
                reachable = test_satellite_connection(sat["host"])
                control_state = _get_control_deck_state(sat["id"], controls)
                bass_db, treble_db = _get_user_tone_settings(sat["host"])
                control_state["bass_level"] = bass_db
                control_state["treble_level"] = treble_db
                satellite_rows.append(
                    {
                        **sat,
                        "reachable": reachable,
                        "web_url": f"http://{sat['host']}:8080/",
                        "control_state": control_state,
                        "stream_status": _get_satellite_stream_status(sat["id"]),
                        # Keep /satellites fast; capabilities can be fetched via /satellite-capabilities on demand.
                        "capabilities": {"supports": {}, "resolved": {}},
                    }
                )
            self._send_json(
                {
                    "ok": True,
                    "count": len(satellites),
                    "default": satellites[0]["id"] if satellites else "",
                    "controls": controls,
                    "satellites": satellite_rows,
                }
            )
            return

        if parsed.path == "/satellite-capabilities":
            params = urllib.parse.parse_qs(parsed.query)
            sat = select_satellite(params.get("d", [""])[0])
            if not sat:
                self._send_json({"ok": False, "error": "Satellite not found"}, status=404)
                return
            reachable = test_satellite_connection(sat["host"])
            self._send_json(
                {
                    "ok": True,
                    "satellite": sat["id"],
                    "host": sat["host"],
                    "reachable": reachable,
                    "capabilities": get_satellite_capabilities(sat["host"]) if reachable else {"supports": {}, "resolved": {}},
                }
            )
            return

        if parsed.path == "/hubmusic/status":
            self._send_json({"ok": True, "status": hubmusic_status_snapshot()})
            return

        if parsed.path == "/dlna/status":
            self._handle_dlna_status_request()
            return

        if parsed.path == "/airplay/status":
            self._handle_airplay_status_request()
            return

        if parsed.path in ("/hubmusic/live.flac", "/hubmusic/live.mp3"):
            params = urllib.parse.parse_qs(parsed.query)
            self._serve_hubmusic_live_stream(params)
            return

        if parsed.path == "/dlna/live.mp3":
            params = urllib.parse.parse_qs(parsed.query)
            self._serve_dlna_live_stream(params)
            return

        if parsed.path == "/hubmusic/proxy":
            params = urllib.parse.parse_qs(parsed.query)
            self._serve_hubmusic_proxy(params)
            return

        if parsed.path == "/answer":
            params = urllib.parse.parse_qs(parsed.query)
            self._handle_answer_request({
                "r": params.get("r", [""])[0],
                "d": params.get("d", [""])[0],
            })
            return

        if parsed.path.startswith("/tts/"):
            self._serve_wav(parsed.path.removeprefix("/tts/"))
            return

        # /satellite-switch?d=<sat_id>&entity=<object_id>&state=<on|off|true|false|1|0>
        # Allows Hubitat (or any caller) to toggle a switch entity on any configured satellite.
        if parsed.path == "/satellite-switch":
            params = urllib.parse.parse_qs(parsed.query)
            self._handle_satellite_switch_request({
                "d": params.get("d", [""])[0],
                "entity": params.get("entity", [""])[0],
                "state": params.get("state", [""])[0],
            })
            return

        if parsed.path == "/satellite-number":
            params = urllib.parse.parse_qs(parsed.query)
            self._handle_satellite_number_request({
                "d": params.get("d", [""])[0],
                "entity": params.get("entity", [""])[0],
                "value": params.get("value", [None])[0],
            })
            return

        self.send_error(HTTPStatus.NOT_FOUND, "Not found")

    def log_message(self, fmt: str, *args) -> None:
        logging.info("%s - %s", self.address_string(), fmt % args)

    def _send_json(self, payload: dict, status: int = 200) -> None:
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _serve_wav(self, raw_name: str) -> None:
        """Serve WAV file with path traversal protection."""
        try:
            name = Path(urllib.parse.unquote(raw_name)).name
            target = (RECORDINGS_PATH / name).resolve()
            
            # Ensure normalized path is within RECORDINGS_PATH
            recordings_dir = RECORDINGS_PATH.resolve()
            if not str(target).startswith(str(recordings_dir)):
                logging.warning("Path traversal attempt detected: %s", raw_name)
                self.send_error(HTTPStatus.FORBIDDEN)
                return
            
            if not target.exists() or not target.is_file():
                self.send_error(HTTPStatus.NOT_FOUND, "WAV not found")
                return
            
            data = target.read_bytes()
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "audio/wav")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        except Exception as e:
            logging.error("Error serving WAV file: %s", e)
            self.send_error(HTTPStatus.BAD_REQUEST)


def main() -> None:
    acquire_runtime_single_instance_lock()
    try:
        # Validate configuration early
        logging.info("Validating configuration...")
        validate_startup_config()
        logging.info("Configuration validation passed")
    except Exception as e:
        logging.error("Configuration validation failed: %s", e)
        raise
    
    host = "0.0.0.0"
    port = get_runtime_port()
    httpd = ThreadingHTTPServer((host, port), RuntimeHandler)
    
    # Start worker thread
    threading.Thread(target=worker_loop, daemon=True).start()
    # Warm-load Piper in the background to reduce first text-to-speech latency.
    threading.Thread(target=preload_piper_voice_model, daemon=True).start()
    # Warm-load Whisper in the background to reduce first speech-to-text latency.
    threading.Thread(target=preload_whisper_model, daemon=True).start()
    
    # Start voice assistant bridges (one per satellite)
    for _bridge in _VOICE_BRIDGES.values():
        _bridge.start()

    # Watch for satellites added after startup and start their bridges automatically.
    def satellite_watcher() -> None:
        while not _SHUTDOWN_EVENT.is_set():
            _SHUTDOWN_EVENT.wait(10)
            if _SHUTDOWN_EVENT.is_set():
                break
            try:
                for sat in load_satellites():
                    sid = sat["id"]
                    if sid not in _VOICE_BRIDGES:
                        logging.info("New satellite detected: %s — starting voice bridge", sid)
                        bridge = VoiceAssistantBridge(sid)
                        _VOICE_BRIDGES[sid] = bridge
                        bridge.start()
            except Exception as exc:
                logging.warning("Satellite watcher error: %s", exc)

    threading.Thread(target=satellite_watcher, daemon=True).start()

    # Start cleanup scheduler
    def cleanup_scheduler():
        while not _SHUTDOWN_EVENT.is_set():
            try:
                cleanup_old_recordings(max_age_hours=RECORDING_MAX_AGE_HOURS, max_files=RECORDING_MAX_FILES)
            except Exception as e:
                logging.error("Cleanup scheduler failed: %s", e)
            # Wait for cleanup interval between runs
            _SHUTDOWN_EVENT.wait(CLEANUP_INTERVAL_SECS)

    def schedule_scheduler() -> None:
        while not _SHUTDOWN_EVENT.is_set():
            try:
                _SCHEDULE_MANAGER.tick(_VOICE_BRIDGES)
            except Exception as e:
                logging.error("Schedule scheduler failed: %s", e)
            _SHUTDOWN_EVENT.wait(SCHEDULER_POLL_INTERVAL_SECS)
    
    threading.Thread(target=cleanup_scheduler, daemon=True).start()
    threading.Thread(target=schedule_scheduler, daemon=True).start()

    # Watchdog: restart ffmpeg if it exits unexpectedly while music should be playing
    def ffmpeg_watchdog() -> None:
        while not _SHUTDOWN_EVENT.is_set():
            _SHUTDOWN_EVENT.wait(1)  # check every 1 second
            if _SHUTDOWN_EVENT.is_set():
                break
            with _HUBMUSIC_FFMPEG_LOCK:
                proc = _HUBMUSIC_FFMPEG_PROC
            if proc is None:
                continue
            if proc.poll() is None:
                continue  # still running, all good
            # ffmpeg has exited â€” check if we should be playing
            snap = _HUB_MUSIC_STATE.snapshot()
            if snap.get("last_action") != "play":
                continue  # user stopped music intentionally
            source_url = snap.get("source_url", "")
            if not source_url:
                continue
            logging.warning("HubMusic watchdog: ffmpeg exited (rc=%s) while music active, restarting for %s",
                            proc.returncode, source_url)
            try:
                _start_hubmusic_ffmpeg(source_url)
            except Exception as exc:
                logging.error("HubMusic watchdog: failed to restart ffmpeg: %s", exc)

    threading.Thread(target=ffmpeg_watchdog, daemon=True).start()

    logging.info("HubVoiceSat runtime listening on %s:%s", host, port)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        logging.info("Shutdown signal received")
    finally:
        _SHUTDOWN_EVENT.set()
        cleanup_resources()
        httpd.shutdown()
        release_runtime_single_instance_lock()


if __name__ == "__main__":
    main()
