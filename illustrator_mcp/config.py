"""
Configuration management for Illustrator MCP.
"""
import logging as _logging
from pathlib import Path
from typing import Tuple, Union

from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict


# Named constants for timeouts (avoids magic numbers)
BRIDGE_STARTUP_TIMEOUT: float = 10.0  # seconds to wait for bridge startup
BRIDGE_EXECUTION_BUFFER: float = 5.0  # extra timeout for thread coordination
RECONNECT_INTERVAL_MS: int = 3000     # CEP panel reconnect interval


# MCP hosts do not all launch the server with the project root as CWD —
# Google Antigravity, for example, starts stdio servers from its own program
# folder. Resolve .env next to the package as well so WS_PORT / TIMEOUT set by
# the user are honoured no matter where the process was spawned from.
PROJECT_ROOT: Path = Path(__file__).resolve().parent.parent
ENV_FILES: Tuple[Union[str, Path], ...] = (PROJECT_ROOT / ".env", ".env")


class Config(BaseSettings):
    """Configuration with validation and .env support."""
    
    model_config = SettingsConfigDict(
        # Later entries win: a .env in the CWD overrides the project-root one.
        env_file=ENV_FILES,
        env_file_encoding='utf-8',
        case_sensitive=False,
        extra='ignore'
    )
    
    # WebSocket settings
    ws_host: str = Field(default="localhost", description="WebSocket host")
    ws_port: int = Field(default=8081, ge=1024, le=65535, description="WebSocket port for bridge")
    
    # Timeout settings
    timeout: float = Field(default=30.0, ge=1.0, le=300.0, description="Operation timeout in seconds")
    
    # Watchdog settings (E1)
    watchdog_interval: float = Field(default=10.0, ge=1.0, le=60.0, description="Seconds between panel health checks")
    watchdog_stale_threshold: float = Field(default=30.0, ge=10.0, le=120.0, description="Seconds of silence before panel is declared dead")
    
    # Log rotation settings (E2)
    max_log_sessions: int = Field(default=50, ge=1, le=1000, description="Maximum session log files to keep")
    max_log_age_days: int = Field(default=30, ge=1, le=365, description="Maximum age in days for session logs")
    
    # Logging
    log_level: str = Field(default="INFO", description="Log level: DEBUG, INFO, WARNING, ERROR")
    
    # I6: Configurable message size limit (WebSocket bridge)
    max_message_size_mb: int = Field(default=10, ge=1, le=100, description="Max WebSocket message size in MB")
    
    @property
    def ws_url(self) -> str:
        """WebSocket URL for CEP panel connection."""
        return f"ws://{self.ws_host}:{self.ws_port}"


# Global config instance
config = Config()

# Set log level once at startup (not dynamically updated if config changes later)
_logging.getLogger("illustrator_mcp").setLevel(
    getattr(_logging, config.log_level.upper(), _logging.INFO)
)
