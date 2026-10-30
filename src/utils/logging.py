"""
Structured logging configuration for Rubick CAST Formatting.

Provides JSON-structured logging with configurable log levels and handlers.
"""

import json
import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional


class JSONFormatter(logging.Formatter):
    """
    JSON-structured log formatter for machine-parseable logs.

    Outputs log records as JSON objects with consistent fields:
    - timestamp: ISO 8601 formatted time
    - level: Log level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
    - logger: Name of the logger
    - message: Log message
    - module: Module name
    - function: Function name
    - line: Line number
    - extra: Any additional context passed via 'extra' kwarg
    """

    def format(self, record: logging.LogRecord) -> str:
        """Format log record as JSON string."""
        log_data = {
            "timestamp": datetime.utcnow().isoformat() + "Z",
            "level": record.levelname,
            "logger": record.name,
            "message": record.getMessage(),
            "module": record.module,
            "function": record.funcName,
            "line": record.lineno,
        }

        # Include extra fields if present
        if hasattr(record, "extra_data"):
            log_data["extra"] = record.extra_data

        # Include exception info if present
        if record.exc_info:
            log_data["exception"] = self.formatException(record.exc_info)

        return json.dumps(log_data)


class ColoredConsoleFormatter(logging.Formatter):
    """
    Console formatter with color support for terminal output.

    Provides human-readable logs with ANSI color coding for different
    log levels to improve readability in development.
    """

    # ANSI color codes
    COLORS = {
        "DEBUG": "\033[36m",  # Cyan
        "INFO": "\033[32m",   # Green
        "WARNING": "\033[33m", # Yellow
        "ERROR": "\033[31m",  # Red
        "CRITICAL": "\033[35m", # Magenta
    }
    RESET = "\033[0m"

    def format(self, record: logging.LogRecord) -> str:
        """Format log record with color coding."""
        color = self.COLORS.get(record.levelname, self.RESET)
        timestamp = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
        level = f"{color}{record.levelname:<8}{self.RESET}"
        logger = f"{record.name}"
        message = record.getMessage()

        if record.exc_info:
            message = f"{message}\n{self.formatException(record.exc_info)}"

        return f"[{timestamp}] [{level}] [{logger}] {message}"


def setup_logging(
    level: str = "INFO",
    log_file: Optional[Path] = None,
    json_format: bool = False
) -> logging.Logger:
    """
    Configure structured logging for the application.

    Args:
        level: Log level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        log_file: Optional path to write logs to file
        json_format: If True, use JSON formatting; otherwise human-readable

    Returns:
        Configured root logger instance
    """
    # Clear any existing handlers
    root_logger = logging.getLogger()
    root_logger.handlers.clear()

    # Set log level
    numeric_level = getattr(logging, level.upper(), logging.INFO)
    root_logger.setLevel(numeric_level)

    # Determine formatter
    if json_format:
        formatter = JSONFormatter()
    else:
        formatter = ColoredConsoleFormatter()

    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    console_handler.setLevel(numeric_level)
    root_logger.addHandler(console_handler)

    # File handler (optional)
    if log_file:
        log_file.parent.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(log_file)
        file_handler.setFormatter(JSONFormatter())  # Always use JSON for files
        file_handler.setLevel(numeric_level)
        root_logger.addHandler(file_handler)

    return root_logger


def get_logger(name: str) -> logging.Logger:
    """
    Get a logger with the given name.

    Args:
        name: Name for the logger (typically __name__)

    Returns:
        Logger instance with configured handlers
    """
    return logging.getLogger(name)
