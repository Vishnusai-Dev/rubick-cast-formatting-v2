"""Application settings (constants)."""

from dataclasses import dataclass


@dataclass(frozen=True)
class AppSettings:
    app_name: str = "Rubick CAST Formatting"
    app_version: str = "0.2.0"
    max_file_size_mb: int = 50
    default_exclude_sheets: str = "masterdata"
    input_generator_output: str = "Input_Generated.xlsx"
    export_output: str = "Formatted_Output.xlsx"


def get_settings() -> AppSettings:
    return AppSettings()
