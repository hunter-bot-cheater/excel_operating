from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class CloudTemplateRequest:
    input_path: Path
    output_dir: Path
    default_year: str | None = None
