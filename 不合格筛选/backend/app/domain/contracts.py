from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class ScreeningRequest:
    input_paths: list[Path]
    output_dir: Path
