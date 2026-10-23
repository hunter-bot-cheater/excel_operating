from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class ReportGenerationRequest:
    bom_path: Path
    report_paths: list[Path]
    output_dir: Path
    round_no: int | None = None
