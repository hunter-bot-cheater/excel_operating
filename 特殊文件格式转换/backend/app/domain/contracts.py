from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class SpecialFormatConversionRequest:
    input_path: Path
    output_dir: Path
    force_openpyxl: bool = False
