from dataclasses import dataclass
from datetime import date

@dataclass(frozen=True)
class DateRange:
    start: date
    end: date

    def __post_init__(self):
        if self.end < self.start:
            raise ValueError("A data final não pode ser anterior à inicial.")

@dataclass
class DownloadResult:
    reference: str
    document_type: str
    status: str
    detail: str = ""
