from dataclasses import dataclass, asdict
from datetime import datetime

@dataclass
class Candidate:
    name: str
    external_id: str = ""

@dataclass
class BackupResult:
    candidate: str
    category: str
    status: str
    files: int = 0
    bytes: int = 0
    detail: str = ""
    timestamp: str = ""

    def row(self):
        if not self.timestamp:
            self.timestamp = datetime.now().isoformat(timespec="seconds")
        return asdict(self)
