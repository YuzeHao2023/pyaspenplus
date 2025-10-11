from dataclasses import dataclass, asdict
from typing import Dict, Any

@dataclass
class Stream:
    """Simple representation of a material/energy stream."""
    name: str
    flow: float
    temperature: float = None
    pressure: float = None
    composition: Dict[str, float] | None = None

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)
