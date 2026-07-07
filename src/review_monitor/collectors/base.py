from __future__ import annotations

from abc import ABC, abstractmethod
from datetime import datetime

from ..models import Review


class Collector(ABC):
    @abstractmethod
    def collect(self, date_from: datetime, date_to: datetime) -> list[Review]:
        raise NotImplementedError
