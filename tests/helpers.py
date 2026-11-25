"""Ajutoare comune pentru testele cu fire."""

from __future__ import annotations

import time
from collections.abc import Callable


def wait_until(condition: Callable[[], object], timeout: float = 5.0) -> bool:
    """Asteapta, cu pauze scurte, pana cand conditia e adevarata (nu un timp fix: pe un runner
    CI lent, firul poate intarzia). False daca limita de timp a trecut."""
    deadline = time.monotonic() + timeout
    while not condition():
        if time.monotonic() > deadline:
            return False
        time.sleep(0.01)
    return True
