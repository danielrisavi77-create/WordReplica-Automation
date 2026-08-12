from __future__ import annotations

import time
from collections.abc import Callable

from word_replica.config import InteractiveOptions
from word_replica.domain.enums import InteractiveSpeedMode


_OBJECT_EVENTS = {
    "BeginTable", "EndTable", "SetTableProperties", "SetColumnWidth", "SetRowProperties",
    "SetCellProperties", "MergeCells", "EnterCell", "LeaveCell", "InsertImage",
    "SetImageSize", "SetImageWrap", "SetImagePosition", "SetImageCrop", "SetImageRotation",
    "BeginSection", "EndSection", "BeginHeader", "EndHeader", "BeginFooter", "EndFooter",
}


class SpeedController:
    def __init__(self, options: InteractiveOptions, *, sleep_fn: Callable[[float], None] = time.sleep) -> None:
        self._sleep = sleep_fn
        self._mode = options.speed_mode
        self._custom_rate = options.characters_per_second
        self._object_delay = options.object_step_delay_ms / 1000.0

    @property
    def mode(self) -> InteractiveSpeedMode:
        return self._mode

    def set_mode(self, mode: InteractiveSpeedMode) -> None:
        self._mode = mode

    def set_custom_rate(self, characters_per_second: float) -> None:
        if characters_per_second <= 0:
            raise ValueError("characters_per_second must be > 0")
        self._custom_rate = float(characters_per_second)

    def _character_delay(self) -> float:
        if self._mode is InteractiveSpeedMode.MAXIMUM:
            return 0.0
        if self._mode is InteractiveSpeedMode.SLOW:
            return 1.0 / 5.0
        if self._mode is InteractiveSpeedMode.FAST:
            return 1.0 / 25.0
        return 1.0 / self._custom_rate

    def delay_after(self, event_type: str) -> None:
        if event_type == "InsertCharacter":
            delay = self._character_delay()
        elif event_type in _OBJECT_EVENTS:
            delay = self._object_delay
        else:
            delay = 0.0
        if delay > 0:
            self._sleep(delay)
