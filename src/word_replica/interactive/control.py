from __future__ import annotations

import threading

from word_replica.domain.enums import InteractiveRunState, InteractiveSpeedMode
from word_replica.domain.reconstruction import ControlDecision


class InteractiveRunControl:
    def __init__(self) -> None:
        self._condition = threading.Condition()
        self._state = InteractiveRunState.CREATED
        self._stop_requested = False
        self._speed_mode = InteractiveSpeedMode.FAST
        self._characters_per_second = 25.0

    @property
    def state(self) -> InteractiveRunState:
        with self._condition:
            return self._state

    @property
    def speed_mode(self) -> InteractiveSpeedMode:
        with self._condition:
            return self._speed_mode

    @property
    def characters_per_second(self) -> float:
        with self._condition:
            return self._characters_per_second

    def start(self) -> None:
        with self._condition:
            if self._state not in {InteractiveRunState.CREATED, InteractiveRunState.PAUSED}:
                return
            self._state = InteractiveRunState.RUNNING
            self._condition.notify_all()

    def pause(self) -> None:
        with self._condition:
            if self._state is InteractiveRunState.RUNNING:
                self._state = InteractiveRunState.PAUSED

    def resume(self) -> None:
        with self._condition:
            if self._state is InteractiveRunState.PAUSED:
                self._state = InteractiveRunState.RUNNING
                self._condition.notify_all()

    def stop(self) -> None:
        with self._condition:
            self._stop_requested = True
            self._state = InteractiveRunState.STOPPED
            self._condition.notify_all()

    def set_speed(
        self,
        mode: InteractiveSpeedMode,
        characters_per_second: float | None = None,
    ) -> None:
        with self._condition:
            if characters_per_second is not None:
                if characters_per_second <= 0:
                    raise ValueError("characters_per_second must be > 0")
                self._characters_per_second = float(characters_per_second)
            self._speed_mode = mode

    def before_next_event(self, last_completed_index: int) -> ControlDecision:
        with self._condition:
            while self._state is InteractiveRunState.PAUSED and not self._stop_requested:
                self._condition.wait()
            return ControlDecision(
                stop_requested=self._stop_requested or self._state is InteractiveRunState.STOPPED,
                last_completed_index=last_completed_index,
            )
