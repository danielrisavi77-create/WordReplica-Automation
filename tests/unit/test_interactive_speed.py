from word_replica.config import InteractiveOptions
from word_replica.domain.enums import InteractiveSpeedMode
from word_replica.interactive.speed import SpeedController


def test_custom_20_cps_sleeps_point_zero_five_seconds_for_character():
    calls = []
    speed = SpeedController(InteractiveOptions(speed_mode=InteractiveSpeedMode.CUSTOM, characters_per_second=20), sleep_fn=calls.append)
    speed.delay_after("InsertCharacter")
    assert calls == [0.05]


def test_maximum_has_no_intentional_character_delay():
    calls = []
    speed = SpeedController(InteractiveOptions(speed_mode=InteractiveSpeedMode.MAXIMUM), sleep_fn=calls.append)
    speed.delay_after("InsertCharacter")
    assert calls == []


def test_slow_and_fast_are_deterministic_and_nonzero():
    slow_calls, fast_calls = [], []
    SpeedController(InteractiveOptions(speed_mode=InteractiveSpeedMode.SLOW), sleep_fn=slow_calls.append).delay_after("InsertCharacter")
    SpeedController(InteractiveOptions(speed_mode=InteractiveSpeedMode.FAST), sleep_fn=fast_calls.append).delay_after("InsertCharacter")
    assert slow_calls == [0.2]
    assert fast_calls == [0.04]


def test_object_event_uses_configured_object_delay():
    calls = []
    speed = SpeedController(InteractiveOptions(object_step_delay_ms=150), sleep_fn=calls.append)
    speed.delay_after("InsertImage")
    assert calls == [0.15]


def test_speed_change_affects_next_delay_only():
    calls = []
    speed = SpeedController(InteractiveOptions(speed_mode=InteractiveSpeedMode.FAST), sleep_fn=calls.append)
    speed.delay_after("InsertCharacter")
    speed.set_custom_rate(10)
    speed.set_mode(InteractiveSpeedMode.CUSTOM)
    speed.delay_after("InsertCharacter")
    assert calls == [0.04, 0.1]
