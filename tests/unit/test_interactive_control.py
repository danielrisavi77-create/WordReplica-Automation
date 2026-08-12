from threading import Event, Thread

from word_replica.domain.enums import InteractiveRunState, InteractiveSpeedMode
from word_replica.interactive.control import InteractiveRunControl


def test_pause_blocks_only_before_next_atomic_event_and_resume_releases():
    control = InteractiveRunControl()
    control.start()
    control.pause()
    entered = Event(); released = Event()

    def worker():
        entered.set()
        decision = control.before_next_event(3)
        assert decision.stop_requested is False
        released.set()

    thread = Thread(target=worker)
    thread.start()
    assert entered.wait(1)
    assert not released.wait(0.05)
    assert control.state is InteractiveRunState.PAUSED
    control.resume()
    assert released.wait(1)
    thread.join(1)
    assert control.state is InteractiveRunState.RUNNING


def test_stop_returns_decision_preserving_last_completed_index():
    control = InteractiveRunControl()
    control.start()
    control.stop()
    decision = control.before_next_event(11)
    assert decision.stop_requested is True
    assert decision.last_completed_index == 11
    assert control.state is InteractiveRunState.STOPPED


def test_speed_state_is_plain_domain_data_not_tkinter():
    control = InteractiveRunControl()
    control.set_speed(InteractiveSpeedMode.CUSTOM, 12.5)
    assert control.speed_mode is InteractiveSpeedMode.CUSTOM
    assert control.characters_per_second == 12.5
    assert "tkinter" not in repr(vars(control)).lower()
