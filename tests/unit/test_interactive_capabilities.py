from word_replica.domain.enums import CapabilityClass
from word_replica.domain.model import DrawingRef
from word_replica.interactive.capabilities import classify_drawing


def drawing(**overrides):
    values = dict(
        element_id="d1", asset_id="a1", source_path="word/media/image1.png",
        representation="inline", width_emu=1000, height_emu=800,
        lock_aspect_ratio=True, wrap_type=None, horizontal_relative_from=None,
        horizontal_position_emu=None, vertical_relative_from=None, vertical_position_emu=None,
        distance_top_emu=0, distance_bottom_emu=0, distance_left_emu=0, distance_right_emu=0,
        crop={}, rotation_degrees=0.0, behind_text=False, z_order=None,
    )
    values.update(overrides)
    return DrawingRef(**values)


def test_png_and_jpeg_word_drawings_are_reconstructable():
    assert classify_drawing(drawing()).classification is CapabilityClass.RECONSTRUCTED
    assert classify_drawing(drawing(source_path="word/media/photo.jpeg")).classification is CapabilityClass.RECONSTRUCTED


def test_preserved_representation_has_reason():
    decision = classify_drawing(drawing(representation="preserved"))
    assert decision.classification is CapabilityClass.PRESERVED
    assert decision.reason


def test_unsupported_representation_has_explicit_reason():
    decision = classify_drawing(drawing(representation="vml-active"))
    assert decision.classification is CapabilityClass.UNSUPPORTED
    assert "unsupported" in decision.reason.lower()


def test_image_zorder_event_closes_restart_transaction_and_applies_relative_stack_order():
    from types import SimpleNamespace
    from word_replica.domain.reconstruction import ReconstructionEvent
    from word_replica.renderers.interactive_word import InteractiveWordController

    class Shape:
        def __init__(self): self.calls=[]
        def ZOrder(self, command): self.calls.append(command)
    shape=Shape()
    controller=InteractiveWordController.for_testing(active_range=SimpleNamespace())
    controller._active_image=shape
    controller._active_image_representation="floating"
    controller._image_transaction_active=True
    controller.execute_event(ReconstructionEvent("SetImageZOrder","d1",{"z_order":7}))
    assert controller.is_restart_safe() is True
    assert shape.calls == [0]


def test_png_magic_bytes_allow_reconstruction_even_when_source_extension_is_undefined():
    from word_replica.domain.model import BinaryAsset

    asset = BinaryAsset(
        asset_id="a1",
        part_name="word/media/image1.undefined",
        content_type="application/octet-stream",
        sha256="sha",
        bytes_data=b"\x89PNG\r\n\x1a\n" + b"payload",
    )
    decision = classify_drawing(drawing(source_path="word/media/image1.undefined"), asset=asset)
    assert decision.classification is CapabilityClass.RECONSTRUCTED

def test_materialized_undefined_png_asset_uses_png_temp_extension(tmp_path):
    from types import SimpleNamespace
    from word_replica.services.interactive_rebuild import InteractiveRebuildService
    png = b"\x89PNG\r\n\x1a\n" + b"x" * 16
    asset = SimpleNamespace(part_name="word/media/image1.undefined", content_type="application/octet-stream", bytes_data=png)
    model = SimpleNamespace(assets={"a1": asset})
    result = InteractiveRebuildService._materialize_assets(model, tmp_path)
    assert result["a1"].suffix == ".png"
    assert result["a1"].read_bytes() == png
