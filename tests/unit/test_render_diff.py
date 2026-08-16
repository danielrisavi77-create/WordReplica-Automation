from PIL import Image, ImageDraw
from word_replica.qa.render import compare_page_images


def test_identical_pages_have_zero_difference(tmp_path):
    a = tmp_path / 'a.png'; b = tmp_path / 'b.png'
    Image.new('RGB', (100, 100), 'white').save(a)
    Image.new('RGB', (100, 100), 'white').save(b)
    metric = compare_page_images(a, b)
    assert metric.changed_pixel_ratio == 0.0
    assert metric.mean_absolute_error == 0.0


def test_page_metric_reports_one_pixel_blurred_mae(tmp_path):
    a = tmp_path / "a.png"
    b = tmp_path / "b.png"
    source = Image.new("RGB", (100, 100), "white")
    rebuilt = Image.new("RGB", (100, 100), "white")
    ImageDraw.Draw(source).line((20, 20, 80, 20), fill="black", width=1)
    ImageDraw.Draw(rebuilt).line((20, 21, 80, 21), fill="black", width=1)
    source.save(a)
    rebuilt.save(b)

    metric = compare_page_images(a, b)

    assert metric.blurred_mean_absolute_error is not None
    assert metric.blurred_mean_absolute_error < metric.mean_absolute_error


def test_compare_pdfs_fails_tolerance_when_page_counts_differ(monkeypatch, tmp_path):
    from word_replica.qa import render
    a=tmp_path/"a.png"; b=tmp_path/"b.png"
    Image.new('RGB',(20,20),'white').save(a); Image.new('RGB',(20,20),'white').save(b)
    calls=[]
    def fake_rasterize(pdf,out_dir,dpi=144):
        calls.append(pdf)
        return [a,b] if len(calls)==1 else [a]
    monkeypatch.setattr(render,"rasterize_pdf",fake_rasterize)
    result=render.compare_pdfs(tmp_path/"source.pdf",tmp_path/"rebuilt.pdf",tmp_path/"qa")
    assert result.page_count_match is False
    assert result.within_tolerance is False
