from word_replica.qa.word_render import (
    _fields_update_leaves_content_unchanged,
    _open_and_repair_changed_content,
)


class FakeContent:
    def __init__(self, text):
        self.Text = text


class OpenAndRepairDocument:
    def __init__(self, text):
        self.Content = FakeContent(text)
        self.close_calls = []

    def Close(self, SaveChanges=False):
        self.close_calls.append(SaveChanges)


class OpenAndRepairDocuments:
    """Returns different content for a normal open vs. OpenAndRepair=True,
    modeling a document Word would silently repair."""

    def __init__(self, *, normal_text, repaired_text):
        self.normal_text = normal_text
        self.repaired_text = repaired_text
        self.open_calls = []

    def Open(self, **kwargs):
        self.open_calls.append(kwargs)
        text = self.repaired_text if kwargs.get("OpenAndRepair") else self.normal_text
        return OpenAndRepairDocument(text)


class FakeApplication:
    def __init__(self, documents):
        self.Documents = documents


def test_open_and_repair_returns_false_when_content_matches(tmp_path):
    documents = OpenAndRepairDocuments(normal_text="Hello world", repaired_text="Hello world")
    app = FakeApplication(documents)
    source = tmp_path / "clean.docx"
    source.write_bytes(b"x")

    assert _open_and_repair_changed_content(app, source) is False
    assert documents.open_calls[0]["ReadOnly"] is True
    assert documents.open_calls[0]["OpenAndRepair"] is False
    assert documents.open_calls[1]["OpenAndRepair"] is True


def test_open_and_repair_returns_true_when_repair_changes_content(tmp_path):
    documents = OpenAndRepairDocuments(normal_text="Hello world", repaired_text="Hello  world")
    app = FakeApplication(documents)
    source = tmp_path / "corrupted.docx"
    source.write_bytes(b"x")

    assert _open_and_repair_changed_content(app, source) is True


class FieldsCollection:
    def __init__(self, on_update):
        self._on_update = on_update
        self.update_calls = 0

    def Update(self):
        self.update_calls += 1
        self._on_update()


class FieldsUpdateDocument:
    def __init__(self, before_text, after_text):
        self.Content = FakeContent(before_text)
        self._after_text = after_text
        self.Fields = FieldsCollection(self._apply_update)
        self.close_calls = []

    def _apply_update(self):
        self.Content = FakeContent(self._after_text)

    def Close(self, SaveChanges=False):
        self.close_calls.append(SaveChanges)


class FieldsUpdateDocuments:
    def __init__(self, before_text, after_text):
        self.before_text = before_text
        self.after_text = after_text
        self.open_calls = []
        self.last_doc = None

    def Open(self, **kwargs):
        self.open_calls.append(kwargs)
        self.last_doc = FieldsUpdateDocument(self.before_text, self.after_text)
        return self.last_doc


def test_fields_update_equality_true_when_cached_results_already_match(tmp_path):
    documents = FieldsUpdateDocuments(before_text="Page 3", after_text="Page 3")
    app = FakeApplication(documents)
    source = tmp_path / "consistent.docx"
    source.write_bytes(b"x")

    assert _fields_update_leaves_content_unchanged(app, source) is True
    assert documents.last_doc.Fields.update_calls == 1
    assert documents.last_doc.close_calls == [False]  # never saved


def test_fields_update_equality_false_when_update_changes_visible_text(tmp_path):
    documents = FieldsUpdateDocuments(before_text="Page 3", after_text="Page 7")
    app = FakeApplication(documents)
    source = tmp_path / "stale-fields.docx"
    source.write_bytes(b"x")

    assert _fields_update_leaves_content_unchanged(app, source) is False


def test_fields_update_opens_with_write_access_but_never_saves(tmp_path):
    documents = FieldsUpdateDocuments(before_text="x", after_text="x")
    app = FakeApplication(documents)
    source = tmp_path / "doc.docx"
    source.write_bytes(b"x")

    _fields_update_leaves_content_unchanged(app, source)

    assert documents.open_calls[0]["ReadOnly"] is False
    assert documents.open_calls[0]["AddToRecentFiles"] is False
