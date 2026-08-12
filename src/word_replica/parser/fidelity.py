from copy import deepcopy

from word_replica.domain.enums import FidelityMode
from word_replica.domain.model import DocumentModel


def project_fidelity(model: DocumentModel, mode: FidelityMode) -> DocumentModel:
    projected = deepcopy(model)
    if mode is FidelityMode.FULL:
        return projected

    projected.comments = {}
    projected.revisions = [revision for revision in projected.revisions if revision.kind == "insert"]
    projected.extras["comments"] = projected.comments
    projected.extras["revisions"] = projected.revisions
    projected.extras["tracked_changes_enabled"] = False

    for paragraph in projected.iter_paragraphs():
        paragraph.runs[:] = [run for run in paragraph.runs if not run.hidden]
    return projected
