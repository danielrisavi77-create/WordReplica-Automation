from word_replica.domain.model import DocumentModel, Paragraph, Table


def _walk_blocks(blocks, prefix: str):
    for index, block in enumerate(blocks):
        path = f"{prefix}/{index}"
        if isinstance(block, Paragraph):
            yield path, block.text()
        elif isinstance(block, Table):
            for r_index, row in enumerate(block.rows):
                for c_index, cell in enumerate(row.cells):
                    yield from _walk_blocks(cell.blocks, f"{path}/row/{r_index}/cell/{c_index}")


def l0_projection(model: DocumentModel) -> list[tuple[str, str]]:
    rows = list(_walk_blocks(model.body, "body"))
    for note_id, blocks in sorted(model.footnotes.items(), key=lambda item: str(item[0])):
        rows.extend(_walk_blocks(blocks, f"footnote/{note_id}"))
    for note_id, blocks in sorted(model.endnotes.items(), key=lambda item: str(item[0])):
        rows.extend(_walk_blocks(blocks, f"endnote/{note_id}"))
    return rows
