from word_replica.domain.model import DocumentModel, Table


def table_shape(table: Table) -> list[list[int]]:
    return [
        [max(1, int(cell.properties.get("grid_span", 1))) for cell in row.cells]
        for row in table.rows
    ]


def l1_projection(model: DocumentModel) -> dict:
    return {
        "body_kinds": [type(block).__name__ for block in model.body],
        "tables": [table_shape(block) for block in model.body if isinstance(block, Table)],
        "asset_hashes": sorted(asset.sha256 for asset in model.assets.values()),
        "headers": len(model.headers),
        "footers": len(model.footers),
        "footnotes": len(model.footnotes),
        "endnotes": len(model.endnotes),
        "bookmarks": len(model.bookmarks),
        "fields": len(model.fields),
    }
