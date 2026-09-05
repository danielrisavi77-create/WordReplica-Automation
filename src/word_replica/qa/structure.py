from word_replica.domain.model import DocumentModel, Table


def table_shape(table: Table) -> list[list[int]]:
    return [
        [max(1, int(cell.properties.get("grid_span", 1))) for cell in row.cells]
        for row in table.rows
    ]


def referenced_story_slot_count(model: DocumentModel, kind: str, story_map: dict) -> int:
    property_name = f"{kind}_refs"
    declared = False
    count = 0
    for section in model.sections:
        if property_name not in section.properties:
            continue
        declared = True
        for ref in section.properties.get(property_name) or []:
            relationship = model.relationships.get(f"word/document.xml:{ref.get('rel_id')}")
            if relationship is not None and relationship.target in story_map:
                count += 1
    return count if declared else len(story_map)


def l1_projection(model: DocumentModel) -> dict:
    referenced_asset_ids = {drawing.asset_id for drawing in model.drawings}
    return {
        "body_kinds": [type(block).__name__ for block in model.body],
        "tables": [table_shape(block) for block in model.body if isinstance(block, Table)],
        "asset_hashes": sorted(
            asset.sha256
            for asset_id, asset in model.assets.items()
            if asset_id in referenced_asset_ids
        ),
        "headers": referenced_story_slot_count(model, "header", model.headers),
        "footers": referenced_story_slot_count(model, "footer", model.footers),
        "footnotes": len(model.footnotes),
        "endnotes": len(model.endnotes),
        "bookmarks": len(model.bookmarks),
        "fields": len(model.fields),
    }
