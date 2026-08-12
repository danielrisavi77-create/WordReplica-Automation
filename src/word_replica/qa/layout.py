from word_replica.domain.model import DocumentModel

_LAYOUT_KEYS = (
    "width",
    "height",
    "orientation",
    "margin_top",
    "margin_right",
    "margin_bottom",
    "margin_left",
    "header_distance",
    "footer_distance",
    "gutter",
    "columns",
    "column_space",
    "page_number_start",
)


def normalize_section_properties(properties: dict) -> dict:
    return {key: properties.get(key) for key in _LAYOUT_KEYS if key in properties}


def l3_projection(model: DocumentModel) -> list[dict]:
    return [normalize_section_properties(section.properties) for section in model.sections]
