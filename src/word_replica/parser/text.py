from word_replica.parser.nodes import local_name

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def run_text(run_node) -> str:
    pieces: list[str] = []
    for child in run_node:
        local = local_name(child)
        if local in {"t", "delText"}:
            pieces.append(child.text or "")
        elif local == "tab":
            pieces.append("\t")
        elif local in {"br", "cr"}:
            pieces.append("\n")
    return "".join(pieces)


def bool_prop(parent, name: str) -> bool | None:
    node = parent.find(f"w:{name}", namespaces=NS) if parent is not None else None
    if node is None:
        return None
    value = node.get(f"{{{W_NS}}}val")
    return value not in {"0", "false", "off", "none"}
