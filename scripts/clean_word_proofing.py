from __future__ import annotations

import argparse
import os
import tempfile
import zipfile
from pathlib import Path

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def qn(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def ensure_child(parent: etree._Element, local: str) -> etree._Element:
    found = parent.find(f"w:{local}", namespaces=NS)
    if found is not None:
        return found
    child = etree.Element(qn(local))
    parent.append(child)
    return child


def patch_settings(root: etree._Element) -> int:
    changed = 0

    proof_state = root.find("w:proofState", namespaces=NS)
    if proof_state is None:
        proof_state = etree.Element(qn("proofState"))
        root.insert(0, proof_state)
        changed += 1
    for attr, value in (("spelling", "clean"), ("grammar", "clean")):
        if proof_state.get(qn(attr)) != value:
            proof_state.set(qn(attr), value)
            changed += 1

    for local in ("hideSpellingErrors", "hideGrammaticalErrors"):
        node = root.find(f"w:{local}", namespaces=NS)
        if node is None:
            root.insert(0, etree.Element(qn(local)))
            changed += 1
        elif node.get(qn("val"), "true").lower() in {"0", "false", "off"}:
            node.set(qn("val"), "true")
            changed += 1

    return changed


def patch_word_xml(root: etree._Element) -> int:
    changed = 0

    for node in list(root.xpath(".//w:proofErr", namespaces=NS)):
        parent = node.getparent()
        if parent is not None:
            parent.remove(node)
            changed += 1

    for run in root.iter(qn("r")):
        if run.find("w:rPr", namespaces=NS) is None:
            run.insert(0, etree.Element(qn("rPr")))
            changed += 1

    for rpr in root.xpath(".//w:rPr", namespaces=NS):
        node = rpr.find("w:noProof", namespaces=NS)
        if node is None:
            rpr.append(etree.Element(qn("noProof")))
            changed += 1
        elif node.get(qn("val"), "true").lower() in {"0", "false", "off"}:
            node.set(qn("val"), "true")
            changed += 1

    for doc_defaults in root.xpath(".//w:docDefaults", namespaces=NS):
        rpr_default = ensure_child(ensure_child(doc_defaults, "rPrDefault"), "rPr")
        if rpr_default.find("w:noProof", namespaces=NS) is None:
            rpr_default.append(etree.Element(qn("noProof")))
            changed += 1

    return changed


def word_xml_parts(base: Path) -> set[str]:
    word_dir = base / "word"
    parts = {
        "word/document.xml",
        "word/styles.xml",
        "word/footnotes.xml",
        "word/endnotes.xml",
        "word/comments.xml",
        "word/settings.xml",
    }
    if word_dir.exists():
        parts.update(f"word/{path.name}" for path in word_dir.glob("header*.xml"))
        parts.update(f"word/{path.name}" for path in word_dir.glob("footer*.xml"))
    return parts


def visible_text(docx_path: Path) -> str:
    text_parts = [
        "word/document.xml",
        "word/footnotes.xml",
        "word/endnotes.xml",
        "word/comments.xml",
    ]
    out: list[str] = []
    with zipfile.ZipFile(docx_path) as zf:
        names = set(zf.namelist())
        text_parts.extend(name for name in names if name.startswith("word/header") and name.endswith(".xml"))
        text_parts.extend(name for name in names if name.startswith("word/footer") and name.endswith(".xml"))
        for part in sorted(set(text_parts)):
            if part not in names:
                continue
            root = etree.fromstring(zf.read(part))
            out.extend(node.text or "" for node in root.iter(qn("t")))
            out.extend("\t" for _ in root.iter(qn("tab")))
            out.extend("\n" for _ in root.iter(qn("br")))
    return "".join(out)


def clean_docx(input_path: Path, output_path: Path, *, verify_text: bool = False) -> dict[str, int]:
    stats: dict[str, int] = {}
    before_text = visible_text(input_path) if verify_text else None
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as tmp_name:
        tmp_path = Path(tmp_name)
        with zipfile.ZipFile(input_path) as zf:
            zf.extractall(tmp_path)

        parser = etree.XMLParser(remove_blank_text=False, recover=False, resolve_entities=False, no_network=True)
        for part in sorted(word_xml_parts(tmp_path)):
            path = tmp_path / part
            if not path.exists():
                continue
            root = etree.parse(str(path), parser).getroot()
            changed = patch_settings(root) if part == "word/settings.xml" else patch_word_xml(root)
            if changed:
                path.write_bytes(
                    etree.tostring(
                        root,
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone=True,
                    )
                )
            stats[part] = changed

        with tempfile.TemporaryDirectory(prefix="word-proofing-", dir=output_path.parent) as output_tmp:
            temp_output = Path(output_tmp) / "cleaned.docx"
            with zipfile.ZipFile(temp_output, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                for path in tmp_path.rglob("*"):
                    if path.is_file():
                        zf.write(path, path.relative_to(tmp_path).as_posix())
            if verify_text and visible_text(temp_output) != before_text:
                raise ValueError("Visible text changed; original and destination were not replaced.")
            os.replace(temp_output, output_path)

    return stats


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Remove Word proofing underlines from a DOCX without changing visible text."
    )
    parser.add_argument("input", type=Path, help="Source .docx file")
    parser.add_argument("output", type=Path, nargs="?", help="Cleaned output .docx file")
    parser.add_argument("--in-place", action="store_true", help="Overwrite the input .docx")
    parser.add_argument("--verify-text", action="store_true", help="Fail if visible text changes")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = args.input.resolve()
    if not input_path.exists():
        raise SystemExit(f"Input not found: {input_path}")
    if input_path.suffix.lower() != ".docx":
        raise SystemExit("Input must be a .docx file")

    if args.in_place:
        if args.output:
            raise SystemExit("Do not provide an output path together with --in-place.")
        output_path = input_path
    elif args.output:
        output_path = args.output.resolve()
    else:
        output_path = input_path.with_name(f"{input_path.stem}_no-proofing-lines.docx")

    if output_path == input_path and not args.in_place:
        raise SystemExit("Use --in-place to overwrite the source document.")
    if output_path.suffix.lower() != ".docx":
        raise SystemExit("Output must be a .docx file")

    stats = clean_docx(input_path, output_path, verify_text=args.verify_text)

    if args.verify_text:
        print("visible_text_equal: True")

    print(f"output: {output_path}")
    for part, changed in sorted(stats.items()):
        print(f"{part}: {changed}")


if __name__ == "__main__":
    main()
