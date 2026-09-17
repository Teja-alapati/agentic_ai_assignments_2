import os
import re
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": WORD_NS}


def extract_word_text(docx_path):
    text_parts = []
    with ZipFile(docx_path, "r") as z:
        for filename in z.namelist():
            if filename.startswith("word/") and filename.endswith(".xml"):
                try:
                    root = etree.fromstring(z.read(filename))
                    for node in root.xpath(".//w:t", namespaces=NS):
                        if node.text:
                            text_parts.append(node.text)
                except Exception:
                    continue
    return " ".join(text_parts)


def extract_placeholders(docx_path):
    text = extract_word_text(docx_path)
    found = re.findall(r"\{\{\s*([^{}]+?)\s*\}\}", text)
    return list(dict.fromkeys(item.strip() for item in found))


def replace_in_xml(xml_data, replacements):
    try:
        root = etree.fromstring(xml_data)
    except Exception:
        return xml_data

    for paragraph in root.xpath(".//w:p", namespaces=NS):
        text_nodes = paragraph.xpath(".//w:t", namespaces=NS)
        if not text_nodes:
            continue

        full_text = "".join(node.text or "" for node in text_nodes)
        new_text = full_text

        for placeholder, value in replacements.items():
            new_text = new_text.replace(placeholder, str(value))

        if new_text != full_text:
            text_nodes[0].text = new_text
            for node in text_nodes[1:]:
                node.text = ""

    return etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )


def replace_text_in_docx(template_path, output_path, replacements):
    temporary_file = output_path + ".tmp"

    with ZipFile(template_path, "r") as source:
        with ZipFile(temporary_file, "w", ZIP_DEFLATED) as destination:
            for item in source.infolist():
                data = source.read(item.filename)

                if (
                    item.filename.startswith("word/")
                    and item.filename.endswith(".xml")
                ):
                    data = replace_in_xml(data, replacements)

                destination.writestr(item, data)

    if os.path.exists(output_path):
        os.remove(output_path)

    os.rename(temporary_file, output_path)


def generate_certificate(template_path, output_path, record, mapping):
    replacements = {}

    for placeholder, column in mapping.items():
        value = record.get(column, "") if column else ""
        if value is None:
            value = ""

        replacements["{{" + placeholder + "}}"] = str(value)

    replace_text_in_docx(template_path, output_path, replacements)
