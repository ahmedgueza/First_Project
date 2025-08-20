#!/usr/bin/env python3
"""
Create a minimal Word (.docx) file containing a plain-text transcript.
This uses only the Python standard library (zipfile) and does not require
third-party packages.

Output: /workspace/certificate_transcript.docx
"""

from zipfile import ZipFile, ZIP_DEFLATED
from datetime import date


TRANSCRIPT_LINES = [
    "Docteur Gouider Ben Hassouna Henda",
    "Spécialiste Oto‑Rhino‑Laryngologie",
    "",
    "CERTIFICAT MÉDICAL",
    "",
    "Je soussigné, Docteur en O.R.L., certifie que l'état de :",
    "Mme, Mlle, Mr : Khamessi Sabafia",  # best-effort reading from image
    "Adresse : __________________________",
    "",
    "1) Nécessite un arrêt de travail de : 1 jour (écrit à la main)",
    "   sauf complications à dater du : 05/06/2021 (écrit à la main)",
    "",
    "2) Nécessite une prolongation d'arrêt de travail de : ____________",
    "   sauf complications à dater du : ____________",
    "",
    "3) Lui permet de reprendre son travail à dater du : ____________",
    "   et ceci sauf complications.",
    "",
    "4) A consulté ce jour le : ____________",
    "",
    "Signature",
]


def build_document_xml(paragraphs: list[str]) -> str:
    def p(text: str) -> str:
        # Escape minimal XML entities
        text = (
            text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )
        return (
            "<w:p><w:r><w:t xml:space=\"preserve\">" + text + "</w:t></w:r></w:p>"
        )

    body = "".join(p(line) for line in paragraphs)
    return (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\""
        " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\""
        " xmlns:o=\"urn:schemas-microsoft-com:office:office\""
        " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
        " xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\""
        " xmlns:v=\"urn:schemas-microsoft-com:vml\""
        " xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\""
        " xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\""
        " xmlns:w10=\"urn:schemas-microsoft-com:office:word\""
        " xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
        " xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\""
        " xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\""
        " xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\""
        " xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\""
        " xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\""
        " mc:Ignorable=\"w14 wp14\"><w:body>" + body + "<w:sectPr/></w:body></w:document>"
    )


def main(output_path: str = "/workspace/certificate_transcript.docx") -> None:
    content_types_xml = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
        "</Types>"
    )

    rels_rels = (
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
        "</Relationships>"
    )

    word_document = build_document_xml(TRANSCRIPT_LINES)

    with ZipFile(output_path, "w", ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types_xml)
        z.writestr("_rels/.rels", rels_rels)
        z.writestr("word/document.xml", word_document)


if __name__ == "__main__":
    main()

