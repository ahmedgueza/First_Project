#!/usr/bin/env python3
"""
Generate a clean, printable French medical certificate (.docx) matching the layout
from the provided photo. The document contains clearly underlined blanks that can be
filled in digitally or printed and completed by hand.

Usage:
  python3 generate_certificate.py \
    --doctor "Docteur Gouider Ben Hassouna Henda" \
    --specialty "Spécialiste Oto‑Rhino‑Laryngologie" \
    --patient_title "Mme, Mlle, Mr" \
    --output certificate_template.docx

Optional arguments allow pre-filling certain fields.
"""

import argparse
from datetime import date
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE


def add_heading_run(paragraph, text: str, bold: bool = False, size: int = 12):
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    return run


def add_labeled_blank(doc: Document, label: str, preset: str | None = None):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(6)
    add_heading_run(paragraph, label + " ", bold=False, size=12)
    run = paragraph.add_run(preset if preset else "____________________________")
    if not preset:
        run.underline = True


def generate_certificate(
    output_path: str,
    doctor_name: str,
    specialty: str,
    patient_title_prefix: str,
    patient_name: str | None,
    address: str | None,
    rest_days_text: str | None,
    rest_from_date: str | None,
    prolongation_text: str | None,
    prolongation_from_date: str | None,
    return_to_work_date: str | None,
    consulted_on_date: str | None,
):
    doc = Document()

    # Header
    header = doc.add_paragraph()
    header.alignment = WD_ALIGN_PARAGRAPH.LEFT
    add_heading_run(header, doctor_name + "\n", bold=True, size=14)
    add_heading_run(header, specialty + "\n", bold=False, size=11)

    # Title
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_heading_run(title, "CERTIFICAT MÉDICAL", bold=True, size=16)

    doc.add_paragraph("")

    # Intro
    intro = doc.add_paragraph()
    intro.paragraph_format.space_after = Pt(10)
    add_heading_run(
        intro,
        "Je soussigné, Docteur en O.R.L., certifie que l'état de :",
        bold=False,
        size=12,
    )

    add_labeled_blank(
        doc,
        f"{patient_title_prefix} :",
        preset=patient_name,
    )
    add_labeled_blank(doc, "Adresse :", preset=address)

    # Items 1-4
    p1 = doc.add_paragraph()
    add_heading_run(p1, "1) Nécessite un arrêt de travail de : ", size=12)
    p1_run = p1.add_run(rest_days_text if rest_days_text else "____________")
    if not rest_days_text:
        p1_run.underline = True
    p1.add_run(" sauf complications à dater du : ")
    p1_from = p1.add_run(rest_from_date if rest_from_date else "____________")
    if not rest_from_date:
        p1_from.underline = True

    p2 = doc.add_paragraph()
    add_heading_run(
        p2,
        "2) Nécessite une prolongation d'arrêt de travail de : ",
        size=12,
    )
    p2_run = p2.add_run(prolongation_text if prolongation_text else "____________")
    if not prolongation_text:
        p2_run.underline = True
    p2.add_run(" sauf complications à dater du : ")
    p2_from = p2.add_run(
        prolongation_from_date if prolongation_from_date else "____________"
    )
    if not prolongation_from_date:
        p2_from.underline = True

    p3 = doc.add_paragraph()
    add_heading_run(
        p3, "3) Lui permet de reprendre son travail à dater du : ", size=12
    )
    p3_run = p3.add_run(return_to_work_date if return_to_work_date else "____________")
    if not return_to_work_date:
        p3_run.underline = True
    p3.add_run(
        " et ceci sauf complications."
    )

    p4 = doc.add_paragraph()
    add_heading_run(p4, "4) A consulté ce jour le : ", size=12)
    p4_run = p4.add_run(
        consulted_on_date if consulted_on_date else date.today().strftime("%d/%m/%Y")
    )
    if not consulted_on_date:
        # The auto-filled date should not be underlined
        pass

    doc.add_paragraph("")

    # Footer: signature
    footer = doc.add_paragraph()
    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_heading_run(footer, "Signature\n", bold=False, size=12)
    footer.add_run("\n\n\n")
    underline = footer.add_run("__________________________")
    underline.underline = True

    doc.save(output_path)


def parse_args():
    parser = argparse.ArgumentParser(description="Generate a medical certificate (.docx)")
    parser.add_argument("--output", default="certificate_template.docx")
    parser.add_argument("--doctor", default="Docteur Gouider Ben Hassouna Henda")
    parser.add_argument(
        "--specialty", default="Spécialiste Oto‑Rhino‑Laryngologie"
    )
    parser.add_argument("--patient_title", default="Mme, Mlle, Mr")
    parser.add_argument("--patient_name", default=None)
    parser.add_argument("--address", default=None)
    parser.add_argument("--rest_days", default=None, help="ex: 1 jour / 3 jours")
    parser.add_argument("--rest_from", default=None, help="ex: 05/06/2021")
    parser.add_argument("--prolongation", default=None, help="ex: 2 jours")
    parser.add_argument("--prolongation_from", default=None, help="ex: 07/06/2021")
    parser.add_argument("--return_to_work", default=None, help="ex: 08/06/2021")
    parser.add_argument("--consulted_on", default=None, help="ex: 05/06/2021")
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    generate_certificate(
        output_path=args.output,
        doctor_name=args.doctor,
        specialty=args.specialty,
        patient_title_prefix=args.patient_title,
        patient_name=args.patient_name,
        address=args.address,
        rest_days_text=args.rest_days,
        rest_from_date=args.rest_from,
        prolongation_text=args.prolongation,
        prolongation_from_date=args.prolongation_from,
        return_to_work_date=args.return_to_work,
        consulted_on_date=args.consulted_on,
    )

