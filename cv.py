import csv
import os
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx_utils import (
    add_custom_heading,
    add_custom_paragraph,
    add_styled_run,
    set_paragraph_format,
)


def generate_cv(csv_file, output_file):
    """
    Lit le CSV de données de CV et génère un document RTF.
    """
    # Dictionnaires pour chaque section
    personal = []
    experiences = []
    education = []
    skills = []
    title = ""

    # Lecture du fichier CSV du CV
    with open(csv_file, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            section = row["section"].strip().lower()
            if section == "personal":
                personal.append(row)
            elif section == "experience":
                experiences.append(row)
            elif section == "education":
                education.append(row)
            elif section == "skills":
                skills.append(row)
            elif section == "title":
                title = row["description"].strip()

    # Construction du document RTF (structure minimale)
    # rtf = r"{\rtf1\ansi" + "\n"
    # Créer un document DOCX
    doc = Document()

    # Informations personnelles
    if personal:
        # doc.add_heading("Informations personnelles", level=2)
        for item in personal:
            # On tente d'afficher une information avec priorité description, content, puis subtitle
            info = (
                item.get("description")
                or item.get("content")
                or item.get("subtitle")
            )
            para_text = (
                f"{item['title']}: {info}"
                if item["title"] == "Tel"
                else f"{info}"
            )
            # doc.add_paragraph(para_text)
            add_custom_paragraph(
                doc,
                para_text,
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
                space_after=Pt(2),
                font_size=Pt(12),
                font_color=RGBColor(0x00, 0x00, 0x00),
            )
        # Titre du CV

    add_custom_heading(
        doc,
        title,
        level=1,
        font_size=Pt(24),
        color=RGBColor(0x00, 0x00, 0x00),
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
    )

    # Expériences professionnelles
    # if experiences:
    #     doc.add_heading("Expériences professionnelles", level=2)
    #     for exp in experiences:
    #         dates = ""
    #         if exp["start_date"] or exp["end_date"]:
    #             dates = f" ({exp['start_date']} - {exp['end_date']})"
    #         # Titre en gras suivi par l'entreprise et les dates
    #         p = doc.add_paragraph()
    #         run = p.add_run(f"{exp['title']} - {exp['subtitle']}{dates}")
    #         run.bold = True
    #         doc.add_paragraph(exp["description"])
    # Expériences professionnelles
    if experiences:
        add_custom_heading(
            doc,
            "Expériences professionnelles",
            level=2,
            font_size=Pt(18),
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
        )
        for exp in experiences:
            dates = ""
            if exp["start_date"] or exp["end_date"]:
                dates = f" ({exp['start_date']} - {exp['end_date']})"
            p = doc.add_paragraph()
            run = add_styled_run(
                p,
                f"{exp['title']} - {exp['subtitle']}{dates}",
                font_size=Pt(14),
                bold=True,
            )
            set_paragraph_format(p, space_after=Pt(4))
            para_desc = add_custom_paragraph(
                doc,
                exp["description"],
                alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                space_after=Pt(6),
            )
    # Formation
    if education:
        add_custom_heading(
            doc,
            "Formation",
            level=2,
            font_size=Pt(18),
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
        )
        for edu in education:
            dates = ""
            if edu["start_date"] or edu["end_date"]:
                dates = f" ({edu['start_date']} - {edu['end_date']})"
            p = doc.add_paragraph()
            run = add_styled_run(
                p,
                f"{edu['title']} - {edu['subtitle']}{dates}",
                font_size=Pt(14),
                bold=True,
            )
            set_paragraph_format(p, space_after=Pt(4))
            add_custom_paragraph(
                doc,
                edu["description"],
                alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                space_after=Pt(6),
            )
        # doc.add_heading("Formation", level=2)
        # for edu in education:
        #     dates = ""
        #     if edu["start_date"] or edu["end_date"]:
        #         dates = f" ({edu['start_date']} - {edu['end_date']})"
        #     p = doc.add_paragraph()
        #     run = p.add_run(f"{edu['title']} - {edu['subtitle']}{dates}")
        #     run.bold = True
        #     doc.add_paragraph(edu["description"])

    # Compétences
    if skills:
        add_custom_heading(
            doc,
            "Compétences",
            level=2,
            font_size=Pt(18),
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
        )
        for skill in skills:
            competence = f"- {skill['title']}: {skill['description']}"
            add_custom_paragraph(
                doc,
                competence,
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
                space_after=Pt(4),
                font_size=Pt(12),
            )
        # doc.add_heading("Compétences", level=2)
        # for skill in skills:
        #     competence = f"- {skill['title']}: {skill['description']}"
        #     doc.add_paragraph(competence)

    # Sauvegarde du document DOCX
    doc.save(output_file)
    print(f"CV généré et sauvegardé dans {output_file}")


if __name__ == "__main__":
    # Fichiers CSV d'entrée
    import sys

    if len(sys.argv) == 2:
        cv_csv = sys.argv[1]
    else:
        print("Usage: python cv.py <cv_csv_file>")
        sys.exit(1)

    cv_output = "cv-test.docx"

    # Vérification de l'existence des fichiers CSV
    if os.path.exists(cv_csv):
        generate_cv(cv_csv, cv_output)
    else:
        print(f"Le fichier {cv_csv} n'existe pas.")
