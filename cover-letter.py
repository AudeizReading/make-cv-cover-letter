import csv
import os
from datetime import datetime
from docx import Document
from docx.shared import Pt


def generate_docx_cover_letter(csv_file, output_file):
    """
    Lit le fichier CSV de la lettre de motivation et génère un document DOCX.
    """
    # Lecture du CSV
    data = {}
    with open(csv_file, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            data[row["title"].strip()] = row["content"].strip()

    # Création du document
    doc = Document()

    doc.add_paragraph()  # Ligne vide pour l'espacement

    # Expéditeur et Destinataire
    if "Expéditeur" in data:
        doc.add_paragraph(data["Expéditeur"])
    if "Destinataire" in data:
        doc.add_paragraph(data["Destinataire"])

    doc.add_paragraph()
    # Ajout de la date
    if "Date" in data:
        p = doc.add_paragraph()
        run = p.add_run("Nice, le ")
        doc.add_paragraph(data["Date"])
    else:
        p = doc.add_paragraph()
        run = p.add_run("Nice, le ")
        today = datetime.today().strftime("%d %B %Y")
        doc.add_paragraph(f"{today}")

    # Objet
    if "Objet" in data:
        p = doc.add_paragraph()
        run = p.add_run("Objet: ")
        run.bold = True
        p.add_run(data["Objet"])

    doc.add_paragraph()

    # Salutation
    if "Salutation" in data:
        p = doc.add_paragraph()
        run = p.add_run(data["Salutation"])
        doc.add_paragraph()
    # Corps de la lettre
    if "Corps" in data:
        # On remplace les "\n" par des retours à la ligne
        corps = data["Corps"].replace("\\n", "\n")
        doc.add_paragraph(corps)

    doc.add_paragraph()

    # Signature
    if "Formule de politesse" in data:
        p = doc.add_paragraph()
        run = p.add_run(data["Formule de politesse"])
        doc.add_paragraph()
    doc.add_paragraph()
    if "Signature" in data:
        doc.add_paragraph(data["Signature"])

    # Sauvegarde du document DOCX
    doc.save(output_file)
    print(f"Lettre de motivation générée et sauvegardée dans {output_file}")


if __name__ == "__main__":
    # Fichiers CSV d'entrée
    import sys

    if len(sys.argv) == 2:
        cover_letter_csv = sys.argv[1]
    else:
        print("Usage: python cv.py <cv_csv_file>")
        sys.exit(1)

    cover_letter_output = "cover-letter-test.docx"

    # Vérification de l'existence des fichiers CSV
    if os.path.exists(cover_letter_csv):
        generate_docx_cover_letter(cover_letter_csv, cover_letter_output)
    else:
        print(f"Le fichier {cover_letter_csv} n'existe pas.")
