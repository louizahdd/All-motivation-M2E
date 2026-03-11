from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

def create_letter(filename, recipient_lines, institution_name_in_body):
    doc = Document()

    # Marges
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    def add_paragraph(text='', bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=6, first_indent=False):
        p = doc.add_paragraph()
        p.alignment = align
        pf = p.paragraph_format
        pf.space_before = Pt(space_before)
        pf.space_after = Pt(space_after)
        if first_indent:
            pf.first_line_indent = Cm(1)
        if text:
            run = p.add_run(text)
            run.bold = bold
            run.font.size = Pt(12)
            run.font.name = 'Times New Roman'
        return p

    # Expéditeur (gauche)
    sender_lines = [
        "Louiza Hadid",
        "06-50-37-56-47",
        "louizahdd@gmail.com",
        "3 rue de Ventspils",
        "56100 Lorient",
    ]
    for i, line in enumerate(sender_lines):
        p = add_paragraph(line, space_after=0)

    doc.add_paragraph()  # espace

    # Destinataire (droite)
    for line in recipient_lines:
        p = add_paragraph(line, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=0)

    doc.add_paragraph()  # espace

    # Objet
    add_paragraph("Objet : Candidature au Master M2E", bold=True, space_before=6, space_after=12)

    # Salutation
    add_paragraph("Madame, Monsieur,", space_after=6)

    # Corps du texte
    paragraphs = [
        "C'est une vocation pour l'enseignement, mûrie au fil des années, qui me pousse aujourd'hui à présenter ma candidature au Master M2E que vous proposez. Ce n'est pas un choix par défaut : votre établissement est celui dans lequel je me projette pleinement. La taille humaine des promotions, l'accompagnement individualisé et l'ancrage dans les valeurs de l'enseignement catholique correspondent à la vision de l'éducation que je porte depuis plusieurs années. C'est dans ce cadre, exigeant et bienveillant, que je souhaite préparer le CRPE et construire mon avenir dans l'enseignement du premier degré.",
        "Mon parcours en licence Sciences de l'éducation à l'Université de Rennes 2 a joué un rôle déterminant dans la construction de ce projet. J'y ai acquis des bases solides en pédagogie, en psychologie du développement et en didactique. Ces trois années m'ont confirmé que c'est dans l'enseignement du premier degré que je veux m'engager durablement, et votre formation représente pour moi l'étape la plus cohérente pour y parvenir.",
        "Ce projet ne s'est pas construit uniquement dans les livres. C'est sur le terrain que ma vocation s'est précisée, étape après étape. D'abord en animation périscolaire, où j'ai découvert qu'un cadre bienveillant compte autant que ce que l'on transmet. Puis en classe de mer, où j'ai compris que c'est la qualité du lien avec les élèves qui rend tout apprentissage possible. Enfin en tutorat auprès d'apprenants étrangers, où j'ai appris à reformuler, différencier et écouter vraiment. Chaque expérience a renforcé la même certitude : c'est dans ce métier que je veux grandir.",
        "Par ailleurs, mon poste actuel de chargée de recouvrement, bien qu'éloigné du milieu scolaire, m'a apporté une rigueur organisationnelle et une aisance dans la communication professionnelle qui seront des atouts pour enseigner. Gérer un portefeuille de comptes, prioriser les actions, travailler en équipe : autant de compétences transversales que je souhaite mettre au service de l'éducation.",
        f"Je suis profondément attachée à l'idée que l'enseignement ne se réduit pas à la transmission de savoirs : il s'agit d'accompagner chaque élève dans son développement, avec attention et exigence. C'est cette conviction qui me pousse à candidater en priorité à {institution_name_in_body}, un établissement dont les valeurs correspondent à ce que je veux incarner en tant qu'enseignante. Je suis déterminée à m'investir pleinement dans cette formation.",
    ]

    for para in paragraphs:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(6)
        pf.first_line_indent = Cm(1)
        run = p.add_run(para)
        run.font.size = Pt(12)
        run.font.name = 'Times New Roman'

    # Formule de politesse
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)
    pf.first_line_indent = Cm(1)
    run = p.add_run("Je me tiens à votre disposition pour un éventuel entretien et vous prie d'agréer, Madame, Monsieur, l'expression de mes respectueuses salutations.")
    run.font.size = Pt(12)
    run.font.name = 'Times New Roman'

    # Signature
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pf = p.paragraph_format
    pf.space_before = Pt(18)
    run = p.add_run("Louiza Hadid")
    run.bold = True
    run.font.size = Pt(12)
    run.font.name = 'Times New Roman'

    doc.save(filename)
    print(f"Created: {filename}")


# 1. INSPE Aix-Marseille - Marseille
create_letter(
    "lettre_INSPE_AixMarseille_Marseille.docx",
    [
        "INSPE d'Aix-Marseille",
        "Institut national supérieur du professorat",
        "et de l'éducation d'Aix-Marseille",
        "Site de Marseille",
        "Marseille (13)",
    ],
    "l'INSPE d'Aix-Marseille (site de Marseille)"
)

# 2. INSPE Aix-Marseille - Aix-en-Provence
create_letter(
    "lettre_INSPE_AixMarseille_Aix.docx",
    [
        "INSPE d'Aix-Marseille",
        "Institut national supérieur du professorat",
        "et de l'éducation d'Aix-Marseille",
        "Site d'Aix-en-Provence",
        "Aix-en-Provence cedex 01 (13)",
    ],
    "l'INSPE d'Aix-Marseille (site d'Aix-en-Provence)"
)

# 3. INSPE Lyon
create_letter(
    "lettre_INSPE_Lyon.docx",
    [
        "INSPE de Lyon",
        "Institut national supérieur du professorat",
        "et de l'éducation",
        "Université Claude Bernard - Lyon 1",
        "Lyon (69)",
    ],
    "l'INSPE de Lyon"
)

# 4. ISFEC Bretagne - Rennes
create_letter(
    "lettre_ISFEC_Bretagne_Rennes.docx",
    [
        "ISFEC Bretagne",
        "UCO – Facultés libres de l'Ouest",
        "Site de Rennes",
        "Rennes (35)",
    ],
    "l'ISFEC Bretagne (site de Rennes)"
)
