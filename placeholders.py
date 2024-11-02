from excel_data import *



def delete_paragraph(paragraph):
    """Remove a paragraph from the document."""
    p = paragraph._element
    p.getparent().remove(p)
    p._element = None

def replace_placeholder_in_run(run, replacements):
    """Replace placeholder text in a run using the replacements dictionary."""
    for placeholder, value in replacements.items():
        if placeholder in run.text:
            run.text = run.text.replace(placeholder, str(value))

def replace_text_in_paragraphs(paragraphs, replacements):
    """Replace text in a list of paragraphs based on a dictionary of replacements."""
    for paragraph in paragraphs:
        for run in paragraph.runs:
            replace_placeholder_in_run(run, replacements)

def remove_empty_paragraph(paragraph):
    """Remove paragraph if it's empty after replacements."""
    if not paragraph.text.strip():
        delete_paragraph(paragraph)

print("_________________________________________________replacing placeholders")
# Check if the Word file exists
if not os.path.exists(word_file):
    print(f"Word file '{word_file}' not found.")
else:
    document = Document(word_file)
    # Dictionary of placeholders and their corresponding replacement values
    replacements = {
        '[NAZEV_PROJEKTU]': nazev_projektu,
        '[MISTO_PROJEKTU]': misto_projektu,
        '[POZARNI_VYSKA]': pozarni_vyska,
        '[ZASTAVENA_PLOCHA]': zastavena_plocha,
        '[POCET_OSOB]': pocet_osob,
        '[JMENO_ZADAVATEL]': jmeno_zadavatel,
        '[ADRESA_ZADAVATEL]': adresa_zadavatel,
        '[JMENO_ZPRACOVATEL]': jmeno_zpracovatel,
        '[ADRESA_ZPRACOVATEL]': adresa_zpracovatel,
        '[ODPOVEDNY_PROJEKTANT]': odpovedny_projektant,
        '[OBOR_AUTORIZACE]': obor_autorizace,
        '[CISLO_AUTORIZACE]': cislo_autorizace,
        '[ZPRACOVATEL_PD]': zpracovatel_pd,
        '[TEL_ZPRACOVATEL]': tel_zpracovatel,
        '[MAIL_ZPRACOVATEL]': mail_zpracovatel,
        '[KATEGORIE]': kategorie,
        '[TRIDA]': trida_vyuziti,
        '[NAZEV_MISTA]': nazev_mista,
        '[KRAJ]': kraj,
        '[SPOLECNOST]': spolecnost,
        '[UCEL_STAVBY]': ucel_stavby,
        '[CHARAKTER_STAVBY]': charakter_stavby
    }

    # Replace placeholder text in document paragraphs
    replace_text_in_paragraphs(document.paragraphs, replacements)

    # Replace placeholder text in tables
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_text_in_paragraphs(cell.paragraphs, replacements)

    # Replace placeholder text in headers
    for section in document.sections:
        replace_text_in_paragraphs(section.header.paragraphs, replacements)

    # Remove paragraphs with specific texts in l_chosen_CSN
    for text_to_remove in l_chosen_CSN:
        for paragraph in document.paragraphs:
            if paragraph.text == text_to_remove:
                delete_paragraph(paragraph)

    document.save(output_word_path)

print("_________________________________________________all placeholders replaced")
print("_________________________________________________opening word")

word = win32.Dispatch("Word.Application")
doc = word.Documents.Open(output_word_path)

print("_________________________________________________word opened")