from placeholders import *

def upload_popis_zmen_paragraph():
    bm_popis_zmen_paragraphs = doc.Bookmarks("POPIS_NAVRZENYCH_ZMEN_PARAGRAPHS").Range
    bm_popis_zmen_paragraphs.InsertParagraphAfter()
    bm_popis_zmen_paragraphs.Collapse(0)
    bm_popis_zmen_paragraphs.InsertAfter("Popis navržených změn")
    bm_popis_zmen_paragraphs.Style = "Podkapitola 1"
    bm_popis_zmen_paragraphs.Collapse(0)
    bm_popis_zmen_paragraphs.InsertParagraphAfter()
    bm_popis_zmen_paragraphs.Collapse(0)
    bm_popis_zmen_paragraphs.Style = "Normální"
    bm_popis_zmen_paragraphs.InsertAfter(popis_navrzenych_zmen)
    bm_popis_zmen_paragraphs.InsertParagraphAfter()

def insert_zmeny_konstrukci_table():
    doc_table_source = word.Documents.Open(r'"C:\Users\Lenovo\OneDrive\Projekty\Skamba\sablony\PBR\dílčí šablony\ZS_1_TZ.docx')
    table_zmeny_konstrukci = doc_table_source.Tables(2)
    table_zmeny_konstrukci.Range.Copy()
    bm_zmeny_konstrukci = doc.Bookmarks("TABULKA_ZMENY_KONSTRUKCI").Range
    bm_zmeny_konstrukci.Paste()  # Paste the table
    doc_table_source.Close()
    print("table copied")