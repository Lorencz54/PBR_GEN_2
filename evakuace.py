from excel_data import *
from placeholders import doc
import win32com.client as win32

def evakuace_obecne():
    bm_unikove_cesty_paragraphs = doc.Bookmarks("UNIKOVE_CESTY_PARAGRAPHS").Range
    bm_unikove_cesty_paragraphs.InsertAfter("Únik osob z objektu bude řešen ")
    bm_unikove_cesty_paragraphs.Collapse(0)
    bm_unikove_cesty_paragraphs.InsertAfter("současnou ")
    bm_unikove_cesty_paragraphs.Font.Bold = True
    bm_unikove_cesty_paragraphs.Collapse(0)
    bm_unikove_cesty_paragraphs.InsertAfter("evakuací skrze únikové cesty (dále jen „ÚC“) na volné prostranství. Počet unikajících osob z posuzovaného objektu jako celku je dle ČSN 73 0818 stanoven na ")
    bm_unikove_cesty_paragraphs.Font.Bold = False
    bm_unikove_cesty_paragraphs.Collapse(0)
    bm_unikove_cesty_paragraphs.InsertAfter("E = " + str(pocet_osob_count) + " os.")
    bm_unikove_cesty_paragraphs.Font.Bold = True
    bm_unikove_cesty_paragraphs.Collapse(0)
    bm_unikove_cesty_paragraphs.InsertParagraphAfter()
    bm_unikove_cesty_paragraphs.InsertParagraphAfter()
    bm_unikove_cesty_paragraphs.Font.Bold = False
    for list in d_PU_types:
        if "OB1" in list:
            bm_unikove_cesty_paragraphs.InsertAfter("Dle ČSN 73 0833, čl. 4.3 se v obytných buňkách budov skupiny OB1 považují za postačující nechráněné únikové cesty šířky 0,9 m s šířkou dveří na únikové cestě 0,8 m. Délky únikových cest se neposuzují.")
            bm_unikove_cesty_paragraphs.InsertParagraphAfter()

        break

def evakuace_garaz_I_obecne():
    bm_unikove_cesty_paragraphs = doc.Bookmarks("UNIKOVE_CESTY_PARAGRAPHS").Range
    bm_unikove_cesty_paragraphs.InsertParagraphAfter()
    if typ_garaze == "jednotlivá":
        bm_unikove_cesty_paragraphs.InsertAfter("Dle ČSN 73 0804, čl. I.6.1 se únikové cesty z jednotlivých garáží neposuzují.")
        bm_unikove_cesty_paragraphs.InsertParagraphAfter()
