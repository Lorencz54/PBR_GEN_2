from excel_data import *
from placeholders import doc

def popis_projektu():
    bm_predmet_pbr = doc.Bookmarks("PREDMET_PBR_PARAGRAPH").Range
    bm_umisteni_objektu = doc.Bookmarks("UMISTENI_OBJEKTU_PARAGRAPH").Range
    bm_zakladni_popis_obj = doc.Bookmarks("ZAKLADNI_POPIS_OBJEKTU_PARAGRAPH").Range
    bm_predmet_pbr.Text = ""
    bm_predmet_pbr.InsertAfter(predmet_PBR)
    bm_umisteni_objektu.InsertAfter(umisteni_obj)
    bm_zakladni_popis_obj.InsertAfter(zakladni_popis_obj)
    bm_zakladni_popis_obj.InsertParagraphAfter()