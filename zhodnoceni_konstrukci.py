from excel_data import *
from placeholders import doc

def upload_konstrukce_paragraphs():
    for list in d_PU_types:
        if "OB1" in list:
            bm_nosne_konstrukce_paragraphs = doc.Bookmarks("NOSNE_KONSTRUKCE_PARAGRAPHS").Range
            bm_pozarni_steny_paragraphs = doc.Bookmarks("POZARNI_STENY_PARAGRAPHS").Range
            if pocet_NP_obj > 1:
                bm_nosne_konstrukce_paragraphs.InsertAfter("Dle ČSN 73 0833, čl. 4.2.1 musí stropní konstrukce uvnitř vícepodlažního PÚ obytné buňky v nadzemních podlažích vykazovat požární odolnost alespoň 15 minut.")
                bm_nosne_konstrukce_paragraphs.InsertParagraphAfter()
            if sousedni_objekt == "ANO RD":
                bm_pozarni_steny_paragraphs.InsertAfter("PŘEVÝŠENÍ STŘEŠNÍHO PLÁŠTĚ dle ČSN 73 0802, čl. 8.2.4")
                bm_pozarni_steny_paragraphs.HighlightColorIndex = 7
                bm_pozarni_steny_paragraphs.InsertParagraphAfter()
            elif sousedni_objekt == "NE":
                bm_pozarni_steny_paragraphs.InsertAfter("Dle ČSN 73 0833, čl. 4.2.2, resp. ČSN 73 0802:2009, čl. 8.2.4 nemusí povrch střešního pláště převyšovat o 300 mm. Objekt stojí samostatně.")
                bm_pozarni_steny_paragraphs.InsertParagraphAfter()
        break

def pozarni_pasy():
    bm_pozarni_pasy_paragraphs = doc.Bookmarks("POZARNI_PASY_PARAGRAPHS").Range
    if sousedni_objekt == "ANO RD":
        bm_pozarni_pasy_paragraphs.InsertAfter("Dle ČSN 73 0833, čl. 4.2.3 nemusí být u styku posuzovaného objektu OB1 se sousedním objektem OB1 v obvodových stěnách zřízeny požární pásy.")
        bm_pozarni_pasy_paragraphs.InsertParagraphAfter()
    elif sousedni_objekt == "NE":  
        if pozarni_vyska_raw <= 12:
            bm_pozarni_pasy_paragraphs.InsertAfter("Dle ČSN 73 0802, čl. 8.4.10 c) nemusí být zřízeny požární pásy s výjimkou svislých požárních pásů mezi objekty, které se nenavrhují. Objekt stojí samostatně.")
            bm_pozarni_pasy_paragraphs.InsertParagraphAfter()
    elif sousedni_objekt == "ANO":
        if pozarni_vyska_raw <= 12:
            bm_pozarni_pasy_paragraphs.InsertAfter("Dle ČSN 73 0802, čl. 8.4.10 c) nemusí být zřízeny požární pásy s výjimkou svislých požárních pásů mezi objekty, které se navrhují.")
            bm_pozarni_pasy_paragraphs.InsertParagraphAfter()
            bm_pozarni_pasy_paragraphs.InsertParagraphAfter()
            bm_pozarni_pasy_paragraphs.Collapse(0)
            bm_pozarni_pasy_paragraphs.InsertAfter("Je navržen svislý požární pás mezi řešeným objektem RD a sousedním objektem jednotlivé garáže v šířce 0,90 m. požární pás bude proveden z cihel POROTHERM tl. 450 mm a vnějšího zateplení z minerální vaty. Skutečná požární odolnost konstrukce (REI 180 DP1) vyhovuje požadavku odolnosti pro svislý požární pás PÚ ve II.SPB (REI 30 DP1).")
            bm_pozarni_pasy_paragraphs.HighlightColorIndex = 6
            bm_pozarni_pasy_paragraphs.InsertParagraphAfter()
def konstrukce_table_insert():
    table = doc.Tables(3)
    constants = win32.constants
    wdAlignParagraphCenter = 1
    for konstrukce in range(len(l_odolnost_svisle_nosne)):
        table.Rows.Add()
        last_row_index = table.Rows.Count
        last_col_index = table.Columns.Count
        cell_1 = table.Cell(last_row_index, last_col_index - 1)
        cell_2 = table.Cell(last_row_index, last_col_index - 2)
        cell_1.Range.Text = l_odolnost_svisle_nosne[konstrukce]
        cell_2.Range.Text = l_svisle_nosne[konstrukce].value
        cell_1.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    for konstrukce in range(len(l_odolnost_vodorovne_nosne)):
        table.Rows.Add()
        last_row_index = table.Rows.Count
        last_col_index = table.Columns.Count
        cell_1 = table.Cell(last_row_index, last_col_index - 1)
        cell_2 = table.Cell(last_row_index, last_col_index - 2)
        cell_1.Range.Text = l_odolnost_vodorovne_nosne[konstrukce]
        cell_2.Range.Text = l_vodorovne_nosne[konstrukce].value
        cell_1.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    for konstrukce in range(len(l_odolnost_svisle_nenosne)):
        table.Rows.Add()
        last_row_index = table.Rows.Count
        last_col_index = table.Columns.Count
        cell_1 = table.Cell(last_row_index, last_col_index - 1)
        cell_2 = table.Cell(last_row_index, last_col_index - 2)
        cell_1.Range.Text = l_odolnost_svisle_nenosne[konstrukce]
        cell_2.Range.Text = l_svisle_nenosne[konstrukce].value
        cell_1.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    for konstrukce in range(len(l_odolnost_vodorovne_nenosne)):
        table.Rows.Add()
        last_row_index = table.Rows.Count
        last_col_index = table.Columns.Count
        cell_1 = table.Cell(last_row_index, last_col_index - 1)
        cell_2 = table.Cell(last_row_index, last_col_index - 2)
        cell_1.Range.Text = l_odolnost_vodorovne_nenosne[konstrukce]
        cell_2.Range.Text = l_vodorovne_nenosne[konstrukce].value
        cell_1.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    for konstrukce in range(len(l_odolnost_konstrukce_strechy)):
        table.Rows.Add()
        last_row_index = table.Rows.Count
        last_col_index = table.Columns.Count
        cell_1 = table.Cell(last_row_index, last_col_index - 1)
        cell_2 = table.Cell(last_row_index, last_col_index - 2)
        cell_1.Range.Text = l_odolnost_konstrukce_strechy[konstrukce]
        cell_2.Range.Text = l_konstrukce_strechy[konstrukce].value
        cell_1.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    for konstrukce in range(len(l_odolnost_schodiste)):
        table.Rows.Add()
        last_row_index = table.Rows.Count
        last_col_index = table.Columns.Count
        cell_1 = table.Cell(last_row_index, last_col_index - 1)
        cell_2 = table.Cell(last_row_index, last_col_index - 2)
        cell_1.Range.Text = l_odolnost_schodiste[konstrukce]
        cell_2.Range.Text = l_schodiste[konstrukce].value
        cell_1.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
