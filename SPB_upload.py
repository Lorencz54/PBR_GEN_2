from placeholders import *

def upload_SPB_paragraphs():
    bm_PU_paragraphs = doc.Bookmarks("SPB_PARAGRAPHS").Range
    oznaceni_OB2_PUs = [sublist[1] for sublist in d_PU_types if sublist[0] == "OB2"]
    SPB_OB2_PUs = [sublist[4] for sublist in d_PU_types if sublist[0] == "OB2"]
    if len(oznaceni_OB2_PUs) > 0:
        nadpis_OB2_PU = "PÚ " + ', '.join(oznaceni_OB2_PUs) + " – byty"
        bm_PU_paragraphs.InsertAfter(nadpis_OB2_PU)
        bm_PU_paragraphs.Style = "Normální"
        bm_PU_paragraphs.Font.Bold = True
        bm_PU_paragraphs.InsertParagraphAfter()
        bm_PU_paragraphs.Collapse(0)
        bm_PU_paragraphs.InsertAfter("dle ČSN 73 0802, tab. 8, (požární výška objektu h = " + str(pozarni_vyska) + " m, " + str(k_system) + " konstrukční systém) je PÚ zařazen do ")
        bm_PU_paragraphs.Style = "Odstavec se seznamem"
        bm_PU_paragraphs.Collapse(0)
        bm_PU_paragraphs.InsertAfter(f"{', '.join(map(str, set(SPB_OB2_PUs)))} stupně požární bezpečnosti.")
        bm_PU_paragraphs.Font.Bold = True
        bm_PU_paragraphs.InsertParagraphAfter()
        bm_PU_paragraphs.Collapse(0)
        bm_PU_paragraphs_2 = bm_PU_paragraphs.Duplicate
        bm_PU_paragraphs_2.Style = "Normální"
        bm_PU_paragraphs.InsertParagraphAfter()
    for list in d_PU_types:
        if list[0] == "OB1":
            nadpis_PU = "PÚ " + list[1] + " – " + list[2]
            bm_PU_paragraphs.InsertAfter(nadpis_PU)
            bm_PU_paragraphs.Style = "Normální"
            bm_PU_paragraphs.Font.Bold = True
            bm_PU_paragraphs.InsertParagraphAfter()
            bm_PU_paragraphs.Collapse(0)
            if pocet_NP_obj == 1:
                clanek_4_1_1 = "a)"
            elif pocet_NP_obj <= 3 and (k_system == "nehořlavý" or k_system == "smíšený"):
                clanek_4_1_1 = "b)"
            elif pocet_NP_obj == 2 and k_system == "hořlavý":
                clanek_4_1_1 = "c)"
            else:
                clanek_4_1_1 = "d)"
            bm_PU_paragraphs.InsertAfter("dle ČSN 73 0833, čl. 4.1.1, pol. " + clanek_4_1_1 + " (RD má " + str(pocet_NP_obj) + " NP a " + str(k_system) + " konstrukční systém) je PÚ zařazen do ")
            bm_PU_paragraphs.Style = "Odstavec se seznamem"
            bm_PU_paragraphs.Collapse(0)
            bm_PU_paragraphs.InsertAfter(str(list[4]) + " stupně požární bezpečnosti.")
            bm_PU_paragraphs.Font.Bold = True
            bm_PU_paragraphs.InsertParagraphAfter()
            bm_PU_paragraphs.Collapse(0)
            bm_PU_paragraphs_2 = bm_PU_paragraphs.Duplicate
            bm_PU_paragraphs_2.Style = "Normální"
            if list == d_PU_types[-1]:
                pass
            else:
                bm_PU_paragraphs.InsertParagraphAfter()
        elif list[0] == "nevýrobní" or list[0] == "nevýrobní (pv)" or list[0] == "garáž I" or list[0] == "garáž III":
            nadpis_PU = "PÚ " + list[1] + " – " + list[2]
            bm_PU_paragraphs.InsertAfter(nadpis_PU)
            bm_PU_paragraphs.Style = "Normální"
            bm_PU_paragraphs.Font.Bold = True
            bm_PU_paragraphs.InsertParagraphAfter()
            bm_PU_paragraphs.Collapse(0)
            bm_PU_paragraphs.InsertAfter("dle ČSN 73 0802, tab. 8, (požární výška objektu h = " + str(pozarni_vyska) + " m, " + str(k_system) + " konstrukční systém) je PÚ zařazen do ")
            bm_PU_paragraphs.Style = "Odstavec se seznamem"
            bm_PU_paragraphs.Collapse(0)
            bm_PU_paragraphs.InsertAfter(str(list[4]) + " stupně požární bezpečnosti.")
            bm_PU_paragraphs.Font.Bold = True
            bm_PU_paragraphs.InsertParagraphAfter()
            bm_PU_paragraphs.Collapse(0)
            bm_PU_paragraphs_2 = bm_PU_paragraphs.Duplicate
            bm_PU_paragraphs_2.Style = "Normální"
            if list == d_PU_types[-1]:
                pass
            else:
                bm_PU_paragraphs.InsertParagraphAfter()
        elif list[0] == "OB2":
            pass