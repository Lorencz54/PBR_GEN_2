from excel_data import *
from placeholders import doc

def vnitrni_odberne_mista():
    bm_vnitrni_odberne_mista_paragraphs = doc.Bookmarks("VNITRNI_ODBERNE_MISTA_PARAGRAPHS").Range
    for list in d_PU_types:
        if "OB1" in list:
            if pocet_osob_count <= 20:
                bm_vnitrni_odberne_mista_paragraphs.InsertAfter("Dle ČSN 73 0873, čl. 4.4 b5) není nutno v posuzovaném objektu zřizovat vnitřní odběrné místo (celkový počet osob v prostorech OB1 není větší než 20).")
                bm_vnitrni_odberne_mista_paragraphs.InsertParagraphAfter()
                bm_vnitrni_odberne_mista_paragraphs.Collapse(0)
        break

def PHP():
    bm_PHP_paragraphs = doc.Bookmarks("PHP_PARAGRAPHS").Range
    for list in d_PU_types:
        nadpis_PU = list[1] + " – " + list[2]
        bm_PHP_paragraphs.InsertAfter(nadpis_PU)
        bm_PHP_paragraphs.Style = "Normální"
        bm_PHP_paragraphs.Font.Bold = True
        bm_PHP_paragraphs.InsertParagraphAfter()
        bm_PHP_paragraphs.Collapse(0)
        if list[0] == "OB1":
            bm_PHP_paragraphs.InsertAfter("dle vyhl. č. 23/2008 Sb., přílohy 1 musí být objekt RD (" + nadpis_PU + ") vybaven alespoň jedním PHP s hasící schopností nejméně 34A. Bude se jednat o 1ks PHP práškový. PHP se doporučuje umístit ")
            bm_PHP_paragraphs.Style = "Odstavec se seznamem"
            bm_PHP_paragraphs.Collapse(0)

            bm_PHP_paragraphs.InsertAfter("v 1.NP, v předsíni (m.č. 1.01)")
            bm_PHP_paragraphs.HighlightColorIndex = 6
            bm_PHP_paragraphs.Collapse(0)

            bm_PHP_paragraphs.InsertAfter(". V případě umístění PHP na přehledné pozici nebudou PHP označeny fotoluminiscenčním značením.")
            bm_PHP_paragraphs.HighlightColorIndex = 0
            bm_PHP_paragraphs.InsertParagraphAfter()

            if garaz_soucasti_RD == "ANO":
                bm_PHP_paragraphs.InsertAfter("dle ČSN 73 0833, čl. 4.5 je dále doporučeno instalovat 1ks PHP 34A/183B v prostoru garáže.")
                bm_PHP_paragraphs.Collapse(0)

                bm_PHP_paragraphs.InsertParagraphAfter()
                bm_PHP_paragraphs.Collapse(0)

                bm_PU_paragraphs_2 = bm_PHP_paragraphs.Duplicate
                bm_PU_paragraphs_2.Style = "Normální"

        elif list[0] == "garáž I":
            bm_PHP_paragraphs.InsertAfter("dle ČSN 73 0804, čl. I.7.3 musí být v jednotlivých garážích instalovány přenosné hasicí přístroje (PHP) s hasicí schopností min. 183 B pro každý samostatně oddělený prostor (stání) ")
            bm_PHP_paragraphs.Style = "Odstavec se seznamem"
            bm_PHP_paragraphs.Collapse(0)
            bm_PHP_paragraphs.InsertAfter("– celkem 1ks")
            bm_PHP_paragraphs.Font.Bold = True
            bm_PHP_paragraphs.Collapse(0)
            if list == d_PU_types[-1]:
                pass
            else:
                bm_PHP_paragraphs.InsertParagraphAfter()