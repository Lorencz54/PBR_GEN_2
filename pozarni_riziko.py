from placeholders import *

def mezni_rozmery_PU():
    bm_mezni_rozmery_PU_paragraph = doc.Bookmarks("MEZNI_ROZMERY_PARAGRAPHS").Range
    oznaceni_OB2_PUs = [sublist[1] for sublist in d_PU_types if sublist[0] == "OB2"]
    if len(oznaceni_OB2_PUs) > 0:
        nadpis_OB2_PU = "PÚ " + ', '.join(oznaceni_OB2_PUs) + " – byty"
        bm_mezni_rozmery_PU_paragraph.InsertAfter(nadpis_OB2_PU)
        bm_mezni_rozmery_PU_paragraph.Style = "Normální"
        bm_mezni_rozmery_PU_paragraph.Font.Bold = True
        bm_mezni_rozmery_PU_paragraph.InsertParagraphAfter()
        bm_mezni_rozmery_PU_paragraph.Collapse(0)
        bm_mezni_rozmery_PU_paragraph.InsertAfter("dle ČSN 73 0833, čl. 5.1.5 se mezní rozměry PÚ s obytnými buňkami a domovním vybavením nestanovují")
        bm_mezni_rozmery_PU_paragraph.Style = "Odstavec se seznamem"
        bm_mezni_rozmery_PU_paragraph.InsertParagraphAfter()
        bm_mezni_rozmery_PU_paragraph.Collapse(0)
        bm_mezni_rozmery_PU_paragraph_2 = bm_mezni_rozmery_PU_paragraph.Duplicate
        bm_mezni_rozmery_PU_paragraph_2.Style = "Normální"
    for list in d_PU_types:
        if list[0] == "OB1":
            nadpis_PU = "PÚ " + list[1] + " – " + list[2]
            bm_mezni_rozmery_PU_paragraph.InsertAfter(nadpis_PU)
            bm_mezni_rozmery_PU_paragraph.Style = "Normální"
            bm_mezni_rozmery_PU_paragraph.Font.Bold = True
            bm_mezni_rozmery_PU_paragraph.InsertParagraphAfter()
            bm_mezni_rozmery_PU_paragraph.Collapse(0)
            bm_mezni_rozmery_PU_paragraph.InsertAfter("Mezní rozměry PÚ jsou dány mezními rozměry pro budovy skupiny OB1. Dle ČSN 73 0833, čl. 3.5 a) ")
            bm_mezni_rozmery_PU_paragraph.Style = "Odstavec se seznamem"
            bm_mezni_rozmery_PU_paragraph.Collapse(0)
            bm_mezni_rozmery_PU_paragraph.InsertAfter("nejsou ")
            bm_mezni_rozmery_PU_paragraph.Font.Bold = True
            bm_mezni_rozmery_PU_paragraph.Collapse(0)
            bm_mezni_rozmery_PU_paragraph.InsertAfter("překročeny mezní rozměry PÚ ani mezní počty podlaží. Rozměry PÚ ")
            bm_mezni_rozmery_PU_paragraph.Font.Bold = False
            bm_mezni_rozmery_PU_paragraph.Collapse(0)
            bm_mezni_rozmery_PU_paragraph.InsertAfter("vyhovují ")
            bm_mezni_rozmery_PU_paragraph.Font.Bold = True
            bm_mezni_rozmery_PU_paragraph.Collapse(0)
            bm_mezni_rozmery_PU_paragraph.InsertAfter("požadavkům normy. Skutečné rozměry PÚ jsou menší nebo rovny největším dovoleným rozměrům dle ČSN 73 0833, čl. 3.5.")
            bm_mezni_rozmery_PU_paragraph.Font.Bold = False
            bm_mezni_rozmery_PU_paragraph.InsertParagraphAfter()
            bm_mezni_rozmery_PU_paragraph.Collapse(0)
            bm_mezni_rozmery_PU_paragraph_2 = bm_mezni_rozmery_PU_paragraph.Duplicate
            bm_mezni_rozmery_PU_paragraph_2.Style = "Normální"
        if list[0] == "garáž I":
            nadpis_PU = "PÚ " + list[1] + " – " + list[2]
            bm_mezni_rozmery_PU_paragraph.InsertAfter(nadpis_PU)
            bm_mezni_rozmery_PU_paragraph.Style = "Normální"
            bm_mezni_rozmery_PU_paragraph.Font.Bold = True
            bm_mezni_rozmery_PU_paragraph.InsertParagraphAfter()
            bm_mezni_rozmery_PU_paragraph.Collapse(0)
            bm_mezni_rozmery_PU_paragraph.InsertAfter("Mezní rozměry PÚ jsou dány mezním počtem stání pro jednotlivé garáže. Dle ČSN 73 0804, čl. I.2.3 a) je stanoven mezní počet stání 3. Řešený objekt je navržen ")
            bm_mezni_rozmery_PU_paragraph.Style = "Odstavec se seznamem"
            if pocet_stani == 1:
                pocet_stani_text = "s jedním "
            elif pocet_stani == 2:
                pocet_stani_text = "se dvěma "
            elif pocet_stani == 3:
                pocet_stani_text = "se třemi "
            bm_mezni_rozmery_PU_paragraph.InsertAfter(pocet_stani_text)
            bm_mezni_rozmery_PU_paragraph.Collapse(0)
            bm_mezni_rozmery_PU_paragraph.InsertAfter("parkovacími stáními ")
            bm_mezni_rozmery_PU_paragraph.Font.Bold = False
            bm_mezni_rozmery_PU_paragraph.Collapse(0)
            bm_mezni_rozmery_PU_paragraph.InsertAfter("– vyhovuje ")
            bm_mezni_rozmery_PU_paragraph.Font.Bold = True
            bm_mezni_rozmery_PU_paragraph.Collapse(0)
            bm_mezni_rozmery_PU_paragraph.InsertParagraphAfter()
            bm_mezni_rozmery_PU_paragraph.InsertAfter("Dle ČSN 73 0804, čl. I.2.3 a) ")
            bm_mezni_rozmery_PU_paragraph.Font.Bold = False
            bm_mezni_rozmery_PU_paragraph.Collapse(0)
            bm_mezni_rozmery_PU_paragraph.InsertAfter("nejsou ")
            bm_mezni_rozmery_PU_paragraph.Font.Bold = True
            bm_mezni_rozmery_PU_paragraph.Collapse(0)
            bm_mezni_rozmery_PU_paragraph.InsertAfter("překročeny mezní rozměry PÚ.")
            bm_mezni_rozmery_PU_paragraph.InsertParagraphAfter()
            bm_mezni_rozmery_PU_paragraph.Collapse(0)
            bm_mezni_rozmery_PU_paragraph_2 = bm_mezni_rozmery_PU_paragraph.Duplicate
            bm_mezni_rozmery_PU_paragraph_2.Style = "Normální"
        if list[0] == "garáž III":
            nadpis_PU = "PÚ " + list[1] + " – " + list[2]
            bm_mezni_rozmery_PU_paragraph.InsertAfter(nadpis_PU)
            bm_mezni_rozmery_PU_paragraph.Style = "Normální"
            bm_mezni_rozmery_PU_paragraph.Font.Bold = True
            bm_mezni_rozmery_PU_paragraph.InsertParagraphAfter()
            bm_mezni_rozmery_PU_paragraph.Collapse(0)
            bm_mezni_rozmery_PU_paragraph.InsertAfter("Mezní rozměry PÚ jsou dány mezním počtem stání pro hromadné garáže. Dle ČSN 73 0804, čl. I.3.4 je stanoven mezní počet stání " + list[8] + ". Posuzovaný prostor hromadné garáže je navržen na max. " + list[9] + " stání.")
            bm_mezni_rozmery_PU_paragraph.Style = "Odstavec se seznamem"
            bm_mezni_rozmery_PU_paragraph.Collapse(0)
            bm_mezni_rozmery_PU_paragraph.InsertAfter("– vyhovuje ")
            bm_mezni_rozmery_PU_paragraph.Font.Bold = True
            bm_mezni_rozmery_PU_paragraph.Collapse(0)
            bm_mezni_rozmery_PU_paragraph.InsertParagraphAfter()
            bm_mezni_rozmery_PU_paragraph.InsertAfter("Dle ČSN 73 0804, čl. I.3.4 ")
            bm_mezni_rozmery_PU_paragraph.Font.Bold = False
            bm_mezni_rozmery_PU_paragraph.Collapse(0)
            bm_mezni_rozmery_PU_paragraph.InsertAfter("nejsou ")
            bm_mezni_rozmery_PU_paragraph.Font.Bold = True
            bm_mezni_rozmery_PU_paragraph.Collapse(0)
            bm_mezni_rozmery_PU_paragraph.InsertAfter("překročeny mezní rozměry PÚ.")
            bm_mezni_rozmery_PU_paragraph.InsertParagraphAfter()
            bm_mezni_rozmery_PU_paragraph.Collapse(0)
            bm_mezni_rozmery_PU_paragraph_2 = bm_mezni_rozmery_PU_paragraph.Duplicate
            bm_mezni_rozmery_PU_paragraph_2.Style = "Normální"
    print("_________________________________________________mezni rozmery paragraphs uploaded")

def samostatne_PU():
    bm_samostatne_PU_paragraph = doc.Bookmarks("SAMOSTATNE_PU").Range
    if objekt_pro_bydleni == "ANO" and pocet_OB <=3:
        bm_samostatne_PU_paragraph.InsertAfter("Objekt bude rozdělen na požární úseky (dále jen „PÚ“) v souladu s požadavky ČSN 73 0802, čl. 5.3.2:")
        bm_samostatne_PU_paragraph.Style = "Normální"
        bm_samostatne_PU_paragraph.InsertParagraphAfter()
        bm_samostatne_PU_paragraph.Collapse(0)
        if typ_garaze == "jednotlivá" and garaz_soucasti_RD == "ANO" and pristresek == "ANO":
            bm_samostatne_PU_paragraph.InsertAfter("dle ČSN 73 0833, čl. 3.9, resp. poznámky I.3.1 může RD tvořit jeden společný PÚ s přístřeškem pro auta (v RD se nachází max. 3 obytné buňky)")
            bm_samostatne_PU_paragraph.Style = "Odstavec se seznamem"
            bm_samostatne_PU_paragraph.Collapse(0)
        elif typ_garaze == "jednotlivá"  and garaz_soucasti_RD == "ANO":
            bm_samostatne_PU_paragraph.InsertAfter("dle ČSN 73 0833, čl. 3.9 může RD tvořit jeden společný PÚ s garáží (v RD se nachází max. 3 obytné buňky)")
            bm_samostatne_PU_paragraph.Style = "Odstavec se seznamem"
            bm_samostatne_PU_paragraph.Collapse(0)
        else:
            bm_samostatne_PU_paragraph.InsertAfter("dle ČSN 73 0833, čl. 3.6 a2) může RD tvořit jeden samostatný PÚ (v RD se nachází max. 3 obytné buňky)")
            bm_samostatne_PU_paragraph.Style = "Odstavec se seznamem"
            bm_samostatne_PU_paragraph.Collapse(0)
    elif objekt_pro_bydleni == "ANO" and pocet_OB > 3:
        bm_samostatne_PU_paragraph.InsertAfter("Objekt bude rozdělen na požární úseky (dále jen „PÚ“) v souladu s požadavky ČSN 73 0802, čl. 5.3.2:")
        bm_samostatne_PU_paragraph.Style = "Normální"
        bm_samostatne_PU_paragraph.InsertParagraphAfter()
        bm_samostatne_PU_paragraph.Collapse(0)
        bm_samostatne_PU_paragraph.InsertAfter("dle ČSN 73 0833, čl. 3.6 a1) musí v budovách skupiny OB2 tvořit každá obytná buňka samostatný PÚ")
        bm_samostatne_PU_paragraph.Style = "Odstavec se seznamem"
        bm_samostatne_PU_paragraph.Collapse(0)
    else:
        bm_samostatne_PU_paragraph.InsertAfter("V objektu se nenachází prostory, které musí dle ČSN 73 0802, čl. 5.3.2 tvořit samostatný požární úsek (dále jen „PÚ“). Objekt bude rozdělen do PÚ s ohledem na mezní parametry dle ČSN 73 0802, čl. 7.3.")

    print("_________________________________________________samostatne PU paragraphs uploaded")

def pozarni_rizika_PU():
    bm_PU_properties = doc.Bookmarks("PU_PROPERTIES").Range
    selected_OB2_elements = [sublist[1] for sublist in d_PU_types if sublist[0] == "OB2"]
    if len(selected_OB2_elements) > 0:
        nadpis_OB2_PU = "PÚ " + ', '.join(selected_OB2_elements) + " – byty"
        bm_PU_properties.InsertAfter(nadpis_OB2_PU)
        bm_PU_properties.Style = "Normální"
        bm_PU_properties.Font.Bold = True
        bm_PU_properties.InsertParagraphAfter()
        bm_PU_properties.Collapse(0)
        bm_PU_properties.InsertAfter("dle ČSN 73 0833, čl. 5.1.2 je požární zatížení PÚ ")
        bm_PU_properties.Style = "Odstavec se seznamem"
        bm_PU_properties.Collapse(0)
        bm_PU_properties.InsertAfter("p")
        bm_PU_properties.Font.Bold = True
        bm_PU_properties.Collapse(0)
        bm_PU_properties.InsertAfter("v")
        bm_PU_properties.Font.Bold = True
        bm_PU_properties.Font.Subscript = True
        bm_PU_properties.Collapse(0)
        bm_PU_properties.InsertAfter(" = 45,00 kg.m")
        bm_PU_properties.Font.Bold = True
        bm_PU_properties.Font.Subscript = False
        bm_PU_properties.Collapse(0)
        bm_PU_properties.InsertAfter("-2")
        bm_PU_properties.Font.Superscript = True
        bm_PU_properties.Style = "Odstavec se seznamem"
        bm_PU_properties.InsertParagraphAfter()
        bm_PU_properties.Collapse(0)
        bm_PU_properties.InsertAfter("dle ČSN 73 0802, přílohy B.1.4 je hodnota součinitele rychlosti odhořívání ")
        bm_PU_properties.Collapse(0)
        bm_PU_properties.InsertAfter("a = 1,00")
        bm_PU_properties.Font.Bold = True
        bm_PU_properties.Collapse(0)
        bm_PU_properties.InsertParagraphAfter()
        bm_PU_properties.InsertAfter("dle ČSN 73 0802, tab. 1 je stálé požární zatížení ")
        bm_PU_properties.Collapse(0)
        bm_PU_properties.InsertAfter("p")
        bm_PU_properties.Font.Bold = True
        bm_PU_properties.Collapse(0)
        bm_PU_properties.InsertAfter("s")
        bm_PU_properties.Font.Bold = True
        bm_PU_properties.Font.Subscript = True
        bm_PU_properties.Collapse(0)
        bm_PU_properties.InsertAfter(" = 10,00 kg.m")
        bm_PU_properties.Font.Bold = True
        bm_PU_properties.Font.Subscript = False
        bm_PU_properties.Collapse(0)
        bm_PU_properties.InsertAfter("-2")
        bm_PU_properties.Font.Superscript = True
        bm_PU_properties.InsertParagraphAfter()
        bm_PU_properties.Collapse(0)
        bm_PU_properties_2 = bm_PU_properties.Duplicate
        bm_PU_properties_2.Style = "Normální"
        bm_PU_properties.InsertParagraphAfter()
    for list in d_PU_types:
        if list[0] == "OB1":
            nadpis_PU = "PÚ " + list[1] + " – " + list[2]
            bm_PU_properties.InsertAfter(nadpis_PU)
            bm_PU_properties.Style = "Normální"
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.InsertParagraphAfter()
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("dle ČSN 73 0802, tab. B.1, pol. 10 je požární zatížení PÚ ")
            bm_PU_properties.Style = "Odstavec se seznamem"
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("p")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("v")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Font.Subscript = True
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter(" = 40 kg.m")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Font.Subscript = False
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("-2")
            bm_PU_properties.Font.Superscript = True
            bm_PU_properties.Style = "Odstavec se seznamem"
            bm_PU_properties.InsertParagraphAfter()
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("dle ČSN 73 0802, přílohy B.1.4 je hodnota součinitele rychlosti odhořívání ")
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("a = 1,00 ")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("(tab. A.1, pol. 8.1)")
            bm_PU_properties.Font.Bold = False
            bm_PU_properties.InsertParagraphAfter()
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("dle ČSN 73 0802, tab. 1 je stálé požární zatížení ")
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("p")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("s")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Font.Subscript = True
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter(" = " + str(list[5]) +" kg.m")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Font.Subscript = False
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("-2")
            bm_PU_properties.Font.Superscript = True
            bm_PU_properties.InsertParagraphAfter()
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("dle ČSN 73 0802, čl. B.1.2 je stanoveno zvýšené výpočtové požární zatížení ")
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("p")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("v")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Font.Subscript = True
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("´ = " + str(list[3]) +" kg.m")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Font.Subscript = False
            bm_PU_properties.Collapse(0)
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.InsertAfter("-2")
            bm_PU_properties.Font.Superscript = True
            bm_PU_properties.InsertParagraphAfter()
            bm_PU_properties.Collapse(0)
            bm_PU_properties_2 = bm_PU_properties.Duplicate
            bm_PU_properties_2.Style = "Normální"
            if list == d_PU_types[-1]:
                pass
            else:
                bm_PU_properties.InsertParagraphAfter()
            print("_________________________________________________OB1 PU paragraphs uploaded")
        elif list[0] == "nevýrobní":
            nadpis_PU = "PÚ " + list[1] + " – " + list[2]
            bm_PU_properties.InsertAfter(nadpis_PU)
            bm_PU_properties.Style = "Normální"
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.InsertParagraphAfter()
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("dle ČSN 73 0802, přílohy A.1 je nahodilé požární zatížení PÚ ")
            bm_PU_properties.Style = "Odstavec se seznamem"
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("p")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("n")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Font.Subscript = True
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter(" = " + str(list[6]) + " kg.m")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Font.Subscript = False
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("-2")
            bm_PU_properties.Font.Superscript = True
            bm_PU_properties.Style = "Odstavec se seznamem"
            bm_PU_properties.InsertParagraphAfter()
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("dle ČSN 73 0802, přílohy B.1.4 je hodnota součinitele rychlosti odhořívání ")
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("a = " + str(list[7]))
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.InsertParagraphAfter()
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("dle ČSN 73 0802, tab. 1 je stálé požární zatížení ")
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("p")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("s")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Font.Subscript = True
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter(" = " + str(list[5]) +" kg.m")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Font.Subscript = False
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("-2")
            bm_PU_properties.Font.Superscript = True
            bm_PU_properties.InsertParagraphAfter()
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("dle ČSN 73 0802, čl. 6.2.1, rovnice (1) je stanoveno ")
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("p")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("v")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Font.Subscript = True
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter(" = " + str(list[3]) +" kg.m")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Font.Subscript = False
            bm_PU_properties.Collapse(0)
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.InsertAfter("-2")
            bm_PU_properties.Font.Superscript = True
            bm_PU_properties.InsertParagraphAfter()
            bm_PU_properties.Collapse(0)
            bm_PU_properties_2 = bm_PU_properties.Duplicate
            bm_PU_properties_2.Style = "Normální"
            if list == d_PU_types[-1]:
                pass
            else:
                bm_PU_properties.InsertParagraphAfter()
            print("_________________________________________________nevyrobni PU paragraphs uploaded")
        elif list[0] == "nevýrobní (pv)" or list[0] == "garáž I" or list[0] == "garáž III":
            nadpis_PU = "PÚ " + list[1] + " – " + list[2]
            bm_PU_properties.InsertAfter(nadpis_PU)
            bm_PU_properties.Style = "Normální"
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.InsertParagraphAfter()
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("dle ČSN 73 0802, tab. B.1 je požární zatížení PÚ ")
            bm_PU_properties.Style = "Odstavec se seznamem"
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("p")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("v")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Font.Subscript = True
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter(" = " + str(list[6]) +" kg.m")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Font.Subscript = False
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("-2")
            bm_PU_properties.Font.Superscript = True
            bm_PU_properties.Style = "Odstavec se seznamem"
            bm_PU_properties.InsertParagraphAfter()
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("dle ČSN 73 0802, přílohy B.1.4 je hodnota součinitele rychlosti odhořívání ")
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("a = " + str(list[7]))
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertParagraphAfter()
            bm_PU_properties.InsertAfter("dle ČSN 73 0802, tab. 1 je stálé požární zatížení ")
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("p")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("s")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Font.Subscript = True
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter(" = " + str(list[5]) +" kg.m")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Font.Subscript = False
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("-2")
            bm_PU_properties.Font.Superscript = True
            bm_PU_properties.InsertParagraphAfter()
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("dle ČSN 73 0802, čl. B.1.2 je stanoveno zvýšené výpočtové požární zatížení ")
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("p")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("v")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Font.Subscript = True
            bm_PU_properties.Collapse(0)
            bm_PU_properties.InsertAfter("´ = " + str(list[3]) +" kg.m")
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.Font.Subscript = False
            bm_PU_properties.Collapse(0)
            bm_PU_properties.Font.Bold = True
            bm_PU_properties.InsertAfter("-2")
            bm_PU_properties.Font.Superscript = True
            bm_PU_properties.InsertParagraphAfter()
            bm_PU_properties.Collapse(0)
            bm_PU_properties_2 = bm_PU_properties.Duplicate
            bm_PU_properties_2.Style = "Normální"
            if list == d_PU_types[-1]:
                pass
            else:
                bm_PU_properties.InsertParagraphAfter()
            print("_________________________________________________OB1 PU paragraphs uploaded")
    if instalace_FVE == "ANO":
        bm_PU_properties.InsertAfter("Dle ČSN 73 0847, čl. 4.2.1 se pro PV systémy nestanovuje požární zatížení.")
        bm_PU_properties.InsertParagraphAfter()
