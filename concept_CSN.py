from placeholders import *



def concept_CSN_ZS_I():
    bm_vychozi_normy = doc.Bookmarks("Vychozi_normy").Range
    bm_vychozi_normy.InsertParagraphAfter()
    bm_vychozi_normy.Collapse(0)
    bm_vychozi_normy.InsertAfter("Posouzení dle ČSN 73 0834")
    bm_vychozi_normy.Style = "Podkapitola 1"
    bm_vychozi_normy.InsertParagraphAfter()

    bm_vychozi_normy.Collapse(0)
    bm_vychozi_normy.InsertAfter("Veškeré změny pro řešený objekt jsou provedeny v mezním rozsahu stanoveném dle položek ČSN 73 0834, čl. 3.2. Žádná z předmětných změn není změnou užívání objektu, prostoru nebo provozu.")
    bm_vychozi_normy.Style = "Normální"
    bm_vychozi_normy.Collapse(0)
    bm_vychozi_normy.InsertAfter("Pozn.: Posouzení změn užívání objektu viz příloha [B] PBŘ.")
    bm_vychozi_normy.Font.Italic = True
    bm_vychozi_normy.Collapse(0)
    bm_vychozi_normy.InsertAfter("Dle ČSN 73 0834, čl. 3.3 je řešený objekt posuzován jako ")
    bm_vychozi_normy.Style = "Normální"
    bm_vychozi_normy.Font.Italic = False
    bm_vychozi_normy.Collapse(0)
    bm_vychozi_normy.InsertAfter("změna stavby skupiny I")
    bm_vychozi_normy.Font.Bold = True
    bm_vychozi_normy.Collapse(0)
    bm_vychozi_normy.InsertAfter(". U objektu je posouzeno pouze dodržení požadavků dle ČSN 73 0834, kap. 4. V objektu nejsou nově zřizovány prostory dle ČSN 73 0834, čl. 3.3 b).")
    bm_vychozi_normy.Font.Bold = False
    bm_vychozi_normy.Collapse(0)
    bm_vychozi_normy.InsertParagraphAfter()
    bm_vychozi_normy.InsertParagraphAfter()
    bm_vychozi_normy.Collapse(0)

    bm_vychozi_normy.InsertAfter("Navrhované změny v objektu:")
    bm_vychozi_normy.HighlightColorIndex = 0
    bm_vychozi_normy.Style = "Normální"
    bm_vychozi_normy.Font.Bold = True
    bm_vychozi_normy.InsertParagraphAfter()
    bm_vychozi_normy.Collapse(0)

    for zmena in l_popis_stavebnich_uprav:
        bm_vychozi_normy.InsertAfter(zmena)
        bm_vychozi_normy.InsertAfter(" (pol. 3.3 ")
        bm_vychozi_normy.InsertAfter(l_clanky_stavebnich_uprav_I[l_popis_stavebnich_uprav.index(zmena)])
        bm_vychozi_normy.InsertParagraphAfter()
    bm_vychozi_normy.Style = "Odstavec se seznamem"
    bm_vychozi_normy.Font.Bold = False
    bm_vychozi_normy.Collapse(0)






def concept_CSN_OB1():
    print("_________________________________________________importing CSN OB info")
    wdCollapseEnd = 0
    bm_vychozi_normy = doc.Bookmarks("Vychozi_normy").Range
    bm_vychozi_normy.InsertParagraphAfter()
    bm_vychozi_normy.Collapse(Direction=wdCollapseEnd)
    bm_vychozi_normy.InsertAfter("Posouzení dle ČSN 73 0833")
    bm_vychozi_normy.Style = "Podkapitola 1"
    bm_vychozi_normy.InsertParagraphAfter()
    bm_vychozi_normy.Collapse(Direction=wdCollapseEnd)
    bm_vychozi_normy.InsertAfter("Objekt RD je navržen jako rodinné bydlení ")
    bm_vychozi_normy.Style = "Normální"
    if int(pocet_OB) == 1:
        bm_vychozi_normy.InsertAfter("s ")
        bm_vychozi_normy.Collapse(Direction=wdCollapseEnd)
        bm_vychozi_normy.InsertAfter("jednou obytnou buňkou")
        bm_vychozi_normy.Font.Bold = True
        bm_vychozi_normy.Collapse(Direction=wdCollapseEnd)
    elif int(pocet_OB) > 1:
        bm_vychozi_normy.InsertAfter("se ")
        bm_vychozi_normy.Collapse(Direction=wdCollapseEnd)
        if int(pocet_OB) == 2:
            bm_vychozi_normy.InsertAfter("dvěma obytnými buňkami")
            bm_vychozi_normy.Font.Bold = True
            bm_vychozi_normy.Collapse(Direction=wdCollapseEnd)
        else:
            bm_vychozi_normy.InsertAfter("třemi obytnými buňkami")
            bm_vychozi_normy.Font.Bold = True
            bm_vychozi_normy.Collapse(Direction=wdCollapseEnd)
    bm_vychozi_normy.InsertAfter(". Tento objekt tak bude dle ČSN 73 0833, čl. 3.5. posuzován jako ")
    bm_vychozi_normy.Font.Bold = False
    bm_vychozi_normy.Collapse(Direction=wdCollapseEnd)
    bm_vychozi_normy.InsertAfter("budova skupiny OB1.")
    bm_vychozi_normy.Font.Bold = True
    bm_vychozi_normy.Collapse(Direction=wdCollapseEnd)
    bm_vychozi_normy.InsertParagraphAfter()
    print("_________________________________________________CSN OB info imported")

def concept_CSN_OB2():
    wdCollapseEnd = 0
    bm_vychozi_normy = doc.Bookmarks("Vychozi_normy").Range
    bm_vychozi_normy.InsertParagraphAfter()
    bm_vychozi_normy.Collapse(Direction=wdCollapseEnd)
    bm_vychozi_normy.InsertAfter("Posouzení dle ČSN 73 0833")
    bm_vychozi_normy.Style = "Podkapitola 1"
    bm_vychozi_normy.InsertParagraphAfter()
    bm_vychozi_normy.Collapse(Direction=wdCollapseEnd)
    if objekt_pro_bydleni == "ANO" and pocet_OB > 3:
        print("____________________________________________importing OB2")
        bm_vychozi_normy.InsertAfter("Objekt BD je navržen jako objekt pro bydlení, ve kterém se nachází celkem ")
        bm_vychozi_normy.Style = "Normální"
        if int(pocet_OB) == 4:
            bm_vychozi_normy.Collapse(Direction=wdCollapseEnd)
            bm_vychozi_normy.InsertAfter(str(pocet_OB))
            bm_vychozi_normy.InsertAfter("obytné buňky.")
            bm_vychozi_normy.Font.Bold = True
            bm_vychozi_normy.Collapse(Direction=wdCollapseEnd)
        else:
            bm_vychozi_normy.Collapse(Direction=wdCollapseEnd)
            bm_vychozi_normy.InsertAfter(str(pocet_OB))
            bm_vychozi_normy.InsertAfter("obytných buněk")
            bm_vychozi_normy.Font.Bold = True
            bm_vychozi_normy.Collapse(Direction=wdCollapseEnd)
        bm_vychozi_normy.InsertAfter(". Tento objekt tak bude dle ČSN 73 0833, čl. 3.5. b) posuzován jako ")
        bm_vychozi_normy.Font.Bold = False
        bm_vychozi_normy.Collapse(Direction=wdCollapseEnd)
        bm_vychozi_normy.InsertAfter("budova skupiny OB2.")
        bm_vychozi_normy.Collapse(Direction=wdCollapseEnd)
        bm_vychozi_normy.InsertParagraphAfter()
    print("_________________________________________________CSN OB info imported")


def concept_garaz():
    bm_vychozi_normy = doc.Bookmarks("Vychozi_normy").Range
    bm_vychozi_normy.InsertParagraphAfter()
    bm_vychozi_normy.Collapse(0)
    bm_vychozi_normy.InsertAfter("Posouzení dle ČSN 73 0804, přílohy I")
    bm_vychozi_normy.Style = "Podkapitola 1"
    bm_vychozi_normy.Collapse(0)
    bm_vychozi_normy.InsertParagraphAfter()
    bm_vychozi_normy.Collapse(0)
    bm_vychozi_normy.Style = "Normální"
    if konstrukce_garaze == "DP1" and pristresek == "ANO" and vice_nez_50_procent_sten == "BEZ STĚN":
        bm_vychozi_normy.InsertAfter("Dle ČSN 73 0804, I.3.8 se nekryté prostory a přístřešky z konstrukcí DP1 bez svislých konstrukcí pro parkování vozidel nepovažují za garáže.")
        bm_vychozi_normy.InsertParagraphAfter()
        bm_vychozi_normy.Style = "Normální"

    elif pristresek == "ANO" and vice_nez_50_procent_sten == "NE" and pocet_stani <= 3:
        bm_vychozi_normy.InsertAfter("Dle ČSN 73 0804, I.3.1 se přístřešky pro max. 3 automobily skupiny 1 za garáže nepovažují. Stěnové konstrukce však mohou být nejvýše na polovině jejich obvodů. Objekt garáže bude posouzen pouze z hlediska odstupových vzdáleností.")
        bm_vychozi_normy.InsertParagraphAfter()
        bm_vychozi_normy.Style = "Normální"

    else:
        if garaz_soucasti_RD == "ANO":
            bm_vychozi_normy.InsertAfter("Garáž v objektu")
            bm_vychozi_normy.Font.Bold = True
            bm_vychozi_normy.Collapse(0)

        else:
            bm_vychozi_normy.InsertAfter("Samostatná garáž")
            bm_vychozi_normy.Font.Bold = True
            bm_vychozi_normy.Collapse(0)

        bm_vychozi_normy.InsertParagraphAfter()
        bm_vychozi_normy.Collapse(0)

        if druh_paliv == "kapalná/elektrická":
            bm_vychozi_normy.InsertAfter("dle ČSN 73 0804, čl. I.2.3.1 a) budou parkovány max. 3 vozidla na kapalná paliva nebo elektrická pohon")
            bm_vychozi_normy.Style = "Odstavec se seznamem"

        elif druh_paliv == "plynná" or druh_paliv == "kombinace":
            bm_vychozi_normy.InsertAfter("dle ČSN 73 0804, čl. I.2.3.1 b) budou parkovány max. 3 vozidla na kapalná nebo plynná paliva nebo na elektrická pohon")
            bm_vychozi_normy.Collapse(0)
        bm_vychozi_normy.InsertParagraphAfter()
        bm_vychozi_normy.InsertAfter("dle ČSN 73 0804, čl. I.2.2 a) se jedná o ")
        bm_vychozi_normy.Style = "Odstavec se seznamem"
        bm_vychozi_normy.Collapse(0)

        bm_vychozi_normy.InsertAfter("garáž skupiny ")
        bm_vychozi_normy.Font.Bold = True
        bm_vychozi_normy.InsertAfter(str(skupina_garaze))
        bm_vychozi_normy.Collapse(0)

        bm_vychozi_normy.InsertParagraphAfter()
        bm_vychozi_normy.InsertAfter("dle ČSN 73 0804, čl. I.2.3 b) je garáž posuzována jako ")
        bm_vychozi_normy.Collapse(0)
        bm_vychozi_normy.InsertAfter("jednotlivá")
        bm_vychozi_normy.Font.Bold = True
        bm_vychozi_normy.Collapse(0)

        bm_vychozi_normy.InsertParagraphAfter()
        bm_vychozi_normy.Collapse(0)

        bm_vychozi_normy.InsertParagraphAfter()
        bm_vychozi_normy.InsertAfter("Dle ČSN 73 0804, čl. I.3.13 mohou být V PÚ jednotlivých a řadových garáží ukládány kapalné pohonné hmoty (nafta, benzin) v nerozbitných přenosných obalech v množství nejvýše 40 litrů a nejvýše 20 l olejů na jedno stání vozidel skupiny 1.")
        bm_vychozi_normy.Style = "Normální"
        bm_vychozi_normy.InsertParagraphAfter()

def concept_FVE():
    bm_vychozi_normy = doc.Bookmarks("Vychozi_normy").Range
    bm_vychozi_normy.InsertParagraphAfter()
    bm_vychozi_normy.Collapse(0)
    bm_vychozi_normy.InsertAfter("Posouzení dle ČSN 73 0847")
    bm_vychozi_normy.Style = "Podkapitola 1"
    bm_vychozi_normy.Collapse(0)
    bm_vychozi_normy.InsertParagraphAfter()
    bm_vychozi_normy.Collapse(0)
    bm_vychozi_normy.InsertAfter("K řešenému objektu je navržena střešní ")
    bm_vychozi_normy.Style = "Normální"
    bm_vychozi_normy.Collapse(0)
    bm_vychozi_normy.InsertAfter("instalace fotovoltaického systému ")
    bm_vychozi_normy.Font.Bold = True
    bm_vychozi_normy.Collapse(0)
    bm_vychozi_normy.InsertAfter("(dále jen „PV systému“), která bude posuzována dle normy ČSN 73 0847.")
    bm_vychozi_normy.Font.Bold = False
    bm_vychozi_normy.InsertParagraphAfter()
    if celkovy_vykon_FVE <= 10:
        bm_vychozi_normy.InsertParagraphAfter()
        bm_vychozi_normy.InsertAfter("Dle ČSN 73 0847, čl. 3.7 a) je navržený PV systém posuzován jako ")
        bm_vychozi_normy.Collapse(0)
        bm_vychozi_normy.InsertAfter("instalace malého rozsahu")
        bm_vychozi_normy.Font.Bold = True
        bm_vychozi_normy.Collapse(0)
        bm_vychozi_normy.InsertAfter(". Instalace nepřevyšuje výkon 10,00 kWp.")
        bm_vychozi_normy.Font.Bold = False
        if FVE_baterie == "NE":
            bm_vychozi_normy.InsertAfter(" Bateriové uložiště není navrženo.")
        bm_vychozi_normy.InsertParagraphAfter()





