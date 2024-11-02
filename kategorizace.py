from excel_data import *
from placeholders import doc

def zduvodneni_kategorizace():
    bm_kategorizace_paragraphs = doc.Bookmarks("KATEGORIZACE_PARAGRAPHS").Range
    bm_kategorizace_paragraphs.InsertParagraphAfter()
    bm_kategorizace_paragraphs.InsertAfter("Na základě odborného zhodnocení navrhovaných prací dle PD bylo shledáno, že provedené stavební úpravy odpovídají ustanovení § 6 odst. (2) vyhl. č. 460/2021 Sb.: ")
    bm_kategorizace_paragraphs.InsertParagraphAfter()
    bm_kategorizace_paragraphs.Collapse(0)
    bm_kategorizace_paragraphs.InsertAfter("udržovací práce nebo stavební úpravy, pokud jejich provedení negativně neovlivní požární bezpečnost stavby nebo nezasáhne trvalý ochranný prostor stálého úkrytu. ")
    bm_kategorizace_paragraphs.Font.Bold = True
    bm_kategorizace_paragraphs.Collapse(0)
    bm_kategorizace_paragraphs.InsertAfter("Takovéto udržovací práce nebo stavební úpravy se bez ohledu na vlastní kategorii stavby, ve které se budou realizovat, posoudí z hlediska požadavků na projektovou dokumentaci nebo dokumentaci stavby jako stavba kategorie 0. Ustanovení § 3 odst. 1 vyhlášky se v těchto případech nepoužije. Tzn. předmětné stavební úpravy spadají do kategorie „0“, a tudíž pro ně")
    bm_kategorizace_paragraphs.Font.Bold = False
    bm_kategorizace_paragraphs.Collapse(0)
    bm_kategorizace_paragraphs.InsertAfter(" nemusí být PBŘ vypracováno.")
    bm_kategorizace_paragraphs.Font.Bold = True
    bm_kategorizace_paragraphs.InsertParagraphAfter()
    bm_kategorizace_paragraphs.InsertParagraphAfter()
    bm_kategorizace_paragraphs.InsertAfter("Zdůvodnění/podmínky:")
    bm_kategorizace_paragraphs.Font.Bold = True
    bm_kategorizace_paragraphs.Collapse(0)
    bm_kategorizace_paragraphs.InsertParagraphAfter()
    bm_kategorizace_paragraphs.Collapse(0)
    bm_kategorizace_paragraphs.InsertAfter("nedochází ke zvýšení požárního rizika")
    bm_kategorizace_paragraphs.InsertParagraphAfter()
    bm_kategorizace_paragraphs.InsertAfter("nedochází ke zvětšení plochy požárních úseků nebo vzniku nových požárních úseků")
    bm_kategorizace_paragraphs.InsertParagraphAfter()
    bm_kategorizace_paragraphs.InsertAfter("nejsou zhoršeny podmínky evakuace osob a zásahu jednotek požární ochrany (zvýšení počtu osob, prodloužení délky únikové cesty, zhoršení větrání chráněné únikové cesty nebo zásahové cesty apod.)")
    bm_kategorizace_paragraphs.InsertParagraphAfter()
    bm_kategorizace_paragraphs.InsertAfter("nejsou zhoršeny vlastnosti stavebních konstrukcí či hmoty z hlediska požární bezpečnosti (např. požární odolnost, třída reakce na oheň a index šíření plamene po povrchu)")
    bm_kategorizace_paragraphs.InsertParagraphAfter()
    bm_kategorizace_paragraphs.InsertAfter("nejsou zvětšeny odstupové vzdálenosti (např. provedení nových požárně otevřených ploch v obvodových konstrukcích, provedení fasády z hořlavých stavebních výrobků apod.)")
    bm_kategorizace_paragraphs.Style = "Odstavec se seznamem"
    bm_kategorizace_paragraphs.Font.Bold = False
    bm_kategorizace_paragraphs.Collapse(0)
    bm_kategorizace_paragraphs.InsertParagraphAfter()
    bm_kategorizace_paragraphs.Collapse(0)
    bm_kategorizace_paragraphs.Style = "Normální"
