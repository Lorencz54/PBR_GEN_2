from placeholders import *

def autonomni_detekce():
    bm_autonomni_detekce_paragraphs = doc.Bookmarks("AUTONOMNI_DETEKCE_PARAGRAPHS").Range

    bm_autonomni_detekce_paragraphs.InsertAfter("Dle ČSN 73 0833, čl. 4.6 musí být rodinný dům vybaven zařízením autonomní detekce a signalizací. Toto zařízení musí být umístěno v části vedoucí k východu z každé obytné buňky a u rodinných domů s více než jednou obytnou buňkou musí být toto zařízení v nejvyšším místě společné chodby, nebo v jiném prostoru nechráněné únikové cesty. U obytných buněk s podlahovou plochou přes 150 m")
    bm_autonomni_detekce_paragraphs.Collapse(0)

    bm_autonomni_detekce_paragraphs.InsertAfter("2")
    bm_autonomni_detekce_paragraphs.Font.Superscript = True
    bm_autonomni_detekce_paragraphs.Collapse(0)

    bm_autonomni_detekce_paragraphs.InsertAfter(" musí být zařízení v další vhodné části bytu.")
    bm_autonomni_detekce_paragraphs.Font.Superscript = False
    bm_autonomni_detekce_paragraphs.InsertParagraphAfter()
    bm_autonomni_detekce_paragraphs.InsertParagraphAfter()
    bm_autonomni_detekce_paragraphs.InsertAfter("Zařízení budou umístěna:")
    bm_autonomni_detekce_paragraphs.InsertParagraphAfter()
    bm_autonomni_detekce_paragraphs.Collapse(0)

    bm_autonomni_detekce_paragraphs.InsertAfter("v chodbě (m.č. 1.01) bytové jednotky 1ks zařízení ")
    bm_autonomni_detekce_paragraphs.Style = "Odstavec se seznamem"
    bm_autonomni_detekce_paragraphs.HighlightColorIndex = 6
    bm_autonomni_detekce_paragraphs.Collapse(0)
    bm_autonomni_detekce_paragraphs.InsertAfter("– celkem 1ks")
    bm_autonomni_detekce_paragraphs.Font.Bold = True
    bm_autonomni_detekce_paragraphs.InsertParagraphAfter()
    bm_autonomni_detekce_paragraphs.Collapse(0)

    bm_autonomni_detekce_paragraphs.InsertParagraphAfter()
    bm_autonomni_detekce_paragraphs.Style = "Normální"
    bm_autonomni_detekce_paragraphs.Collapse(0)

    bm_autonomni_detekce_paragraphs.InsertAfter("Zařízení autonomní detekce a signalizace musí splňovat požadavky ČSN EN 14 604.")
    bm_autonomni_detekce_paragraphs.Style = "Normální"
    bm_autonomni_detekce_paragraphs.HighlightColorIndex = 0
    bm_autonomni_detekce_paragraphs.Font.Bold = False
    bm_autonomni_detekce_paragraphs.InsertParagraphAfter()







