from placeholders import doc

def upload_rekreacni_obj():
    bm_pristup_komunikace = doc.Bookmarks("PRISTUP_KOMUNIKACE_PARAGRAPHS").Range
    wdCollapseEnd = 0
    bm_pristup_komunikace.InsertAfter("Dle ČSN 73 0833, čl. 4.4.2 nemusí být u staveb pro rodinnou rekreaci zřízena přístupová komunikace.")
    bm_pristup_komunikace.Style = "Normální"
    bm_pristup_komunikace.InsertParagraphAfter()
    bm_pristup_komunikace.Collapse(Direction=wdCollapseEnd)