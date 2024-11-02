from excel_data import *
from placeholders import doc

def konstrukcni_system():
    bm_k_system_paragraphs = doc.Bookmarks("KONSTRUKCNI_SYSTEM").Range
    all_lists_only_DP1 = all(all(item == "DP1" for item in list) for list in d_DP_konstrukce)
    svisle_nosne_only_DP1 = all(item == "DP1" for item in l_DP_svisle_nosne)
    svisle_pozarne_delici_only_DP1 = False
    for boolean, dp in zip(l_pozarne_delici_svisle_nosne, l_DP_svisle_nosne):
        if (boolean == "ANO" and dp == "DP1") or (boolean == "NE"):
            svisle_pozarne_delici_only_DP1 = True
        else:
            svisle_pozarne_delici_only_DP1 = False
            break

    ostatni_pozarne_delici_least_DP2 = False
    ostatni_nosne_least_DP2 = False

    if pocet_NP_obj == 1:
        for list in d_DP_konstrukce[:-1]:
            if "DP3" in list:
                ostatni_nosne_least_DP2 = False
            else:
                ostatni_nosne_least_DP2 = True
        for boolean_list, dp_list in zip(d_pozarne_delici_konstrukce[1:], d_DP_konstrukce[1:]):
            for boolean, dp in zip(boolean_list, dp_list):
                if (boolean == "ANO" and dp == "DP1") or (boolean == "ANO" and dp == "DP2") or (boolean == "NE"):
                    ostatni_pozarne_delici_least_DP2 = True

                else:
                    ostatni_pozarne_delici_least_DP2 = False
                    break
            else:
                continue
            break
    else:
        for list in d_DP_konstrukce:
            if "DP3" in list:
                ostatni_nosne_least_DP2 = False
            else:
                ostatni_nosne_least_DP2 = True
            for boolean_list, dp_list in zip(d_pozarne_delici_konstrukce[1:], d_DP_konstrukce[1:]):
                for boolean, dp in zip(boolean_list, dp_list):
                    if (boolean == "ANO" and dp == "DP1") or (boolean == "ANO" and dp == "DP2") or (boolean == "NE"):
                        ostatni_pozarne_delici_least_DP2 = True

                    else:
                        ostatni_pozarne_delici_least_DP2 = False
                        break
                else:
                    continue
                break


    if all_lists_only_DP1:
        bm_k_system_paragraphs.InsertAfter("Dle ČSN 73 0802, čl. 7.2.8 a) se konstrukční systém posuzovaného objektu posuzuje jako ")
        bm_k_system_paragraphs.Collapse(0)

        bm_k_system_paragraphs.InsertAfter("nehořlavý")
        bm_k_system_paragraphs.Font.Bold = True
        bm_k_system_paragraphs.Collapse(0)

        bm_k_system_paragraphs.InsertAfter(". Všechny požárně dělící a nosné konstrukce zajišťující stabilitu objektu nebo jeho části jsou druhu DP1.")
        bm_k_system_paragraphs.Font.Bold = False
        bm_k_system_paragraphs.Collapse(0)
        bm_k_system_paragraphs.InsertParagraphAfter()

    elif svisle_nosne_only_DP1 and svisle_pozarne_delici_only_DP1 and ostatni_pozarne_delici_least_DP2 and ostatni_nosne_least_DP2:
        print(svisle_nosne_only_DP1, svisle_pozarne_delici_only_DP1, ostatni_pozarne_delici_least_DP2)
        bm_k_system_paragraphs.InsertAfter("Dle ČSN 73 0802, čl. 7.2.8 b) se konstrukční systém posuzovaného objektu posuzuje jako ")
        bm_k_system_paragraphs.Collapse(0)

        bm_k_system_paragraphs.InsertAfter("smíšený")
        bm_k_system_paragraphs.Font.Bold = True
        bm_k_system_paragraphs.Collapse(0)

        bm_k_system_paragraphs.InsertAfter(". Svislé nosné konstrukce a požárně dělící konstrukce jsou druhu DP1. Ostatní požárně dělící a nosné konstrukce jsou druhu min. DP2.")
        if pocet_NP_obj == 1:
            bm_k_system_paragraphs.InsertAfter(" Konstrukce střechy může být druhu DP3.")
        bm_k_system_paragraphs.Font.Bold = False
        bm_k_system_paragraphs.Collapse(0)
        bm_k_system_paragraphs.InsertParagraphAfter()
    else:
        bm_k_system_paragraphs.InsertAfter("Dle ČSN 73 0802, čl. 7.2.8 c) se konstrukční systém posuzovaného objektu posuzuje jako ")
        bm_k_system_paragraphs.Collapse(0)

        bm_k_system_paragraphs.InsertAfter("hořlavý")
        bm_k_system_paragraphs.Font.Bold = True
        bm_k_system_paragraphs.Collapse(0)

        bm_k_system_paragraphs.InsertAfter(".")
        bm_k_system_paragraphs.Font.Bold = False
        bm_k_system_paragraphs.Collapse(0)
        bm_k_system_paragraphs.InsertParagraphAfter()