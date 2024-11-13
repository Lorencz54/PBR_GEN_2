from placeholders import *

def popis_konstrukci_a_trida_reakce_table_insert():
    if l_svisle_nosne or l_svisle_nenosne:
        bm_svisle_konstrukce = doc.Bookmarks("SVISLE_KONSTRUKCE").Range
        bm_svisle_konstrukce.Style = doc.Styles("Normální")
        bm_trida_svisle_nenosne = doc.Bookmarks("TRIDA_SVISLE_NENOSNE").Range
        bm_material_svisle_nenosne = doc.Bookmarks("MATERIAL_SVISLE_NENOSNE").Range
        for konstrukce in l_svisle_nenosne + l_svisle_nosne:
            bm_svisle_konstrukce.InsertAfter(konstrukce.value)
            bm_svisle_konstrukce.InsertParagraphAfter()
        bm_svisle_konstrukce_2 = bm_svisle_konstrukce.Duplicate
        bm_svisle_konstrukce_2.Style = "Odstavec se seznamem"
        for material, trida in zip(l_material_svisle_nenosne, l_tridy_svisle_nenosne):
            bm_material_svisle_nenosne.InsertAfter(material.value)
            bm_trida_svisle_nenosne.InsertAfter(trida.value)
            if material != l_material_svisle_nenosne[-1]:
                bm_material_svisle_nenosne.InsertParagraphAfter()
                bm_trida_svisle_nenosne.InsertParagraphAfter()
    print("_______________________________________ svisle nenosne imported")
    if l_vodorovne_nenosne or l_vodorovne_nosne:
        bm_vodorovne_konstrukce = doc.Bookmarks("VODOROVNE_KONSTRUKCE").Range
        bm_vodorovne_konstrukce.Style = doc.Styles("Normální")
        bm_material_vodorovne_nenosne = doc.Bookmarks("MATERIAL_VODOROVNE_NENOSNE").Range
        bm_trida_vodorovne_nenosne = doc.Bookmarks("TRIDA_VODOROVNE_NENOSNE").Range
        for konstrukce in l_vodorovne_nenosne + l_vodorovne_nosne:
            bm_vodorovne_konstrukce.InsertAfter(konstrukce.value)
            bm_vodorovne_konstrukce.InsertParagraphAfter()
        bm_vodorovne_konstrukce_2 = bm_vodorovne_konstrukce.Duplicate
        bm_vodorovne_konstrukce_2.Style = "Odstavec se seznamem"
        for material, trida in zip(l_material_vodorovne_nenosne, l_tridy_vodorovne_nenosne):
            bm_material_vodorovne_nenosne.InsertAfter(material.value)
            bm_trida_vodorovne_nenosne.InsertAfter(trida.value)
            if material != l_material_vodorovne_nenosne[-1]:
                bm_material_vodorovne_nenosne.InsertParagraphAfter()
                bm_trida_vodorovne_nenosne.InsertParagraphAfter()
    print("_______________________________________ vodorovne nenosne imported")
    if l_svisle_nosne:
        bm_material_svisle_konstrukce = doc.Bookmarks("MATERIAL_SVISLE_NOSNE_KONSTRUKCE").Range
        bm_trida_svisle_konstrukce = doc.Bookmarks("TRIDA_SVISLE_NOSNE").Range
        for material, trida in zip(l_material_svisle_nosne, l_tridy_svisle_nosne):
            bm_material_svisle_konstrukce.InsertAfter(material.value)
            bm_trida_svisle_konstrukce.InsertAfter(trida.value)
            if material != l_material_svisle_nosne[-1]:
                bm_material_svisle_konstrukce.InsertParagraphAfter()
                bm_trida_svisle_konstrukce.InsertParagraphAfter()
    print("_______________________________________ svisle imported")
    if l_vodorovne_nosne:
        bm_material_vodorovne_konstrukce = doc.Bookmarks("MATERIAL_VODOROVNE_NOSNE").Range
        bm_trida_vodorovne_konstrukce = doc.Bookmarks("TRIDA_VODOROVNE_NOSNE").Range
        for material, trida in zip(l_material_vodorovne_nosne, l_tridy_vodorovne_nosne):
            bm_material_vodorovne_konstrukce.InsertAfter(material.value)
            bm_trida_vodorovne_konstrukce.InsertAfter(trida.value)
            if material != l_material_vodorovne_nosne[-1]:
                bm_material_vodorovne_konstrukce.InsertParagraphAfter()
                bm_trida_vodorovne_konstrukce.InsertParagraphAfter()
    print("_______________________________________ vodorovne imported")
    if l_stresni_krytiny:
        bm_stresni_krytiny = doc.Bookmarks("STRECHA").Range
        bm_stresni_krytiny.Style = doc.Styles("Normální")
        bm_material_stresni_krytiny = doc.Bookmarks("MATERIAL_STRESNI_KRYTINA").Range
        bm_trida_stresni_krytiny = doc.Bookmarks("TRIDA_STRESNI_KRYTINA").Range
        for konstrukce in l_stresni_krytiny + l_konstrukce_strechy:
            bm_stresni_krytiny.InsertAfter(konstrukce.value)
            bm_stresni_krytiny.InsertParagraphAfter()
        bm_stresni_krytiny_2 = bm_stresni_krytiny.Duplicate
        bm_stresni_krytiny_2.Style = "Odstavec se seznamem"
        for material, trida in zip(l_material_stresni_krytiny, l_tridy_stresni_krytiny):
            bm_material_stresni_krytiny.InsertAfter(material.value)
            bm_trida_stresni_krytiny.InsertAfter(trida.value)
            if material != l_material_stresni_krytiny[-1]:
                bm_material_stresni_krytiny.InsertParagraphAfter()
                bm_trida_stresni_krytiny.InsertParagraphAfter()
    print("_______________________________________ krytiny imported")
    if l_konstrukce_strechy:
        bm_material_strechy = doc.Bookmarks("MATERIAL_NOSNA_KONSTRUKCE_STRECHY").Range
        bm_trida_strechy = doc.Bookmarks("TRIDA_KONSTRUKCE_STRECHY").Range
        for material, trida in zip(l_material_konstrukce_strechy, l_tridy_konstrukce_strechy):
            bm_material_strechy.InsertAfter(material.value)
            bm_trida_strechy.InsertAfter(trida.value)
            if material != l_material_konstrukce_strechy[-1]:
                bm_material_strechy.InsertParagraphAfter()
                bm_trida_strechy.InsertParagraphAfter()
    print("_______________________________________ strecha imported")
    if l_tepelne_izolace:
        bm_tepelne_izolace = doc.Bookmarks("TEPELNE_IZOLACE").Range
        bm_tepelne_izolace.Style = doc.Styles("Normální")
        bm_material_tepelne_izolace = doc.Bookmarks("MATERIAL_TEPELNE_IZOLACE").Range
        bm_trida_tepelne_izolace = doc.Bookmarks("TRIDA_TEPELNA_IZOLACE").Range
        for konstrukce in l_tepelne_izolace:
            bm_tepelne_izolace.InsertAfter(konstrukce.value)
            bm_tepelne_izolace.InsertParagraphAfter()
        bm_tepelne_izolace_2 = bm_tepelne_izolace.Duplicate
        bm_tepelne_izolace_2.Style = "Odstavec se seznamem"
        for material, trida in zip(l_material_tepelne_izolace, l_tridy_tepelne_izolace):
            bm_material_tepelne_izolace.InsertAfter(material.value)
            bm_trida_tepelne_izolace.InsertAfter(trida.value)
            if material != l_material_tepelne_izolace[-1]:
                bm_material_tepelne_izolace.InsertParagraphAfter()
                bm_trida_tepelne_izolace.InsertParagraphAfter()
    print("_______________________________________ izolace imported")
    if l_schodiste:
        bm_schodiste = doc.Bookmarks("SCHODISTE").Range
        bm_schodiste.Style = doc.Styles("Normální")
        bm_material_schodiste = doc.Bookmarks("MATERIAL_VODOROVNE_NOSNE").Range
        bm_trida_vodorovne_konstrukce = doc.Bookmarks("TRIDA_VODOROVNE_NOSNE").Range
        for konstrukce in l_schodiste:
            bm_schodiste.InsertAfter(konstrukce.value)
            bm_schodiste.InsertParagraphAfter()
        bm_schodiste_2 = bm_schodiste.Duplicate
        bm_schodiste_2.Style = "Odstavec se seznamem"
        for material, trida in zip(l_material_schodiste, l_tridy_schodiste):
            bm_material_schodiste.InsertAfter(material.value)
            bm_trida_vodorovne_konstrukce.InsertAfter(trida.value)
            if material != l_material_schodiste[-1] or len(l_vodorovne_nosne) != 0:
                bm_material_schodiste.InsertParagraphAfter()
                bm_trida_vodorovne_konstrukce.InsertParagraphAfter()
    print("_______________________________________ schodiste imported")
    if l_podlahy:
        bm_podlahy = doc.Bookmarks("PODLAHY").Range
        bm_podlahy.Style = doc.Styles("Normální")
        bm_material_podlahy = doc.Bookmarks("MATERIAL_PODLAHA").Range
        bm_trida_podlahy = doc.Bookmarks("TRIDA_PODLAHY").Range
        for konstrukce in l_podlahy:
            bm_podlahy.InsertAfter(konstrukce.value)
            bm_podlahy.InsertParagraphAfter()
        bm_podlahy_2 = bm_podlahy.Duplicate
        bm_podlahy_2.Style = "Odstavec se seznamem"
        for material, trida in zip(l_material_podlahy, l_tridy_podlahy):
            bm_material_podlahy.InsertAfter(material.value)
            bm_trida_podlahy.InsertAfter(trida.value)
            if material != l_material_podlahy[-1]:
                bm_material_podlahy.InsertParagraphAfter()
                bm_trida_podlahy.InsertParagraphAfter()
    print("_______________________________________ podlahy imported")
    if l_vyplne_otvoru:
        bm_vyplne_otvoru = doc.Bookmarks("VYPLNE_OTVORU").Range
        bm_vyplne_otvoru.Style = doc.Styles("Normální")
        bm_material_vyplne_otvoru = doc.Bookmarks("MATERIAL_VYPLNE_OTVORU").Range
        bm_trida_vyplne_otvoru = doc.Bookmarks("TRIDA_VYPLNE_OTVORU").Range
        for konstrukce in l_vyplne_otvoru:
            bm_vyplne_otvoru.InsertAfter(konstrukce.value)
            bm_vyplne_otvoru.InsertParagraphAfter()
        bm_vyplne_otvoru_2 = bm_vyplne_otvoru.Duplicate
        bm_vyplne_otvoru_2.Style = "Odstavec se seznamem"
        for material, trida in zip(l_material_vyplne_otvoru, l_tridy_vyplne_otvoru):
            bm_material_vyplne_otvoru.InsertAfter(material.value)
            bm_trida_vyplne_otvoru.InsertAfter(trida.value)
            if material != l_material_vyplne_otvoru[-1]:
                bm_material_vyplne_otvoru.InsertParagraphAfter()
                bm_trida_vyplne_otvoru.InsertParagraphAfter()
    print("_______________________________________ vyplne imported")
    if l_vnejsi_povrchy:
        bm_vnejsi_povrchy = doc.Bookmarks("VNEJSI_POVRCHY").Range
        bm_vnejsi_povrchy.Style = doc.Styles("Normální")
        bm_material_vnejsi_povrchy = doc.Bookmarks("MATERIAL_VNEJSI_POVRCHY").Range
        bm_trida_vnejsi_povrchy = doc.Bookmarks("TRIDA_VNEJSI_POVRCHY").Range
        for konstrukce in l_vnejsi_povrchy:
            bm_vnejsi_povrchy.InsertAfter(konstrukce.value)
            bm_vnejsi_povrchy.InsertParagraphAfter()
        bm_vnejsi_povrchy_2 = bm_vnejsi_povrchy.Duplicate
        bm_vnejsi_povrchy_2.Style = "Odstavec se seznamem"
        for material, trida in zip(l_material_vnejsi_povrchy, l_tridy_vnejsi_povrchy):
            bm_material_vnejsi_povrchy.InsertAfter(material.value)
            bm_trida_vnejsi_povrchy.InsertAfter(trida.value)
            if material != l_material_vnejsi_povrchy[-1]:
                bm_material_vnejsi_povrchy.InsertParagraphAfter()
                bm_trida_vnejsi_povrchy.InsertParagraphAfter()
    print("_______________________________________ vnejsi povrchy imported")
    print("_________________________________________________constructions imported")


