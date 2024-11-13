from concept_CSN import *
from popis_konstrukci import *
from SPB_upload import *
from pozarni_riziko import *
from zhodnoceni_konstrukci import *
from evakuace import *
from hasiva import *
from moznosti_zasahu import *
from PBZ import *
from konstrukcni_system import *
from kategorizace import *
from zakladni_popis_projektu import *
from zmeny_staveb import *

if typ_garaze == "jednotlivá":
    concept_garaz()
    evakuace_garaz_I_obecne()

if instalace_FVE == "ANO":
    concept_FVE()

if objekt_pro_bydleni == "ANO":
    if pocet_OB <= 3 and pocet_PP_obj <= 1 and pocet_NP_obj <= 3:
        concept_CSN_OB1()
        autonomni_detekce()
        if rekreacni_objekt == "ANO":
            upload_rekreacni_obj()
    elif pocet_OB > 3:
        concept_CSN_OB2()


    if kategorie == 0:
        zduvodneni_kategorizace()
    else:
        if skupina_zmeny_stavby == "I.":
            upload_popis_zmen_paragraph()
            insert_zmeny_konstrukci_table()
            concept_CSN_ZS_I()
            popis_projektu()
            popis_konstrukci_a_trida_reakce_table_insert()
            vnitrni_odberne_mista()
            PHP()
        else:
            popis_projektu()
            konstrukcni_system()
            popis_konstrukci_a_trida_reakce_table_insert()
            samostatne_PU()
            pozarni_rizika_PU()
            upload_SPB_paragraphs()
            mezni_rozmery_PU()
            upload_konstrukce_paragraphs()
            evakuace_obecne()
            konstrukce_table_insert()
            pozarni_pasy()


doc.Save()
word.Quit()