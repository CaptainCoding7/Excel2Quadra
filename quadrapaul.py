

#UNUSED
# Import openpyxl 
import openpyxl
import tabula
import fuckit

import pandas as pd
import datetime as dt
from enum import Enum
from kivy.app import App
from kivy.uix.widget import Widget

from readBankStatement import ReadBankStatement
from utils import QuadraPyUtils as Utils
from writeEntries import WriteEntries


class Mode(Enum):
    WRITE_DEBIT = 1
    WRITE_CREDIT = 2
    WRITE_ALL_BY_DATES = 3
    WRITE_ALL_DEB_CRED = 4

def quadrapaul():

    enableExtraLibelle = True

    mode = Mode.WRITE_ALL_BY_DATES

    path = "../releves_excel/Releve_bah.xlsx"
    #path = "../releves_excel/Relevé_cameleon.xlsx"
    #path = "../releves_excel/releve_vintagebar.xlsx"
    #path =  "../releves_excel/Fiducial_RELEVE MENSUEL.xlsx"
    releve = ReadBankStatement(path)
    
    # Si l option ajoutant la deuxieme ligne au libelle est activée
    if enableExtraLibelle == True:
        releve.libellevalues = Utils.addExtraLibelle(releve.datevalues, releve.libellevalues)

    # Instancier un objet d ecriture des donnees dans l excel de sortie
    writer = WriteEntries(releve.datevalues, releve.debitvalues, releve.creditvalues, releve.libellevalues)

    # Editer le fichier excel de sortie selon le mode choisi
    if mode == Mode.WRITE_ALL_BY_DATES:
        writer.writeAllByDate()
    elif mode == Mode.WRITE_ALL_DEB_CRED:
        writer.writeAllDebitsThenCredits()
    elif mode == Mode.WRITE_DEBIT:
        writer.writeDebits()
    elif mode == Mode.WRITE_CREDIT:
        writer.writeCredits()
    
    # Ecrire le dataframe globale dans le fichier excel de sortie
    writer.dfout.to_excel('./outputs/Output.xlsx', startcol=1, startrow=3, sheet_name="data", index=False)


quadrapaul()




'''
Si un debit ou credit n est pas aligne avec la date (mauvaise ligne ou meme mauvaise colonne) 
ce qui arrive dans le cas des cellules de dates fusionnées,
on detecte que la key du dict des dates n a pas d equivalent coté dict credit ou debit
et on met 0 dans credit et debit avec par defaut debit en prime et credit en contrepartie.
On avertit l user avec le nombre de fail et un message invitant à changer le formattage des cellules
dans excel ou a corriger dans Qudrapaie



Liste des courses:
- gui import/export excel
- 4 options impressions debit/credit:
    > only credits
    > only debits
    > debits/credits par date
    > debits puis credits (default)
- Cocher case pour inclure seconde ligne de libelle
- Nice to have:
    + Indiquer format de date manuellement
    + Indiquer colonne manuellement
    + Atribuer un compte au libelle (ex: libelle contient "oxyplast" => 9OXYPLAST)
      en utilisant des couples de valeurs  [ substring  | compte ]

(- Choisir les attributs généraux suivants:      
    _DEFCOMPTE = 47100000
    _CTRPARTIE = 51200000
    _YEAR = '2023'
)
'''