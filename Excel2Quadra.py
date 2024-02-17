

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
from tkinter import Tk     # from tkinter import Tk for Python 3.x
from tkinter.filedialog import askopenfilename
from aenum import Enum

  

class Mode(Enum):

    _init_ = 'value string'

    WRITE_DEBIT = 1, 'Ecrire débits'
    WRITE_CREDIT = 2, 'Ecrire crédits'
    WRITE_ALL_BY_DATES = 3, 'Ecrire débits et crédits (selon date)'
    WRITE_ALL_DEB_CRED = 4, 'Ecrire débits PUIS crédits'

    def __str__(self):
        return self.string


class Excel2Quadra:

    def __init__(self, fileIn, fileOut, enableExtraLibelle, mode):
        self.fileIn = fileIn
        self.fileOut = fileOut
        self.enableExtraLibelle = enableExtraLibelle
        self.mode = mode


    def runApp(self):

        print(self.fileIn)
        print(self.fileOut)
        print(self.enableExtraLibelle)
        print(self.mode)

        #path = "../releves_excel/Relevé_cameleon.xlsx"
        #path = "../releves_excel/releve_vintagebar.xlsx"
        #path =  "../releves_excel/Fiducial_RELEVE MENSUEL.xlsx"
        releve = ReadBankStatement(self.fileIn)
        
        # Si l option ajoutant la deuxieme ligne au libelle est activée
        if self.enableExtraLibelle == True:
            releve.libellevalues = Utils.addExtraLibelle(releve.datevalues, releve.libellevalues)

        # Instancier un objet d ecriture des donnees dans l excel de sortie
        writer = WriteEntries(releve.datevalues, releve.debitvalues, releve.creditvalues, releve.libellevalues)

        # Editer le fichier excel de sortie selon le mode choisi
        if self.mode == Mode.WRITE_ALL_BY_DATES:
            writer.writeAllByDate()
        elif self.mode == Mode.WRITE_ALL_DEB_CRED:
            writer.writeAllDebitsThenCredits()
        elif self.mode == Mode.WRITE_DEBIT:
            writer.writeDebits()
        elif self.mode == Mode.WRITE_CREDIT:
            writer.writeCredits()
        
        # Ecrire le dataframe globale dans le fichier excel de sortie
        writer.dfout.to_excel("outputs/"+self.fileOut, startcol=1, startrow=3, sheet_name="data", index=False)


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