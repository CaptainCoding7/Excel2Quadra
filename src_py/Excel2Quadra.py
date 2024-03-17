
import pandas as pd
import datetime as dt
from enum import Enum
from tkinter.messagebox import showinfo
import os

from readBankStatement import ReadBankStatement
from utils import QuadraPyUtils as Utils
from writeEntries import WriteEntries


class Mode(Enum):

    WRITE_DEBIT = 'Ecrire débits'
    WRITE_CREDIT = 'Ecrire crédits'
    WRITE_ALL_BY_DATES = 'Ecrire débits et crédits (selon date)'
    WRITE_ALL_DEB_CRED = 'Ecrire débits PUIS crédits'

class Excel2Quadra:

    def __init__(self, fileIn, fileOut, enableExtraLibelle, mode, numCompte, ctrpartie):
        self.fileIn = fileIn
        self.fileOut = fileOut
        self.enableExtraLibelle = enableExtraLibelle
        self.mode = mode
        self.defNumCompte = numCompte
        self.contrepartie = ctrpartie


    def runApp(self):

        releve = ReadBankStatement(self.fileIn)
        
        # Si l option ajoutant la deuxieme ligne au libelle est activée
        if self.enableExtraLibelle == True:
            releve.libellevalues = Utils.addExtraLibelle(releve.datevalues, releve.libellevalues)

        # Instancier un objet d ecriture des donnees dans l excel de sortie
        writer = WriteEntries(releve.datevalues, releve.debitvalues, releve.creditvalues, releve.libellevalues, self.defNumCompte, self.contrepartie)

        # Editer le fichier excel de sortie selon le mode choisi
        if self.mode == Mode.WRITE_ALL_BY_DATES:
            writer.writeAllByDate()
            strDescr = "des débits et crédits "
        elif self.mode == Mode.WRITE_ALL_DEB_CRED:
            writer.writeAllDebitsThenCredits()
            strDescr = "des débits et crédits "
        elif self.mode == Mode.WRITE_DEBIT:
            writer.writeDebits()
            strDescr = "des débits "
        elif self.mode == Mode.WRITE_CREDIT:
            writer.writeCredits()
            strDescr = "des crédits "
        else:
            print("Error unknown mode")
            return -1
        
        output_dir = "outputs_releves_transformes"

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        # Ecrire le dataframe globale dans le fichier excel de sortie
        writer.dfout.to_excel(output_dir+"/"+self.fileOut, startcol=1, startrow=3, sheet_name="data", index=False)
        showinfo(
            title = "Fin",
            message = "Fin de l'écriture " + strDescr +"dans " + self.fileOut 
        )

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
- Afficher fichier d'entrée chargé
- Choisir dossier de sortie
- Message popup chargement ok + fails => class pop up appelée par class Excel2Quadra

- Nice to have:

    + Indiquer format de date manuellement
    + Indiquer colonne manuellement
    + Atribuer un compte au libelle (ex: libelle contient "oxyplast" => 9OXYPLAST)
      en utilisant des couples de valeurs  [ substring  | compte ]
    + Choisir année (et mois ?)
    + Ne choisir que les dates du mois considéré (pas les dates du début du second mois)

(- Choisir les attributs généraux suivants:      
    _DEFCOMPTE = 47100000
    _CTRPARTIE = 51200000
    _YEAR = '2023'
)
'''