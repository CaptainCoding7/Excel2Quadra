import pandas as pd
import logging
import sys

'''
logger = logging.getLogger()
logger.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')

stdout_handler = logging.StreamHandler(sys.stdout)
stdout_handler.setLevel(logging.DEBUG)
stdout_handler.setFormatter(formatter)


file_handler = logging.FileHandler('logs.log')
file_handler.setLevel(logging.DEBUG)
file_handler.setFormatter(formatter)

logger.addHandler(file_handler)
logger.addHandler(stdout_handler)
'''

class QuadraPyUtils():

    # Function to find recursively the next key of a dictionary
    def next_key(dict, key):
        keys = iter(dict)
        key in keys
        return next(keys, False)


    def getNumPiece(libelle):

        # Par defaut le numPiece de retour est vide
        retNumPiece = ''

        virtList = {"Virement", "virement", "VIREMENT", "VIRT", "VIR "}
        cbList = {"CB", "CARTE", "Carte"}
        prlvList = {"Prélèvement", "Prelevement", "Prlv", "PRLV", "PRVL"}
        abonList = {"Abonnement", "Abon"}
        rcList = {"Remise", "remise", "chèque", "cheque", "Chèque", "Cheque"}

        allLists = {
            "VIRT" : virtList,
            "CB " : cbList,
            "PRLV" : prlvList, 
            "ABON" : abonList,
            "RC" : rcList
        }

        try:
            retNumPiece = [key for key, list in allLists.items() for label in list if label in libelle].pop()
        except IndexError:
            pass

        return retNumPiece


    def formatNumber(cell):
        # SI nombre delimité par espaces
        if(',' not in cell) and ('.' not in cell) and (' ' in cell):
            # SI 2 espaces
            if cell.count(' ') == 2:
                # delete first space for number > 1 000
                cell = str(cell).replace(" ","",1)
                #Identifier les centimes
                if len(cell.split(' ')[1]) == 2:
                    # replace space by point
                    cell = str(cell).replace(" ",".")
            else:
                # Cas centimes
                if len(cell.split(' ')[1]) == 2:
                    # replace space by point
                    cell = str(cell).replace(" ",".")
                # Cas millier
                else:
                    # delete dot, spaces in number > 1 000
                    cell = str(cell).replace(" ","")
        # Sinon nombre classique
        else:
            # replace comma by point
            cell = str(cell).replace(",",".")
            # delete dot, spaces in number > 1 000
            cell = str(cell).replace(" ","")

        return cell
    
    def addExtraLibelle(datevalues, libellevalues):

        # Pour chaque date
        for daterow in datevalues:      
            # Recuperer le libelle de l operation courante
            libelle = libellevalues[daterow]
            # Recuperer la prochaine key (ligne) des dates
            newtDateRow = QuadraPyUtils.next_key(datevalues, daterow )
            # If the gap between two dates is greater than 1 row
            if(newtDateRow - daterow) > 1:
                try:
                    # Try to get the second libelle cell
                    extralibelle = ' ' + libellevalues[daterow+1]
                except KeyError:
                    # There is no second libelle cell
                    extralibelle = ''
                # Add the extralibelle
                libellevalues[daterow] = libelle +  extralibelle

        return libellevalues
    

    def readListeCompte(dfComptes):

        dictComptes = dict()

        # Pour chaque ligne
        for rowIdx, row in dfComptes.iterrows():
            # Le premier element de la ligne est un num de compte
            compte = row[0]
            if compte is not None :
                patternList = []
                # Parcourir la ligne courante
                for pattern in row[1:]:
                    # Si la cellule est non vide
                    if (pattern is not None) and (pd.isna(pattern)==False):
                        # Ajouter le pattern a la lite associée à la ligne
                        patternList.append(str(pattern))
                # Définir la nouvelle paire clé/valeurS pour ce numéro de compte
                dictComptes[compte] = patternList

        return dictComptes