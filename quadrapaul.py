

# Import openpyxl 
import openpyxl
import pandas as pd
import datetime as dt
import fuckit
from dateutil import parser
from enum import Enum
from kivy.app import App
from kivy.uix.widget import Widget


_DEFCOMPTE = 47100000
_CTRPARTIE = 51200000
_YEAR = '2023'
  
class Mode(Enum):
    WRITE_DEBIT = 1
    WRITE_CREDIT = 2
    WRITE_ALL_BY_DATES = 3
    WRITE_ALL_DEB_CRED = 4

path = "../releves_excel/Releve_bah.xlsx"
path = "../releves_excel/Relevé_cameleon.xlsx"
path = "../releves_excel/releve_vintagebar.xlsx"

date_parser = lambda x: dt.datetime.strptime(x, "%d/%m/%Y")

# Open the spreadsheet 
workbook = openpyxl.load_workbook(path) 
dfIn = pd.read_excel(path, dtype = str)#,parse_dates=['Local Date'], date_parser=date_parser)

# Le format des tuples est le suivant:
# (Date, Compte, Libellé, C.partie, Num.Pièce, Débit, Crédit)
entryLines = {
    'Date': [],
    'Compte': [],
    'Libellé': [],
    'Contrepartie': [],
    'Num.Pièce': [],
    'Débit': [],
    'Crédit': [],
}

dfout = pd.DataFrame(entryLines)

# Get the first sheet 
sheet = workbook._sheets[0] 

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

# Function to determine a column with a label
def findColumnFromLabel(labellist):
    # Iterate over columns
        # WITH OPENPYXL : for column in sheet.iter_cols():
    for (idx,col) in enumerate(dfIn.columns):
        column = dfIn[col]
        # Iterate over column cells
        for cellIdx, cell in column.items():
            # Parmi les labels
            for label in labellist:
                # Si un label est trouvé
                if (label in str(cell)) and cellIdx < 20:
                    # Retourner la colonne
                    return column[cellIdx:]

    return -1


# This function return the dates from the corresponding column
def getDateDict():

    # Create a list to store the values 
    dates = dict() 
    
    # Find the corresponding column
    datecolumn = findColumnFromLabel(["Date", "DATE"])

    # Iterate over the cells in the column 
    for i, cell in enumerate(datecolumn): 
        # If the cell is not empty
        if not pd.isnull(cell):      
            # SI il n y a pas deja une clé i                         
            if i not in dates.keys():    
            # Essayer de formatter la date
                try:
                    # Gerer d abord un cas particulier avec le mois d octobre qui peut etre confondu avec janvier
                    dates[i] = dt.datetime.strptime(str(cell), '%d.1').strftime('%d/10/'+ _YEAR)
                except ValueError or TypeError:
                    try:
                        dates[i] = dt.datetime.strptime(str(cell), '%d.%m').strftime('%d/%m/'+ _YEAR)
                    except ValueError or TypeError:
                        try:
                            dates[i] = dt.datetime.strptime(str(cell), '%d.%m.%y').strftime('%d/%m/'+ _YEAR)
                        except ValueError or TypeError:
                            try:
                                dates[i] = dt.datetime.strptime(str(cell), '%Y-%m-%d %H:%M:%S').strftime('%d/%m/'+ _YEAR)
                                # THIS IS NOT A DATE !
                            except ValueError or TypeError:
                                pass
            else:
                print("Error, entrée de date dupliquée pour la colonne ", i)
    return dates

# This function return the debits from the corresponding column
def getDebitDict():

    # Create a list to store the values 
    debits = dict()

    # Find the corresponding column
    debitcolumn = findColumnFromLabel(["Débit", "DEBIT"])

    # Iterate over the cells in the column 
    for i, cell in enumerate(debitcolumn): 
        # If the cell is not empty
        if not pd.isnull(cell):
            if i not in debits.keys():
                cell = formatNumber(cell)
                try:
                    # Try to convert the string into float
                    float(cell)
                except ValueError:
                    try:
                        # delete first dot (ex: 2.150.15 = 2150.15)
                        cell = str(cell).replace(".","",1)  
                        # Try to convert the string into float
                        float(cell)  
                    except ValueError:
                        # If it's not a float do nothing
                        pass
                    else:
                        #If it's a float add the cell to the list
                        debits[i] = cell
                else:
                    #If it's a float add the cell to the list
                    debits[i] = cell
            else:
                print("Error, entrée de débit dupliquée pour la colonne ", i)

    return debits

# This function return the credits from the corresponding column
def getCreditDict():

    # Create a list to store the values 
    credits = dict()

    # Find the corresponding column
    creditcolumn = findColumnFromLabel(["Crédit", "CREDIT"])

    # Iterate over the cells in the column 
    for i, cell in enumerate(creditcolumn): 
        # If the cell is not empty
        if not pd.isnull(cell):
            if i not in credits.keys():
                cell = formatNumber(cell)
                try:
                    # Try to convert the string into float
                    float(cell)
                except ValueError:
                    try:
                        # delete first dot (ex: 2.150.15 = 2150.15)
                        cell = str(cell).replace(".","",1)  
                        # Try to convert the string into float
                        float(cell)  
                    except ValueError:
                        # If it's not a float do nothing
                        pass
                    else:
                        #If it's a float add the cell to the list
                        credits[i] = cell
                else:
                    #If it's a float add the cell to the list
                    credits[i] = cell
            else:
                print("Error, entrée de débit dupliquée pour la colonne ", i)

    return credits
    

# This function return the libelles from the corresponding column
def getLibelleDict():

    # Create a list to store the values 
    libelles = dict()

    # Find the corresponding column
    libellecolumn = findColumnFromLabel(["Libellé", "Référence", "NATURE"])

    # Iterate over the cells in the column 
    for i, cell in enumerate(libellecolumn): 
        # If the cell is not empty
        if not pd.isnull(cell):
            #If it's a float add the cell to the list
            #debits.append((cell,i)) 
            if i not in libelles.keys():
                libelles[i] = str(cell)[0:30]
            else:
                print("Error, entrée de crédit dupliquée pour la colonne ", i)

    return libelles  


def getNumPiece(libelle):

    # Par defaut le numPiece de retour est vide
    retNumPiece = ''

    virtList = {"Virement", "virement", "VIREMENT", "VIRT"}
    cbList = {"CB", "CARTE", "Carte"}
    prlvList = {"Prélèvement", "Prelevement", "Prlv", "PRLV"}
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

# Function to find recursively the next key of a dictionary
def next_key(dict, key):
    keys = iter(dict)
    key in keys
    return next(keys, False)


def addDebit(date, libelle, numPiece, debit):
    
    newEntry = {
        'Date' : date, 
        'Compte': _DEFCOMPTE, 
        'Libellé': libelle , 
        'Contrepartie':_CTRPARTIE, 
        'Num.Pièce': numPiece, 
        'Débit': debit, 
        'Crédit': ''
    }
    newCtPrtEntry = {
        'Date' : date, 
        'Compte': _CTRPARTIE, 
        'Libellé': libelle, 
        'Contrepartie':_CTRPARTIE, 
        'Num.Pièce': numPiece, 
        'Débit': '', 
        'Crédit': debit
    }

    return (newEntry, newCtPrtEntry)


def addCredit(date, libelle, numPiece, credit):
    newEntry = {
        'Date' : date, 
        'Compte': _DEFCOMPTE, 
        'Libellé': libelle , 
        'Contrepartie':_CTRPARTIE, 
        'Num.Pièce': numPiece, 
        'Débit': '', 
        'Crédit': credit
    }
    newCtPrtEntry = {
        'Date' : date, 
        'Compte': _CTRPARTIE, 
        'Libellé': libelle, 
        'Contrepartie':_CTRPARTIE, 
        'Num.Pièce': numPiece, 
        'Débit': credit, 
        'Crédit': ''
    }

    return (newEntry, newCtPrtEntry)


def writeNewEntry(newEntry, newCtrpEntry):

    global dfout
    # Ajouter la nouvelle entrée et sa contrepartie
    dfout = pd.concat([dfout,pd.DataFrame([newEntry, newCtrpEntry,])], ignore_index=True)
    # Add an empty row for readability
    dfout.loc[len(dfout)] = pd.Series()


def readReleve():


    # Get a dictionnary with all the dates
    datevalues = getDateDict()
    # Get a dictionnary with all the debits
    debitvalues = getDebitDict()
    # Get a dictionnary with all the credits
    creditvalues = getCreditDict()
    # Get a dictionnary with all the libelles
    libellevalues = getLibelleDict()    

    # ENABLE/DISABLE DEBUG PRINT
    if 1:
        print("\n***** DATES **** \n")
        print(datevalues)
        print("\n***** DEBITS **** \n")
        print(debitvalues)
        print("\n***** CREDITS **** \n")
        print(creditvalues)
        print("\n***** LIBELLES **** \n")
        print(libellevalues)

    return (datevalues, debitvalues, creditvalues, libellevalues)



# Ecrire les credits dans le fichier excel de sortie 
def writeDebits(datevalues, debitvalues, libellevalues):

    for debitrow in debitvalues:
        
        # Oui, un debit est sur la meme ligne !
        debit = debitvalues[debitrow]
        try:
            # Recuperer le libelle de l operation courante
            libelle = libellevalues[debitrow]
        except KeyError:
            print("Pas de key pour le libelle associe au debit courant: ",debit)
            pass            
        else:
            # Si il ne s'agit pas d'une ligne décrivant le solde
            if not (("solde" in libelle) or ("Solde" in libelle)):
                try:
                    # Recuperer la date courante
                    date = datevalues[debitrow]
                except KeyError:
                    print("Pas de key pour la date associe au debit courant: ", debit)
                    pass  
                else: 
                    # Get the numPiece
                    numPiece = getNumPiece(libelle)

                    # Creer une entree de crédit
                    (newEntry, newCtrpEntry) = addDebit(date,libelle, numPiece, debit)
                    # Ajouter la nouvelle entrée
                    writeNewEntry(newEntry, newCtrpEntry)


    # Ecrire le dataframe dans le fichier excel de sortie
    dfout.to_excel('Output.xlsx', startcol=1, startrow=3, sheet_name="data", index=False)


# Ecrire les debits dans le fichier excel de sortie 
def writeCredits(datevalues, creditvalues, libellevalues):

    for creditrow in creditvalues:

        # Oui, un credit est sur la meme ligne !
        credit = creditvalues[creditrow]

        try:
            # Recuperer le libelle de l operation courante
            libelle = libellevalues[creditrow]
        except KeyError:
            print("Pas de key pour le libelle associe au credit courant : ", credit)
            pass            
        else:
            # Si il ne s'agit pas d'une ligne décrivant le solde
            if not (("solde" in libelle) or ("Solde" in libelle)):
                try:
                    # Recuperer la date courante
                    date = datevalues[creditrow]
                except KeyError:
                    print("Pas de key pour la date associe au credit courant : ", credit)
                    pass  
                else: 
                    # Get the numPiece
                    numPiece = getNumPiece(libelle)


                    # Creer une entree de crédit
                    (newEntry, newCtrpEntry) = addCredit(date,libelle, numPiece, credit)
                    # Ajouter la nouvelle entrée
                    writeNewEntry(newEntry, newCtrpEntry)


# Ecrire les debits puis les credits dans le fichier excel de sortie 
def writeAllDebitsThenCredits(datevalues, debitvalues, creditvalues, libellevalues):

    writeDebits(datevalues, debitvalues, libellevalues)
    writeCredits(datevalues, creditvalues, libellevalues)



# Ecrire d'après l'ordre des dates les debits et les credits dans le fichier excel de sortie 
def writeAllByDate(datevalues, debitvalues, creditvalues, libellevalues):

    # Pour chaque date
    for daterow in datevalues:

        # Recuper la date courante
        date = datevalues[daterow]

        # Recuperer le libelle de l operation courante
        libelle = libellevalues[daterow]
       
        # Get the numPiece
        numPiece = getNumPiece(libelle)

        # Si il ne s'agit pas d'une ligne décrivant le solde
        if not (("solde" in libelle) or ("Solde" in libelle)):

            # Est ce un debit
            try:
                # Oui, un debit est sur la meme ligne !
                debit = debitvalues[daterow]
                # Creer une entree de débit
                (newEntry, newCtrpEntry) = addDebit(date,libelle, numPiece, debit)
            # Non !
            except KeyError:
                # Est ce un credit
                try:
                    # Oui, un credit est sur la meme ligne !
                    credit = creditvalues[daterow]
                    # Creer une entree de crédit
                    (newEntry, newCtrpEntry) = addCredit(date,libelle, numPiece, credit)
                except KeyError:
                    print("Pas de données de débit ou crédit pour la date du ",datevalues[daterow])
                else:
                    # Ajouter la nouvelle entrée
                    writeNewEntry(newEntry, newCtrpEntry)
            else:   
                # Ajouter la nouvelle entrée
                writeNewEntry(newEntry, newCtrpEntry)
   


def quadrapaul():

    enableExtraLibelle = True

    mode = Mode.WRITE_ALL_DEB_CRED

    (datevalues, debitvalues, creditvalues, libellevalues) = readReleve()

    # Si l option ajoutant la deuxieme ligne au libelle est activée
    if enableExtraLibelle == True:
        # Pour chaque date
        for daterow in datevalues:      
            # Recuperer le libelle de l operation courante
            libelle = libellevalues[daterow]
            # Recuperer la prochaine key (ligne) des dates
            newtDateRow = next_key(datevalues, daterow )
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

    
    # Editer le fichier excel de sortie selon le mode choisi
    if mode == Mode.WRITE_ALL_BY_DATES:
        writeAllByDate(datevalues, debitvalues, creditvalues, libellevalues)
    elif mode == Mode.WRITE_ALL_DEB_CRED:
        writeAllDebitsThenCredits(datevalues, debitvalues, creditvalues, libellevalues)
    elif mode == Mode.WRITE_DEBIT:
        writeDebits(datevalues, debitvalues, libellevalues)
    elif mode == Mode.WRITE_CREDIT:
        writeCredits(datevalues, creditvalues, libellevalues)
    
    # Ecrire le dataframe globale dans le fichier excel de sortie
    dfout.to_excel('./outputs/Output.xlsx', startcol=1, startrow=3, sheet_name="data", index=False)


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