import pandas as pd
import datetime as dt
from utils import QuadraPyUtils as Utils
from tkinter.messagebox import showinfo


class ReadBankStatement:

    def __init__(self,df):
        self.dfIn = df

        self._YEAR = '2023'
        self.colIdxToExclude = -1

        # Get a dictionnary with all the dates
        self.datevalues = self.getDateDict()
        # Get a dictionnary with all the debits
        self.debitvalues = self.getDebitDict()
        # Get a dictionnary with all the credits
        self.creditvalues = self.getCreditDict()
        # Get a dictionnary with all the libelles
        self.libellevalues = self.getLibelleDict()    

        # ENABLE/DISABLE DEBUG PRINT
        if 0:
            print("\n***** DATES **** \n")
            print(self.datevalues)
            print("\n***** DEBITS **** \n")
            print(self.debitvalues)
            print("\n***** CREDITS **** \n")
            print(self.creditvalues)
            print("\n***** LIBELLES **** \n")
            print(self.libellevalues)

    # Function to determine a column with a label
    def findColumnFromLabel(self, labellist):
        # Iterate over columns
            # WITH OPENPYXL : for column in sheet.iter_cols():
        print("********")
        for (idx,col) in enumerate(self.dfIn.columns):
            # If the current column is not the one to exclude
            if idx != self.colIdxToExclude:
                column = self.dfIn[col]
                # Iterate over column cells
                for cellIdx, cell in column.items():
                    # Parmi les labels
                    for label in labellist:
                        # Si un label est trouvé
                        if (label in str(cell)) and cellIdx < 50:
                            # Retourner la colonne
                            return (column[cellIdx:],idx)
        print("ERROR : No column returned")
        return -1


    # This function return the dates from the corresponding column
    def getDateDict(self):

        # Create a list to store the values 
        dates = dict() 
        
        # Find the corresponding column
        (datecolumn, dateColIdx) = self.findColumnFromLabel(["Date", "DATE"])

        # Indicate that the date column must not be find anymore 
        self.colIdxToExclude = dateColIdx

        # Iterate over the cells in the column 
        for i, cell in enumerate(datecolumn): 
            # If the cell is not empty
            if not pd.isnull(cell):      
                # SI il n y a pas deja une clé i                         
                if i not in dates.keys():    
                # Essayer de formatter la date 
                # CheatSheet formats: https://strftime.org/
                    try:
                        # Gerer d abord un cas particulier avec le mois d octobre qui peut etre confondu avec janvier
                        dates[i] = dt.datetime.strptime(str(cell), '%d.1').strftime('%d/10/'+ self._YEAR)
                    except ValueError or TypeError:
                        try:
                            dates[i] = dt.datetime.strptime(str(cell), '%d.%m').strftime('%d/%m/'+ self._YEAR)
                        except ValueError or TypeError:
                            try:
                                dates[i] = dt.datetime.strptime(str(cell), '%d.%m.%y').strftime('%d/%m/'+ self._YEAR)
                            except ValueError or TypeError:
                                try:
                                    dates[i] = dt.datetime.strptime(str(cell), '%-d.%m.%y').strftime('%d/%m/'+ self._YEAR)
                                except ValueError or TypeError:
                                    try:
                                        dates[i] = dt.datetime.strptime(str(cell), '%Y-%m-%d %H:%M:%S').strftime('%d/%m/'+ self._YEAR)
                                        # THIS IS NOT A DATE !
                                    except ValueError or TypeError:
                                        print(cell)
                                        pass
                else:
                    print("Error, entrée de date dupliquée pour la colonne ", i)
        return dates

    # This function return the debits from the corresponding column
    def getDebitDict(self):

        # Create a list to store the values 
        debits = dict()

        # Find the corresponding column
        (debitcolumn, idx) = self.findColumnFromLabel(["Débit", "DEBIT"])

        # Iterate over the cells in the column 
        for i, cell in enumerate(debitcolumn): 
            # If the cell is not empty
            if not pd.isnull(cell):
                if i not in debits.keys():
                    cell = Utils.formatNumber(cell)
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
    def getCreditDict(self):

        # Create a list to store the values 
        credits = dict()

        # Find the corresponding column
        (creditcolumn, idx) = self.findColumnFromLabel(["Crédit", "CREDIT"])

        # Iterate over the cells in the column 
        for i, cell in enumerate(creditcolumn): 
            # If the cell is not empty
            if not pd.isnull(cell):
                if i not in credits.keys():
                    cell = Utils.formatNumber(cell)
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
    def getLibelleDict(self):

        # Create a list to store the values 
        libelles = dict()

        # Find the corresponding column
        (libellecolumn, idx) = self.findColumnFromLabel(["Libellé", "Référence", "Nature", "NATURE"])

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
    