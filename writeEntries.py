import pandas as pd
from utils import QuadraPyUtils as Utils


class WriteEntries:
        
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

    def __init__(self, dates, debits, credits, libelles, defcompte, ctrpartie):
        self._DEFCOMPTE = defcompte
        self._CTRPARTIE = ctrpartie
        self.datevalues = dates
        self.debitvalues = debits
        self.creditvalues = credits
        self.libellevalues = libelles
        self.dfout = pd.DataFrame(self.entryLines)

    def addDebit(self, date, libelle, numPiece, debit):
        
        newEntry = {
            'Date' : date, 
            'Compte': self._DEFCOMPTE, 
            'Libellé': libelle , 
            'Contrepartie': self._CTRPARTIE, 
            'Num.Pièce': numPiece, 
            'Débit': debit, 
            'Crédit': ''
        }
        newCtPrtEntry = {
            'Date' : date, 
            'Compte': self._CTRPARTIE, 
            'Libellé': libelle, 
            'Contrepartie': self._CTRPARTIE, 
            'Num.Pièce': numPiece, 
            'Débit': '', 
            'Crédit': debit
        }

        return (newEntry, newCtPrtEntry)


    def addCredit(self, date, libelle, numPiece, credit):
        newEntry = {
            'Date' : date, 
            'Compte': self._DEFCOMPTE, 
            'Libellé': libelle , 
            'Contrepartie': self._CTRPARTIE, 
            'Num.Pièce': numPiece, 
            'Débit': '', 
            'Crédit': credit
        }
        newCtPrtEntry = {
            'Date' : date, 
            'Compte': self._CTRPARTIE, 
            'Libellé': libelle, 
            'Contrepartie': self._CTRPARTIE, 
            'Num.Pièce': numPiece, 
            'Débit': credit, 
            'Crédit': ''
        }

        return (newEntry, newCtPrtEntry)



    def writeNewEntry(self, newEntry, newCtrpEntry):

        # Ajouter la nouvelle entrée et sa contrepartie
        self.dfout = pd.concat([self.dfout,pd.DataFrame([newEntry, newCtrpEntry,])], ignore_index=True)
        # Add an empty row for readability
        self.dfout.loc[len(self.dfout)] = pd.Series()


    # Ecrire les credits dans le fichier excel de sortie 
    def writeDebits(self):

        for debitrow in self.debitvalues:
            
            # Oui, un debit est sur la meme ligne !
            debit = self.debitvalues[debitrow]
            try:
                # Recuperer le libelle de l operation courante
                libelle = self.libellevalues[debitrow]
            except KeyError:
                print("Pas de key pour le libelle associe au debit courant: ",debit)
                pass            
            else:
                # Si il ne s'agit pas d'une ligne décrivant le solde
                if not (("solde" in libelle) or ("Solde" in libelle)):
                    try:
                        # Recuperer la date courante
                        date = self.datevalues[debitrow]
                    except KeyError:
                        print("Pas de key pour la date associe au debit courant: ", debit)
                        pass  
                    else: 
                        # Get the numPiece
                        numPiece = Utils.getNumPiece(libelle)

                        # Creer une entree de crédit
                        (newEntry, newCtrpEntry) = self.addDebit(date,libelle, numPiece, debit)
                        # Ajouter la nouvelle entrée
                        self.writeNewEntry(newEntry, newCtrpEntry)


    # Ecrire les debits dans le fichier excel de sortie 
    def writeCredits(self):

        for creditrow in self.creditvalues:

            # Oui, un credit est sur la meme ligne !
            credit = self.creditvalues[creditrow]

            try:
                # Recuperer le libelle de l operation courante
                libelle = self.libellevalues[creditrow]
            except KeyError:
                print("Pas de key pour le libelle associe au credit courant : ", credit)
                pass            
            else:
                # Si il ne s'agit pas d'une ligne décrivant le solde
                if not (("solde" in libelle) or ("Solde" in libelle)):
                    try:
                        # Recuperer la date courante
                        date = self.datevalues[creditrow]
                    except KeyError:
                        print("Pas de key pour la date associe au credit courant : ", credit)
                        pass  
                    else: 
                        # Get the numPiece
                        numPiece = Utils.getNumPiece(libelle)

                        # Creer une entree de crédit
                        (newEntry, newCtrpEntry) = self.addCredit(date,libelle, numPiece, credit)
                        # Ajouter la nouvelle entrée
                        self.writeNewEntry(newEntry, newCtrpEntry)


    # Ecrire les debits puis les credits dans le fichier excel de sortie 
    def writeAllDebitsThenCredits(self):

        self.writeDebits()
        self.writeCredits()



    # Ecrire d'après l'ordre des dates les debits et les credits dans le fichier excel de sortie 
    def writeAllByDate(self):

        # Pour chaque date
        for daterow in self.datevalues:

            # Recuper la date courante
            date = self.datevalues[daterow]

            # Recuperer le libelle de l operation courante
            libelle = self.libellevalues[daterow]
        
            # Get the numPiece
            numPiece = Utils.getNumPiece(libelle)

            # Si il ne s'agit pas d'une ligne décrivant le solde
            if not (("solde" in libelle) or ("Solde" in libelle)):

                # Est ce un debit
                try:
                    # Oui, un debit est sur la meme ligne !
                    debit = self.debitvalues[daterow]
                    # Creer une entree de débit
                    (newEntry, newCtrpEntry) = self.addDebit(date,libelle, numPiece, debit)
                # Non !
                except KeyError:
                    # Est ce un credit
                    try:
                        # Oui, un credit est sur la meme ligne !
                        credit = self.creditvalues[daterow]
                        # Creer une entree de crédit
                        (newEntry, newCtrpEntry) = self.addCredit(date,libelle, numPiece, credit)
                    except KeyError:
                        print("Pas de données de débit ou crédit pour la date du ",self.datevalues[daterow])
                    else:
                        # Ajouter la nouvelle entrée
                        self.writeNewEntry(newEntry, newCtrpEntry)
                else:   
                    # Ajouter la nouvelle entrée
                    self.writeNewEntry(newEntry, newCtrpEntry)