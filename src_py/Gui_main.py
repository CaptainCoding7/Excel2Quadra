from tkinter import *
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from Excel2Quadra import Excel2Quadra
from Excel2Quadra import Mode


class Gui:

    def __init__(self):
        # create the root window
        root = Tk()
        root.title('Excel2Quadra')
        root.resizable(False, False)
        root.geometry('600x400')
        root.eval('tk::PlaceWindow . center')

        self.mode = Mode.WRITE_ALL_DEB_CRED
        self.extraLibelle = True
        self.dictCompte = dict()

        # Choix fichier entrée
        open_button = Button(
            root,
            text='Sélectionner le relevé bancaire à transformer (format Excel)',
            command=self.select_file
        )
        open_button.pack(expand=True)
        
        # Champ de saisie pour le nom du fichier de sortie
        label = Label(root, text="Nom du fichier Excel de sortie")
        label.pack(expand=True)

   
        defaulOutFile = StringVar() 
        defaulOutFile.set("output") 
        self.entryFileout = Entry(root, bd =5, textvariable = defaulOutFile)
        self.entryFileout.pack(expand=True)

    
        # setting variable for Integers
        self.modeButtonValue = StringVar()
        self.modeButtonValue.set(self.mode.value)

        # Menu des modes d ecriture
        dropdownMode = OptionMenu(
            root,
            self.modeButtonValue,
            *[mode.value for mode in Mode],
            command=self.getMode
        )
        dropdownMode.pack(expand=True)

        # Champ de saisie pour le numero de compte
        label2 = Label(root, text="Numéro de compte par défaut")
        label2.pack(expand=True)

        defNumCompte = StringVar() 
        defNumCompte.set("47100000") 
        self.entryNumCompte = Entry(root, bd =5, textvariable = defNumCompte)
        self.entryNumCompte.pack(expand=True)


        # Champ de saisie pour le numero de contrepartie
        label3 = Label(root, text="Numéro de contrepartie")
        label3.pack(expand=True)

        defContrpartie = StringVar() 
        defContrpartie.set("51200000") 
        self.entryCtrPartie = Entry(root, bd =5, textvariable = defContrpartie)
        self.entryCtrPartie.pack(expand=True)

        # enable extralibelle  checkbox
        self.exLibButtonValue = BooleanVar()
        self.exLibButtonValue.set(self.extraLibelle )
        extraLibelleBut = Checkbutton(
            root,
            text='Ecrire la 2e ligne de libellé',
            variable=self.exLibButtonValue,
            command=self.setExtraLibelle
        )
        extraLibelleBut.pack(expand=True)


        # newCompte button
        newCompte_button = Button(
            root,
            text='Ajouter un compte récurrent',
            command=self.addNewCompte
        )
        newCompte_button.pack(expand=True)

        # run button
        run_button = Button(
            root,
            text='Transformer le relevé',
            command=self.run
        )
        run_button.pack(expand=True)


        self.fileIn = None
        self.App = root

    def select_file(self):

        filetypes = (
            ('Excel', '*.xlsx'),
            ('All files', '*.*')
        )

        self.fileIn = fd.askopenfilename(
            title='Open a File',
            initialdir='./',
            filetypes=filetypes)


    def getMode(self, choosedMode):
        self.mode = Mode(choosedMode)

    def setExtraLibelle(self):
        self.extraLibelle = self.exLibButtonValue.get()

    def addNewCompte(self):

        newCompteFen = Tk()
        newCompteFen.title('Ajouter un numéro de compte récurrent')
        newCompteFen.resizable(False, False)
        newCompteFen.geometry('200x200')
        newCompteFen.eval('tk::PlaceWindow . center')  

        # Champ de saisie pour le nouveau numero de compte
        label1 = Label(newCompteFen, text="Nouveau numéro de compte")
        label1.pack(expand=True)
        newCompte = StringVar() 
        newCompte.set("47100000") 
        self.entryNewCompte = Entry(newCompteFen, bd =5, textvariable = newCompte)
        self.entryNewCompte.pack(expand=True)

        # Champ de saisie pour l extrait de libelle associ&é
        label2 = Label(newCompteFen, text="Extrait de libellé associé")
        label2.pack(expand=True)
        libelleCompte = StringVar() 
        self.entryLibelleCompte = Entry(newCompteFen, bd =5, textvariable = libelleCompte)
        self.entryLibelleCompte.pack(expand=True)

        # run button
        addCompte_button = Button(
            newCompteFen,
            text='Ajouter',
            command=self.run_addCompte
        )
        addCompte_button.pack(expand=True)
        self.newCompteFen = newCompteFen


    def popUpMsg(self, strTitle, strDescr):
        showinfo(
            title = strTitle,
            message = strDescr
        )

    def run_addCompte(self):

        compte = str(self.entryNewCompte.get()) 
        libelle = str(self.entryLibelleCompte.get()) 
        self.dictCompte[compte] = libelle
        self.newCompteFen.destroy()

    def run(self):

        self.fileOut = str(self.entryFileout.get()) + '.xlsx'
        self.defNumCompte = str(self.entryNumCompte.get()) 
        self.ctrPartie = str(self.entryCtrPartie.get()) 
        e2q = Excel2Quadra(self.fileIn, self.fileOut, self.extraLibelle, self.mode, self.defNumCompte, self.ctrPartie, self.dictCompte)
        e2q.runApp()

gui = Gui()
# run the application
gui.App.mainloop()

