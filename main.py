import tkinter as tk
from tkinter import *
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from Excel2Quadra import Excel2Quadra
from Excel2Quadra import Mode

class Gui:

    def __init__(self):
        # create the root window
        root = tk.Tk()
        root.title('QuadraPyImport')
        root.resizable(False, False)
        root.geometry('600x400')
        root.eval('tk::PlaceWindow . center')

        self.mode = Mode.WRITE_ALL_DEB_CRED
        self.extraLibelle = True

        # Choix fichier entrée
        open_button = Button(
            root,
            text='Sélectioner le relevé bancaire à transformer (format Excel)',
            command=self.select_file
        )
        open_button.pack(expand=True)
        
        # Champ de saisie pour le nom du fichier de sortie
        label = Label(root, text="Nom du fichier Excel de sortie")
        label.pack(expand=True)
        self.entry = Entry(root, bd =5)
        self.entry.pack(expand=True)

    
        # setting variable for Integers
        self.modeButtonValue = StringVar()
        self.modeButtonValue.set(self.mode )

        # Menu des modes d ecriture
        dropdownMode = OptionMenu(
            root,
            self.modeButtonValue,
            *[mode for mode in Mode],
            command=self.getMode
        )
        dropdownMode.pack(expand=True)

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

        # run button
        run_button = Button(
            root,
            text='Lancer la conversion',
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
        '''
        showinfo(
            title='Erreur',
            message=fileIn
        )
        '''
    def getMode(self, choosedMode):
        self.mode = self.modeButtonValue.get()
        
    def setExtraLibelle(self):
        self.extraLibelle = self.exLibButtonValue.get()

    def run(self):
        
        self.fileOut = str(self.entry.get()) + '.xlsx'
        e2q = Excel2Quadra(self.fileIn, self.fileOut, self.extraLibelle, self.mode)
        e2q.runApp()



gui = Gui()
# run the application
gui.App.mainloop()

