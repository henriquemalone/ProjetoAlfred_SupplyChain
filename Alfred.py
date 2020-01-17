import tkinter as tk
from tkinter import *
from tkinter import messagebox
import pygubu
import scripts.popular33n
import scripts.popularLips
import scripts.attCliente
import scripts.attTransportadora

class MyApplication(pygubu.TkApplication):

    def _create_ui(self):
        #1: Create a builder
        self.builder = builder = pygubu.Builder()

        #2: Load an ui file
        builder.add_from_file('C:\\Alfred\\Interface\\interface2.ui')

        #3: Create the widget using self.master as parent
        self.mainwindow = builder.get_object('telaPrincipal', self.master)
        self.set_title("Alfred")

        # self.filepathmacro33 = builder.get_object('path33n1')
        self.filepathbase33 = builder.get_object('path33n2')

        # self.filepathmacrolips = builder.get_object('pathlips1')
        self.filepathbase33lips = builder.get_object('pathlips2')
        self.filepathbaselips = builder.get_object('pathlips3')

        self.filepathcliente = builder.get_object('pathattcliente')

        self.filepathtransp = builder.get_object('pathatttransp')

        # Configure callbacks
        builder.connect_callbacks(self)

    def updateTransp(self):
        pathtransp = self.filepathtransp.cget('path')
        scripts.attTransportadora.atualizar(pathtransp)

    def updateCliente(self):
        pathcliente = self.filepathcliente.cget('path')
        scripts.attCliente.atualizar(pathcliente)
    
    def populate33n(self):
        pathbase = self.filepathbase33.cget('path')
        if(scripts.popular33n.tratar(pathbase) == 1):
            scripts.popular33n.popular()

    def populatelips(self):
        pathbase33 = self.filepathbase33lips.cget('path')
        pathbaselips = self.filepathbaselips.cget('path')
        if(scripts.popularLips.tratar(pathbase33, pathbaselips) == 1):
            scripts.popularLips.popular()

if __name__ == '__main__':
    root = tk.Tk()
    app = MyApplication(root)
    app.run()