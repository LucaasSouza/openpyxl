from openpyxl import load_workbook
from requests import get
from tkinter import *
from tkinter import messagebox

class Application:
    def __init__(self, master = None):
        self.frame = Frame(master, padx=15, pady=5, width=100)
        self.frame.pack()

        self.label = Label(self.frame, text='Digite o caminho do arquivo (.xlsx):')
        self.label.grid(row=1, column=1)

        self.input = Entry(self.frame, width=50)
        self.input.grid(row=1, column=2)

        self.frame_bottom = Frame(master, padx=10, pady=5)
        self.frame_bottom.pack()

        self.button = Button(self.frame_bottom, text='Atualizar planilha', command=lambda: self.atualizarPlanilha(), width=70)
        self.button.pack()

    def atualizarPlanilha(self):
        path = self.input.get()

        try:
            planilha = load_workbook(path)

            lista_usuarios = get('https://jsonplaceholder.typicode.com/users')
            lista_usuarios = lista_usuarios.json()

            celula = planilha.active
            coords = [
                { "col": "A", "key": "id" },
                { "col": "B", "key": "name" },
                { "col": "C", "key": "username" },
                { "col": "D", "key": "email" },
                { "col": "E", "key": "phone" },
            ]

            for c in coords:
                for i in range(len(lista_usuarios)):
                    celula[c['col'] + str(i + 1)].value = lista_usuarios[i][c['key']]

            planilha.save(path)
            messagebox.showinfo(title='SUCESSO', message='Planilha atualizada com sucesso')
            self.input.delete(0, 'end')
        except:
            messagebox.showerror(title='ERRO', message='Não foi possível executar a operação')

root = Tk()
root.title('Atualizar planilha')
root.resizable(False, False)
Application(root)
root.mainloop()