from tkinter import *
from tkinter import messagebox
from requests import get
from openpyxl import load_workbook

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
        path = self.input.get() # Pega o que foi digitado pelo usuário

        try:
            # Instância da planilha com base no caminho especificado
            planilha = load_workbook(path)

            # Instância das células da planilha
            celula = planilha.active

            # GET dos dados que irão para a planilha
            lista_dados = get('https://dummyjson.com/products')
            lista_dados = lista_dados.json()['products']

            # Definição do que cada coluna irá apresentar. Key é o nome do objeto vindo da API 
            coords = [
                { "col": "A", "key": "id", "format": None },
                { "col": "B", "key": "brand", "format": None },
                { "col": "C", "key": "title", "format": None },
                { "col": "D", "key": "description", "format": None },
                { "col": "E", "key": "price", "format": lambda price: "R$ " + str(price) },
                { "col": "F", "key": "discountPercentage", "format": lambda percentage: str(percentage) + '%' },
                { "col": "G", "key": "rating", "format": None },
                { "col": "H", "key": "stock", "format": None },
                { "col": "I", "key": "category", "format": None },
            ]

            for c in coords: # Loop nas colunas/coordenadas
                celula[c['col'] + '1'].value = c['key']

                for i in range(len(lista_dados)): # Loop nos objetos/dict da API
                    if c['format'] is not None: # Verifica se a coluna possui alguma função para formatar o valor da linha
                        celula[c['col'] + str(i + 2)].value = c['format'](lista_dados[i][c['key']])
                    else:
                        celula[c['col'] + str(i + 2)].value = lista_dados[i][c['key']] # Atualização do valor da célula [A1, A2, A3 ...]

            planilha.save(path) # Salva as alterações feitas na planilha
            messagebox.showinfo(title='SUCESSO', message='Planilha atualizada com sucesso') # Notificação de que a planilha foi alterada

            self.input.delete(0, 'end') # Limpa o input
            root.destroy() # Fecha a janela
        except:
            messagebox.showerror(title='ERRO', message='Não foi possível executar a operação') # Notificação de erro

root = Tk() # Instância do módulo de interface gráfica
root.title('Atualizar planilha')
root.resizable(False, False)
Application(root)
root.mainloop()