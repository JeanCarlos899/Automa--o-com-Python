from openpyxl import load_workbook
import pyautogui as p
import time
import random
import PySimpleGUI as sg

class autoSisMoura:
    def __init__(self, path, TEMPO_ESPERA, valor_total):
        time.sleep(5)

        self.TEMPO_ESPERA = float(TEMPO_ESPERA)

        self.valor_total = valor_total

        self.path = path
        self.value_atual = 0
        self.contador = 0

        if valor_total == 0:
            raise ValueError('Valor total não pode ser 0')

        self.QTD_RODAR = valor_total // 1000

        while True:
            if self.valor_total - self.value_atual >= 1000:
                valor_venda = self.runAplication(self.path, self.TEMPO_ESPERA).run() 
                if type(valor_venda) == int or type(valor_venda) == float:
                    self.value_atual += valor_venda
                    self.contador += 1
                    self.finishing(hotkeyCloseSale = 'f5', hotkeyFinalize = 'f5')
                else:
                    raise ValueError('O arquivo não tem mais estoque para vender')
            else:
                break
        
        value = valor_total - self.value_atual
        self.value_atual += self.runAplication(self.path, self.TEMPO_ESPERA, value).run()
        self.finishing(hotkeyCloseSale = 'f5', hotkeyFinalize = 'f5')
        self.contador += 1

        p.alert(f"FINALIZADO: R$ {self.value_atual:.2f}, VENDAS: {self.contador}")

        print()
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        print(f"FINALIZADO: R$ {self.value_atual:.2f}, VENDAS: {self.contador}")
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        print()

    def finishing(self, hotkeyCloseSale, hotkeyFinalize):
        p.press(hotkeyCloseSale)
        time.sleep(3)
        p.click(x=69, y=233)
        time.sleep(3)
        p.press(hotkeyFinalize)
        time.sleep(3)

    class runAplication:
        def __init__(self, path, TEMPO_ESPERA, value_personalizado = None):
            # valueES DE EXECUÇÃO
            self.index = 0
            self.value_venda = 0
            self.value_personalizado = value_personalizado

            # CONFIGURAÇÕES
            self.path = path
            self.TEMPO_ESPERA = TEMPO_ESPERA

            # INICIAR PLANILHA
            self.planilha = load_workbook(self.path)
            self.aba_ativa = self.planilha.active

        def gerar_lista(self, coluna) -> list:
            lista = []
            for celula in self.aba_ativa[coluna]:
                values = celula.value
                if values != None and values != "" and type(values) != str:
                        lista.append(values)
            return lista

        def quantidade(self):
            estoque = self.gerar_lista("G")[self.index]

            if estoque <= 20:
                return 0
            elif estoque > 20 and estoque <= 50:
                return random.randint(1, 5)
            elif estoque > 50 and estoque <= 100:
                return random.randint(5, 10)
            elif estoque > 100:
                return random.randint(15, 20)
            
        def run(self):

            for produto in self.gerar_lista('A'):

                qtd = self.quantidade()

                if qtd == 0:
                    print("Produto com estoque baixo.")
                    self.index += 1
                    continue
                
                if self.value_personalizado:
                    if self.value_venda > self.value_personalizado or self.value_venda > 1000:
                        print('')                
                        print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
                        print(f"Venda concluída, valor total: R$ {self.value_venda:.2f}")
                        print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
                        print('')
                        print("==============================================================")
                        p.alert(f"Venda concluída, valor total: R$ {self.value_venda:.2f}")
                        
                        return self.value_venda

                else: 
                    if self.value_venda + (qtd * self.gerar_lista("N")[self.index]) > 1000:
                        print('')                
                        print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
                        print(f"Venda concluída, valor total: R$ {self.value_venda:.2f}")
                        print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
                        print('')
                        print("==============================================================")

                        p.alert(f"Venda concluída, value total: R$ {self.value_venda:.2f}")

                        return self.value_venda
                    
                p.click(x=304, y=114) 
                p.write(str(qtd))
                time.sleep(self.TEMPO_ESPERA)
                p.press('*')
                time.sleep(self.TEMPO_ESPERA)
                p.write(str(produto))
                time.sleep(self.TEMPO_ESPERA)
                p.press('enter')
                
                self.aba_ativa[f"G{self.index+4}"] = self.gerar_lista("G")[self.index] - qtd

                print(f"{qtd} unidades do produto {produto} adicionadas ao carrinho.")

                #ADICIONAR O PREÇO DO PRODUTO AO CONSOLE
                self.value_venda = self.value_venda + self.gerar_lista("N")[self.index] * qtd
                print(f"value total: {self.value_venda:.2f}")
                print("==============================================================")

                self.index += 1

                self.planilha.save(self.path)     
            
class windowAutoSisMoura:

    def window(self):
        sg.theme('DefaultNoMoreNagging')

        layout = [
            [sg.Text("Atomação Sistema Moura", size=(100, 1), font=("Helvetica", 25), justification="center")],
            [sg.Text("", font=(None, 1))],  
            [sg.Text('Informe o caminho do arquivo:', size=(30, 1))],
            [sg.InputText(key='-PATH-', size=(55, 1)), sg.FileBrowse(file_types=(("Excel", "*.xlsx"), ("All Files", "*.*")), size=(100, 1))],
            [sg.Text("Valor total da saída: (use '.' para separar casas decimais)")],
            [sg.InputText(key="-value-", size=(100, 1), default_text=0)],
            [sg.Text("Tempo de espera:", size=(15, 1))],
            [sg.InputText(key="-TIME-", size=(100, 1), default_text="0")],
            [sg.Button("Iniciar", size=(100, 2), button_color=("White", "#027F9E"), border_width=0)],

            [sg.Text("_________________________________________________________________________________", text_color="#FF8C01")],
            [sg.Text("", font=(None, 1))],
            [sg.Output(size=(200, 15), font=("Courier", 10), key="-OUTPUT-")],
            [sg.Text("", font=(None, 1))],
            [sg.Text("Criado por: Jean Carlos Rodrigues Sousa - Acauã - PI")],

        ]
        
        return sg.Window("AutoSisMoura", layout=layout, finalize=True, size=(600, 605))


if __name__ == "__main__":

    sysWindow = windowAutoSisMoura().window()
    
    while True:
        
        window, event, value = sg.read_all_windows()

        if event == sg.WIN_CLOSED:
            break

        elif event == "Iniciar":
            local_path = value["-PATH-"]
            valor = value["-value-"]
            tempo = value["-TIME-"]

            if local_path == "":
                sg.popup("Informe o caminho do arquivo.")
                continue

            else:
                try:
                    autoSisMoura(path=local_path, TEMPO_ESPERA=float(tempo), valor_total=float(valor))
                except Exception as e:
                    sg.popup("Erro, verificar o arquivo e entradas.")
                    print(f"Erro: {e}")

