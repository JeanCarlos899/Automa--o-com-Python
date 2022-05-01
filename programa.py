from openpyxl import load_workbook
import pyautogui as p
import time
import random
import PySimpleGUI as sg

class autoSisMoura:
    def __init__(self, path, TEMPO_ESPERA, valor_total, x, y):
        time.sleep(5)

        self.TEMPO_ESPERA = float(TEMPO_ESPERA)

        self.valor_total = valor_total

        self.path = path
        self.value_atual = 0
        self.contador = 0
        self.x = x
        self.y = y

        if valor_total == 0:
            raise ValueError('Valor total não pode ser 0')

        self.QTD_RODAR = valor_total // 1000

        while True:
            if self.valor_total - self.value_atual >= 1000:
                valor_venda = self.runAplication(self.path, self.TEMPO_ESPERA, x=self.x, y=self.y).run() 
                if type(valor_venda) == int or type(valor_venda) == float:
                    self.value_atual += valor_venda
                    self.contador += 1
                    self.finishing(hotkeyCloseSale = 'f5', hotkeyFinalize = 'f5')
                else:
                    raise ValueError('O arquivo não tem mais estoque para vender')
            else:
                break
        
        value = valor_total - self.value_atual
        self.value_atual += self.runAplication(self.path, self.TEMPO_ESPERA, x=self.x, y=self.y, value_personalizado=value).run()
        self.finishing(hotkeyCloseSale = 'f5', hotkeyFinalize = 'f5')
        self.contador += 1

        p.alert(f"FINALIZADO: R$ {self.value_atual:.2f}, VENDAS: {self.contador}")

        print()
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        print(f"FINALIZADO: R$ {self.value_atual:.2f}, VENDAS: {self.contador}")
        print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
        print()

    def finishing(self, hotkeyCloseSale, hotkeyFinalize):
        # p.press(hotkeyCloseSale)
        # time.sleep(3)
        # p.click(x=69, y=233)
        # time.sleep(3)
        # p.press(hotkeyFinalize)
        # time.sleep(3)
        pass

    class runAplication:

        def __init__(self, path, TEMPO_ESPERA, x, y, value_personalizado = None):
            # valueES DE EXECUÇÃO
            self.index = 0
            self.value_venda = 0
            self.value_personalizado = value_personalizado

            # CONFIGURAÇÕES
            self.path = path
            self.TEMPO_ESPERA = TEMPO_ESPERA
            self.x = x
            self.y = y

            # INICIAR PLANILHA
            self.planilha = load_workbook(self.path)
            self.aba_ativa = self.planilha.active

            # GERAR LISTA
            self.estoque = self.gerar_lista("G")
            self.codigos = self.gerar_lista("A")
            self.precos = self.gerar_lista("N")

        def gerar_lista(self, coluna) -> list:
            lista = []
            for celula in self.aba_ativa[coluna]:
                values = celula.value
                if values != None and values != "" and type(values) != str:
                        lista.append(values)
            return lista

        def quantidade(self):
            if self.estoque[self.index] <= 20:
                return 0
            elif self.estoque[self.index] > 20 and self.estoque[self.index] <= 50:
                return random.randint(1, 5)
            elif self.estoque[self.index] > 50 and self.estoque[self.index] <= 100:
                return random.randint(5, 10)
            elif self.estoque[self.index] > 100:
                return random.randint(15, 20)
            
        def atualizar_estoque(self):
            print("Atualizando o arquivo...")
            for index in range(len(self.codigos)):
                self.aba_ativa[f"G{index+4}"] = self.estoque[index]
            self.planilha.save(self.path)  
            print("Arquivo atualizado com sucesso!")

        def run(self):

            for produto in self.codigos:

                qtd = self.quantidade()

                if qtd == 0:
                    print("Produto com estoque baixo.")
                    self.index += 1
                    continue
                
                if self.value_personalizado:
                    if self.value_venda >= self.value_personalizado or self.value_venda + (qtd * self.precos[self.index]) > 1000:
                        print('')                
                        print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
                        print(f"Venda concluída, valor total: R$ {self.value_venda:.2f}")
                        print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
                        print('')
                        print("==============================================================")
                        p.alert(f"Venda concluída, valor total: R$ {self.value_venda:.2f}")

                        self.atualizar_estoque()
                        return self.value_venda

                else: 
                    if self.value_venda + (qtd * self.precos[self.index]) > 1000:
                        print('')                
                        print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
                        print(f"Venda concluída, valor total: R$ {self.value_venda:.2f}")
                        print("++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++")
                        print('')
                        print("==============================================================")
                        p.alert(f"Venda concluída, value total: R$ {self.value_venda:.2f}")

                        self.atualizar_estoque()
                        return self.value_venda
                    
                p.click(self.x, self.y) 
                p.write(str(qtd))
                time.sleep(self.TEMPO_ESPERA)
                p.press('*')
                time.sleep(self.TEMPO_ESPERA)
                p.write(str(produto))
                time.sleep(self.TEMPO_ESPERA)
                p.press('enter')
                
                self.estoque[self.index] = self.estoque[self.index] - qtd
  
                print(f"{qtd} unidades do produto {produto} adicionadas ao carrinho.")

                #ADICIONAR O PREÇO DO PRODUTO AO CONSOLE
                self.value_venda = self.value_venda + self.precos[self.index] * qtd
                print(f"value total: {self.value_venda:.2f}")
                print("==============================================================")

                self.index += 1
                
class windowAuto:

    def window(self):
        sg.theme('DefaultNoMoreNagging')

        layout = [
            [sg.Text("", font=(None, 1))],  
            [sg.Frame("Selecione o sistema",
                [
                    [
                        sg.Radio("SISMOURA", "RADIO1", key="-SISMOURA-", default=True),
                        sg.Radio("SOFTCOM", "RADIO1", key="-SOFTCOM-"),
                    ]

                ], size=(1920, 60)
            )],
            [sg.Text('Informe o caminho do arquivo:', size=(30, 1))],
            [sg.InputText(key='-PATH-', size=(55, 1)), sg.FileBrowse(file_types=(("Excel", "*.xlsx"), ("All Files", "*.*")), size=(200, 1))],
            [sg.Text("Valor total da saída: (use '.' para separar casas decimais)")],
            [sg.InputText(key="-value-", size=(1920, 1), default_text=0)],
            [sg.Text("Tempo de espera:", size=(1920, 1))],
            [sg.InputText(key="-TIME-", size=(1920, 1), default_text="0")],
            [sg.Button("Iniciar", size=(1920, 2), button_color=("White", "#027F9E"), border_width=0)],
            [sg.Button("Calibrar clique do mouse", size=(1920, 2), button_color=("White", "#027F9E"), border_width=0, key="-MOUSE-")],

            [sg.Text(500*"_", text_color="#FF8C01")],
            [sg.Text("", font=(None, 1))],
            [sg.Output(size=(200, 15), font=("Courier", 10), key="-OUTPUT-")],
            [sg.Text("", font=(None, 1))],
            [sg.Text("Criado por: Jean Carlos Rodrigues Sousa | Acauã - PI", justification="center", size=(1920, 1))],

        ]
        
        return sg.Window("AutoSisMoura", layout=layout, finalize=True, size=(600, 780), resizable=True)


if __name__ == "__main__":

    sysWindow = windowAuto().window()
    
    while True:
        
        window, event, value = sg.read_all_windows()

        if event == sg.WIN_CLOSED:
            break

        elif event == '-MOUSE-':
            print("Mova o mouse para a posição desejada em 5 segundos.")
            time.sleep(5)
            x, y = p.position()
            print("Posição do mouse:", p.position())
            print("Calibração concluída.")

        elif event == "Iniciar":
            local_path = value["-PATH-"]
            valor = value["-value-"]
            tempo = value["-TIME-"]

            sismoura = value["-SISMOURA-"]

            if local_path == "":
                sg.popup("Informe o caminho do arquivo.")
                continue

            else:
                if sismoura == True:
                    try:
                        autoSisMoura(path=local_path, TEMPO_ESPERA=float(tempo), valor_total=float(valor), x=x, y=y)
                    except Exception as e:
                        sg.popup("Erro, verificar o arquivo, calibração e entradas.")
                        print(f"Erro: {e}")
                else:
                    # try:
                    #     autoSoftcom(path=local_path, TEMPO_ESPERA=float(tempo), valor_total=float(valor))
                    # except Exception as e:
                    #     sg.popup("Erro, verificar o arquivo e entradas.")
                    #     print(f"Erro: {e}")
                    pass

