import xlwings as xw
import os
import glob
from datetime import datetime

def Hello():
    workbook = xw.books.active

    workbook.sheets[0].range("A1").value = "Batata"

#Não da pra polir muito o programa pq tenho 3 das pra fazer.
#Vou fazer o mvp e ir polindo com calma depois.

#Possiveis bugs para corrigir depois:
#   FileManager -> seleciona_planilha_por_nome
#       Pode dar erro se tiver mais de uma planilha de contagem


class FileManager:

    def __init__(self):

        # Sistema de rastreador de pedidos 
        self.relatorio_pedidos_pasta = r"\\121.137.1.5\manutencao1\Lucas\21_Luiz\Programas\controle-perfis-moducolor\relatorio-de-pedidos"

    def abrir_relatorio_pedidos(self):

        arquivos = os.listdir(self.relatorio_pedidos_pasta)

        arquivo = f"{self.relatorio_pedidos_pasta}\\{arquivos[0]}"

        xw.Book(arquivo)

    #Pode dar parte do nome
    def seleciona_planilha_por_nome(self, nome):
        for book in xw.books:
            if nome in book.name:
                book.activate(True)

    def abrir_planilha_a_faturar(self):
        ano = datetime.now().strftime('%y')
        mes = datetime.now().strftime('%m')
        dia = datetime.now().strftime('%d')

        arquivos = glob.glob(f"\\\\121.137.1.5\\manutencao1\\Lucas\\12_Relatorios\\20{ano}\\01_Relatorios Diarios\\01_Relatorios TecSerp\\{ano}_{mes}_*\\{ano}_{mes}_{dia}*A FATURAR*.xlsx")

        planilha_a_faturar = xw.Book(arquivos[0])

        planilha_a_faturar.sheets['Macro'].activate()


 



class DataProcessor:

    def __init__(self):
        pass

    def pegar_pedidos_relatorio(self):

        #Planilha ativa deve ser relatorio de pedidos
        planilha = xw.books.active.sheets[0]

        linha = planilha.range('A2').end('down').row

        data = planilha.range(f'A2:C{linha}').value

        return data

    #Retorna array com lista de pedidos já verificados 
    def pegar_log_pedidos(self):

        #Planilha ativa deve ser CONTROLE PERFIS

        log_pedidos = xw.sheets['log-pedidos']

        log_pedidos.activate()

        data = log_pedidos.tables[0].data_body_range

        data = data.value

        return data

    #Retorna array com endereço e numero do pedido
    def procurar_pedidos_por_range(self, range_pedidos, pedidos_procurar):
        pedidos_encontrados = []

        for p in range_pedidos:
        
            if None != p.value:
                numero = str(p.value)
                numero = numero[:-2]
    

                if numero in pedidos_procurar:

                    pedidos_encontrados.append([numero, p.address])

        return pedidos_encontrados




class Program:

    def __init__(self):
        self.file_man = FileManager()
        self.data_proc = DataProcessor()
        self.NOME_PLAN_CONTROLE = 'CONTROLE PERFIS'
        self.NOME_PLAN_FATURAR = 'A FATURAR'

    def pega_pedidos_relatorio(self):

        def filter_pedido(pedido):
            if pedido[1] == "Pronto":
                return True
            elif pedido[2] == "Análise de Custo" or pedido[2] == "—":
                return False
            else:
                return True


        pedidos_filtrados = []

        self.file_man.abrir_relatorio_pedidos()

        pedidos = self.data_proc.pegar_pedidos_relatorio()

        pedidos = filter(filter_pedido, pedidos)

        pedidos_filtrados = [p for p in pedidos]

        xw.books.active.close()

        return pedidos_filtrados

    def pega_pedidos_verificados(self):

        self.file_man.seleciona_planilha_por_nome("CONTROLE PERFIS")

        pedidos = self.data_proc.pegar_log_pedidos()

        return pedidos

    # Copia itens do pedido planilha do Lucas (Tecserp). Cola na planilha de controle
    # Ativar planilha Lucas antes de chamar
    def copia_itens_pedido_plan_tecserp(self, celula_num_pedido):
            celula_cima = celula_num_pedido.offset(-1, 0).value

            if celula_cima == None:
                range_pedido = xw.Range(f'A{celula_num_pedido.row}:AJ{celula_num_pedido.end('up').offset(1, 0).row}')
            else:
                range_pedido = celula_num_pedido

            range_pedido.copy()

            self.file_man.seleciona_planilha_por_nome(self.NOME_PLAN_CONTROLE)

            planilha_controle = xw.books.active

            planilha_controle.sheets['pedidos'].activate()

            xw.Range("A1").end('down').paste()

    def pega_produtos_pedidos(self, pedidos :list):

        pedidos_para_procurar = pedidos

        # A FATURAR

        self.file_man.abrir_planilha_a_faturar()

        numero_linhas = xw.Range('A2').end('down').row

        pedidos_a_faturar = xw.Range(f'E2:E{numero_linhas}')

        pedidos_encontrados = self.data_proc.procurar_pedidos_por_range(pedidos_a_faturar, pedidos_para_procurar)

        
        
        for pedido in pedidos_encontrados:

            celula = xw.Range(pedido[1])

            self.copia_itens_pedido_plan_tecserp(celula)

            self.file_man.seleciona_planilha_por_nome(self.NOME_PLAN_FATURAR)

            pedidos_para_procurar.remove(pedido[0])



        print('PEDIDOS PROCURAR:')
        
        [print(p) for p in pedidos_para_procurar]

        print('PEDIDOS ENCONTRADOS:')

        [print(p[0]) for p in pedidos_encontrados]

        #Saber se achou todos os pedidos
        for p in pedidos_para_procurar:
            foi_encontrado = False

            for pe in pedidos_encontrados:
                if pe[0] == p:
                    foi_encontrado = True
                    break
                    
            print(f'{p}     {foi_encontrado}')

        # FATURADO

        





    def atualiza_estoque_perfis(self):

        #Pedidos do sistema de rastreio
        pedidos_sistema = self.pega_pedidos_relatorio()

        #Pedidos dos logs
        pedidos_verificados = self.pega_pedidos_verificados()

        pedidos_novos = [p[0] for p in pedidos_sistema if p[0] not in pedidos_verificados]

        self.pega_produtos_pedidos(pedidos_novos)

        # Registrar pedido que esta pronto e que não esta pronto
        # Pensar bem em como fazer isso. Já que isso vai ditar o que as formulas vao fazer com o pedido 


def run():
    p = Program()

    p.atualiza_estoque_perfis()

    

if __name__ == '__main__':
    run()