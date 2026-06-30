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


    def abrir_planilha_faturado(self):
        ano = datetime.now().strftime('%y')
        mes = datetime.now().strftime('%m')
        dia = datetime.now().strftime('%d')

        arquivos = glob.glob(f"\\\\121.137.1.5\\manutencao1\\Lucas\\12_Relatorios\\20{ano}\\01_Relatorios Diarios\\01_Relatorios TecSerp\\{ano}_{mes}_*\\{ano}_{mes}_{dia}*FATURADOS*.xlsx")

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

        primeira_celula = xw.Range('A2').value

        if primeira_celula != None:
            data = log_pedidos.tables[0].data_body_range

            data = data.value
        else:
            data = []

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

    # pedidos_sistema -> [['122018', 'Pronto', 'Pedido Pronto'], ...]
    # pedidos_não_encontrados -> ['122151', '122155', ...]
    # Retorna pedidos_status -> [['122018', 'Pronto'], ...]
    def verifica_status_pedidos(self, pedidos_sistema, pedidos_não_encontrados):

        def pega_status(pedido):
            if pedido[1] == 'Pronto':
                return [pedido[0], 'PRONTO']
            else:
                return [pedido[0], 'PRODUZINDO']

        def coloca_pedidos_não_encontrados(pedido):
            if pedido[0] in pedidos_não_encontrados:
                return [pedido[0], 'NÃO ENCONTRADO']
            else:
                return pedido

        pedidos_status = map(pega_status, pedidos_sistema)

        pedidos_status = map(coloca_pedidos_não_encontrados, pedidos_status)

        return list(pedidos_status)




class Program:

    def __init__(self):
        self.file_man = FileManager()
        self.data_proc = DataProcessor()
        self.NOME_PLAN_CONTROLE = 'CONTROLE PERFIS'
        self.NOME_PLAN_FATURAR = 'A FATURAR'
        self.NOME_PLAN_FATURADOS = 'FATURADOS'

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
                range_pedido = xw.Range(f'A{celula_num_pedido.row}:AJ{celula_num_pedido.row}')

            range_pedido.copy()

            self.file_man.seleciona_planilha_por_nome(self.NOME_PLAN_CONTROLE)

            planilha_controle = xw.books.active

            planilha_controle.sheets['pedidos'].activate()

            if xw.Range("A1").end('down').value == None:
                xw.Range("A1").end('down').paste()
            else:
                xw.Range("A1").end('down').offset(1, 0).paste()
            

    #Copia pedidos da planilha a faturar pra planilha de controle
    def copia_pedidos_a_faturar(self, pedidos):
        pedidos_para_procurar = pedidos
        
        self.file_man.abrir_planilha_a_faturar()

        numero_linhas = xw.Range('A2').end('down').row

        pedidos_a_faturar = xw.Range(f'E2:E{numero_linhas}')

        pedidos_encontrados = self.data_proc.procurar_pedidos_por_range(pedidos_a_faturar, pedidos_para_procurar)
        
        for pedido in pedidos_encontrados:

            celula = xw.Range(pedido[1])

            self.copia_itens_pedido_plan_tecserp(celula)

            self.file_man.seleciona_planilha_por_nome(self.NOME_PLAN_FATURAR)

            pedidos_para_procurar.remove(pedido[0])

        #Fecha A FATURAR
        xw.books.active.close()

        #Retorna pedidos que não achou no A FATURAR
        return pedidos_para_procurar

    def copia_pedidos_faturados(self, pedidos):
        pedidos_para_procurar = pedidos
                
        self.file_man.abrir_planilha_faturado()

        numero_linhas = xw.Range('A2').end('down').row

        pedidos_faturados = xw.Range(f'E2:E{numero_linhas}')

        pedidos_encontrados = self.data_proc.procurar_pedidos_por_range(pedidos_faturados, pedidos_para_procurar)
        
        for pedido in pedidos_encontrados:

            celula = xw.Range(pedido[1])

            self.copia_itens_pedido_plan_tecserp(celula)

            self.file_man.seleciona_planilha_por_nome(self.NOME_PLAN_FATURADOS)

            pedidos_para_procurar.remove(pedido[0])

        #Fecha faturado
        xw.books.active.close()

        #Retorna pedidos que não achou no A FATURAR
        return pedidos_para_procurar

    
    #Retorna pedidos que não foram encontrados
    def pega_produtos_pedidos(self, pedidos :list):   

        pedidos_restantes = self.copia_pedidos_a_faturar(pedidos)

        pedidos_restantes = self.copia_pedidos_faturados(pedidos_restantes)

        return pedidos_restantes
        
    def coloca_status_na_planilha(self, pedidos_status):

        xw.sheets['log-pedidos'].activate()

        primeira_celula = xw.Range('A2')

        if primeira_celula.value == None:
            primeira_celula.value = pedidos_status
        else:

            ultima_celula = xw.Range('A2').end('down')

            pedidos_na_planilha = xw.Range(F'A2:A{ultima_celula.row}')

            pedidos_novos = []

            #Atualiza status dos pedidos que já estão na planilha
            for pedido in pedidos_status:

                for pedido_planilha in pedidos_na_planilha:
                    numero_pedido_planilha = str(int(pedido_planilha.value))

                    if numero_pedido_planilha == pedido[0]:
                        status_planilha = pedido_planilha.offset(0, 1)
                        status_planilha.value = pedido[1]
                        break
                else:
                    pedidos_novos.append(pedido)

            ultima_celula.offset(1, 0).value = pedidos_novos

                



            



    def atualiza_estoque_perfis(self):

        #Pedidos do sistema de rastreio
        pedidos_sistema = self.pega_pedidos_relatorio()

        #Pedidos dos logs
        pedidos_verificados = self.pega_pedidos_verificados()

        pedidos_novos = [p[0] for p in pedidos_sistema if p[0] not in pedidos_verificados]

        pedidos_nao_encontrados = self.pega_produtos_pedidos(pedidos_novos)

        pedidos_status = self.data_proc.verifica_status_pedidos(pedidos_sistema, pedidos_nao_encontrados)

        self.coloca_status_na_planilha(pedidos_status)




        #Fazer função que pega pedidos_sistema e pedidos_nao_encontrados e coloca status neles
        #Coloca status convorme a planilha
        #Coloca status de nao encontrado nos pedidos nao encontrados




def run():
    p = Program()

    pedidos_status = [['122051', 'PRONTO'], ['122137', 'PRODUZINDO'], ['122136', 'PRODUZINDO'], ['122141', 'PRODUZINDO'], ['122139', 'PRODUZINDO'], ['122142', 'PRODUZINDO'], ['122149', 'PRODUZINDO'], ['122122', 'PRONTO'], ['122143', 'PRONTO'], ['122098', 'PRONTO'], ['122133', 'PRONTO'], ['122114', 'PRONTO'], ['122070', 'PRONTO'], ['122152', 'PRONTO'], ['122018', 'PRONTO'], ['122151', 'NÃO ENCONTRADO'], ['121906', 'PRODUZINDO'], ['122155', 'NÃO ENCONTRADO'], ['122144', 'PRONTO'], ['122163', 'NÃO ENCONTRADO'], ['122145', 'PRONTO'], ['122164', 'NÃO ENCONTRADO'], ['122007', 'PRONTO'], ['122146', 'NÃO ENCONTRADO'], ['122170', 'NÃO ENCONTRADO'], ['116572', 'NÃO ENCONTRADO'], ['122179', 'PRONTO'], ['122165', 'PRONTO'], ['122158', 'PRODUZINDO'], ['121991', 'PRODUZINDO'], ['122212', 'PRODUZINDO'], ['122153', 'NÃO ENCONTRADO'], ['122134', 'PRONTO'], ['122058', 'PRONTO'], ['122198', 'NÃO ENCONTRADO'], ['122127', 'PRODUZINDO'], ['122044', 'PRODUZINDO'], ['122080', 'PRODUZINDO'], ['122205', 'PRODUZINDO'], ['122206', 'PRODUZINDO'], ['122197', 'PRONTO'], ['122215', 'PRONTO'], ['120991', 'PRODUZINDO'], ['121370', 'NÃO ENCONTRADO'], ['121737', 'NÃO ENCONTRADO'], ['122250', 'PRODUZINDO'], ['121895', 'PRONTO'], ['122236', 'PRONTO'], ['122242', 'PRONTO'], ['122231', 'PRONTO'], ['122251', 'PRONTO'], ['122219', 'NÃO ENCONTRADO'], ['122258', 'PRODUZINDO'], ['122257', 'PRONTO'], ['122126', 'PRODUZINDO'], ['122275', 'PRONTO'], ['122283', 'PRONTO'], ['122314', 'NÃO ENCONTRADO'], ['122299', 'PRONTO'], ['120525', 'PRONTO']]

    p.atualiza_estoque_perfis()

    

    

if __name__ == '__main__':
    run()