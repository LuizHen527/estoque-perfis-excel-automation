import xlwings as xw

def coloca_formula_aba_perfis_negativos():
    # Loop pelos perfis aba negativo
    # Pega nome e procura na planilha de balanco
    # Quando achar, salvar celula
    # colocar formula na lista
    planilha = xw.books.active
    perfil_anterior = ''
    row_perfil = False
    ACABAMENTOS = {
        'FOSCO': 'B',
        'BRANCO': 'C',
        'BRILHO': 'D',
        'PRETO': 'E',
        'BRONZE': 'F',
        'DOURADO': 'G',
        'ROSE': 'H',
        'INOX': 'I'
    }

    planilha.sheets['balanco'].activate()

    range_perfis_balanco = xw.Range("A4:A73")

    planilha.sheets['perfis balanco negativo'].activate()

    

    for cell in xw.Range("A2:A561"):
        perfil_procurado = cell.value

        column_perfil = ACABAMENTOS[cell.offset(0, 1).value]

        if perfil_anterior == perfil_procurado:
            cell.offset(0, 2).formula = f'=balanco!{column_perfil}${row_perfil}'
            continue

        planilha.sheets['balanco'].activate()
        # Loop pelos perfis pra achar a linha

        for perfil in range_perfis_balanco:
            if perfil.value == perfil_procurado:
                row_perfil = perfil.row
                break

        planilha.sheets['perfis balanco negativo'].activate()  

        perfil_anterior = perfil_procurado

        cell.offset(0, 2).formula = f'=balanco!{column_perfil}${row_perfil}'





        


def run():
    coloca_formula_aba_perfis_negativos()

if __name__ == '__main__':
    run()