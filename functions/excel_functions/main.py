from openpyxl import load_workbook


def iniciar_planilha(conteudo):
    """
    Inicia o arquivo Excel contendo os dados.
    :return: Retorna a planilha.
    """
    try:
        # open('PlanilhaExcel\\Arquivo.xlsx', 'a').close()
        planilha = load_workbook(filename=conteudo, data_only=True)
    except PermissionError:
        print('O arquivo já está aberto, feche o mesmo antes de prosseguir.')
        exit()
    except:
        print('Arquivo não encontrado.')
        exit()
    else:
        return planilha


def pegar_dados_intervalo_planilha(conteudo, intervalo: str, ultima_linha: bool = False) -> list:
    """
    Retorna os valores presentes no intervalo informado.
    Os valores são retornados dentro de uma lista.
    A lista retornada contém listas para cada linha do intervalo informado.
    Ex: [['Pessoa1', 46, 2500, 'Jogador', '987654321'], ['Pessoa2', 22, 8000, 'Streamer', '768948302']]
    :param intervalo: Intervalo da planilha.
    :param ultima_linha: Define se deverá ser pego até a ultima linha do intervalo informado.
    :return: Retorna uma lista contendo os valores.
    """
    if ultima_linha:
        intervalo: str = intervalo + descobrir_linha_vazia_planilha_excel(conteudo, intervalo[0])

    planilha = iniciar_planilha(conteudo)
    aba_ativa = planilha.active

    try:
        valores: list = []
        valores_linha: list = []

        # Adicionei os if's para impedir que dados vazios sejam obtidos.
        for celula in aba_ativa[intervalo]:
            for elemento in celula:
                if elemento.value is not None:
                    valores_linha.append(elemento.value)
                else:
                    break  # -> Talvez eu possa tirar esse else e deixar ele pegar uma linha onde um elemento seja None.
            if len(valores_linha) > 0:
                valores.append(valores_linha.copy())
                valores_linha.clear()
            # else:
                # break  # -> Para a procura se achar algum registro vazio.
    except:
        planilha.close()
        print('Error - pegar_dados_intervalo_planilha()')
    else:
        planilha.close()
        return valores


def descobrir_linha_vazia_planilha_excel(conteudo, coluna: str) -> str:
    """
    Descobre o número da ultima linha preenchida na planilha da coluna informada.
    Mesmo se houver linhas vazias entres linhas preenchidas, a última será pega.
    :return: Retorna o número da ultima linha como uma string.
    """
    planilha = iniciar_planilha(conteudo)
    aba_ativa = planilha.active

    ultima_posicao: int = aba_ativa[coluna][-1].row

    ultima_linha: int = 1
    cont: int = 0

    for i, celula in enumerate(aba_ativa[f'{coluna}1:{coluna}{ultima_posicao}']):
        if celula[0].value is not None:
            ultima_linha: int = celula[0].row
        else:
            # Ignoro a quantidade de linhas que eu percebi que eu preciso passar para chegar na proxima linha NÃO vazia:
            if cont > 0:
                cont -= 1  # Diminuo o cont pois acabei de passar por uma linha que estou ignorando.
            else:
                # aba_ativa[coluna][i].row == celula[0].row  | Ambos são a mesma coisa.
                if aba_ativa[coluna][i].row < ultima_posicao:
                    # VERIFICANDO SE EXISTE ALGUMA LINHA PREENCHIDA DA POSIÇÃO ATUAL ATE A ULTIMA LINHA DA COLUNA:
                    quantidade_linha_faltante: int = ultima_posicao - celula[0].row
                    elemento: int = 0  # -> Variável utilizada para verificar se foi encontrada uma linha NÃO vazia.
                    # Loop baseado na quantidade de linhas faltantes, se faltam 7 linhas para acabar,
                    # o loop vai rodar 7 vezes.
                    for c in range(1, quantidade_linha_faltante+1):
                        valor = aba_ativa[coluna][i+c].value  # PEGANDO O VALOR PRESENTE NA LINHA *SEGUINTE*.
                        # ^ Para verificar em seguida se a mesma é uma linha vazia ou não.
                        if valor is not None:
                            elemento: int = 1  # -> Indicador de que encontrei uma linha NÃO vazia.
                            cont = c-1  # -> Quantidade de linhas que posso ignorar até a proxima linha NÃO vazia.
                            break
                            # ^ Se encontro alguma linha não vazia no meio do caminho, nem continuo olhando os próximos,
                            # pois já achei o que queria.
                    if elemento == 0:  # Se não encontrei nenhuma linha preenchida saio do loop e já sei a última linha.
                        break
                else:
                    break

    planilha.close()

    return str(ultima_linha)
