import os
import tkinter as tk
from tkinter import filedialog

#selecionar pasta raiz colando caminho no terminal
def carregar_pasta():
    print("=== Leitor de Configurações de Equipamentos ===\n")
    pasta = input("Digite o caminho da pasta onde estão os arquivos: ")
    return pasta

#selecionar pasta raiz atraves do explorer do windows
def selecionar_pasta_raiz():
    """
    Abre o Windows Explorer para o usuário selecionar a pasta raiz.
    Retorna o caminho escolhido como string, ou None se o usuário cancelar.
    """
    # Cria uma janela principal "fantasma" (invisível)
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal do tkinter

    # Define as opções da janela de diálogo (opcional, mas recomendado)
    opcoes = {
        'title': 'Selecione a Pasta Raiz (Ex: MINHAPASTA)',
        'initialdir': os.path.expanduser('~')  # Abre por padrão na pasta do usuário (Documentos/Desktop)
    }

    # Abre o Explorer do Windows e aguarda a seleção
    pasta_selecionada = filedialog.askdirectory(**opcoes)

    # Verifica se o usuário escolheu uma pasta ou clicou em "Cancelar"
    if pasta_selecionada:
        # Se escolheu, o tkinter retorna o caminho com barras invertidas duplas.
        # O os.path.normpath ajuda a padronizar o caminho para o SO atual.
        return os.path.normpath(pasta_selecionada)
    else:
        # Se clicou em cancelar ou fechou a janela
        return None

def procurar_arquivo(pasta_raiz, nome_arquivo_procurado):
    """
    Varre a pasta raiz e todas as suas subpastas procurando por um arquivo específico.
    Retorna uma lista com os caminhos completos onde o arquivo foi encontrado.
    """
    # Verifica se a pasta base existe
    if not os.path.isdir(pasta_raiz):
        print(f"[ERRO] A pasta raiz '{pasta_raiz}' não foi encontrada.")
        return []

    #print(f"Iniciando busca por '{nome_arquivo_procurado}' em: {pasta_raiz}...\n")

    arquivos_encontrados = []

    # os.walk percorre a raiz e TODAS as subpastas em qualquer nível de profundidade
    for diretorio_atual, subpastas, arquivos in os.walk(pasta_raiz):

        # Verifica se o arquivo procurado está na lista de arquivos da pasta atual
        if nome_arquivo_procurado in arquivos:
            # Monta o caminho completo (Pasta + Nome do Arquivo)
            caminho_completo = os.path.join(diretorio_atual, nome_arquivo_procurado)

            # Adiciona na nossa lista de resultados
            arquivos_encontrados.append(caminho_completo)

            # Se você quiser executar uma função imediatamente ao achar o arquivo,
            # você pode chamá-la aqui, assim como fizemos com o analisar_cxf()
            # minha_funcao_de_analise(caminho_completo)

    # Resumo da operação
    #if arquivos_encontrados:
    #    print(f"Busca concluída! '{nome_arquivo_procurado}' encontrado em {len(arquivos_encontrados)} local(is).")
    #else:
    #    print(f"Busca concluída! Nenhum arquivo '{nome_arquivo_procurado}' foi encontrado.")

    return arquivos_encontrados

#percorre uma pasta em busca de configurações de caixa dentro de pastas CAIXAn/FLEX/CAIXA
def processar_todas_as_caixas(pasta_raiz):
    # Verifica se a pasta base existe
    if not os.path.isdir(pasta_raiz):
        print(f"[ERRO] A pasta raiz '{pasta_raiz}' não foi encontrada.")
        return

    print(f"Iniciando varredura profunda na pasta raiz: {pasta_raiz}...\n")
    arquivos_processados = 0

    # os.walk percorre a raiz e TODAS as subpastas, não importa o quão fundo estejam
    for diretorio_atual, subpastas, arquivos in os.walk(pasta_raiz):

        if "INSTALL.CXF" in arquivos:

            # Trava de segurança: Garante que só vamos ler se terminar com FLEX/CAIXA
            # (Isso protege contra arquivos perdidos na raiz MINHAPASTA ou MINHAPASTA/CAIXAS)
            caminho_esperado = os.path.join("FLEX", "CAIXA")

            if diretorio_atual.endswith(caminho_esperado):
                analisar_cxf(diretorio_atual)
                arquivos_processados += 1

    print(f"\nVarredura concluída! Total de arquivos INSTALL.CXF processados: {arquivos_processados}")

#arquivo INSTALL.CXF da caixa
def analisar_cxf(pasta):
    nome_arquivo = "INSTALL.CXF"
    caminho_arquivo = os.path.join(pasta, nome_arquivo)

    # Gera o nome do log dinamicamente mudando a extensão para .log
    nome_log = os.path.splitext(nome_arquivo)[0] + ".log"
    caminho_log = os.path.join(pasta, nome_log)

    if not os.path.isfile(caminho_arquivo):
        print(f"[AVISO] Arquivo não encontrado: {nome_arquivo}\n")
        return

    # Abre o arquivo em modo leitura binária ('rb')
    try:
        with open(caminho_arquivo, 'rb') as f:
            conteudo = f.read()
    except Exception as e:
        print(f"[ERRO] Falha ao ler o arquivo {nome_arquivo}: {e}\n")
        return

    if len(conteudo) < 1:
        print(f"[ERRO] O arquivo {nome_arquivo} está vazio.\n")
        return

    # Inicia a Análise do Binário
    num_blocos = conteudo[0]
    offset = 1  # O offset funciona como o nosso "cursor" de leitura
    dados_extraidos = {}
    tamanho_esperado = 1

    try:
        for _ in range(num_blocos):
            tamanho_bloco = conteudo[offset]
            offset += 1

            # Posição Inicial (High e Low byte)
            pos_high = conteudo[offset]
            pos_low = conteudo[offset + 1]
            posicao_inicial = (pos_high << 8) | pos_low
            offset += 2

            tamanho_esperado += 1 + ((tamanho_bloco + 1) * 2)

            # Leitura das informações do bloco
            for i in range(tamanho_bloco):
                val_high = conteudo[offset]
                val_low = conteudo[offset + 1]
                valor = (val_high << 8) | val_low
                offset += 2

                posicao_atual = posicao_inicial + i
                dados_extraidos[posicao_atual] = valor

    except IndexError:
        print("\n[AVISO] Arquivo terminou inesperadamente (pode estar corrompido).")

    arquivo_ok = (offset == len(conteudo)) and (len(conteudo) == tamanho_esperado)

    # ==========================================
    # IMPRESSÃO DOS RESULTADOS E GERAÇÃO DO LOG
    # ==========================================
    try:
        with open(caminho_log, 'w', encoding='utf-8') as f_out:
            def smart_print(texto):
                # print(texto)
                f_out.write(texto + "\n")

            smart_print("=" * 40)
            smart_print(" INFORMAÇÕES - CAIXA DE COMUTAÇÃO")
            smart_print("=" * 40)

            # Exceções Iniciais
            tipo_equipamento = dados_extraidos.get(32000, "[Não encontrado]")
            serial = dados_extraidos.get(32001, "[Não encontrado]")
            cod_cliente = dados_extraidos.get(32040, "[Não encontrado]")

            # Tratamento ASCII para Nome da Obra (32002 a 32033)
            nome_obra_chars = []
            for pos in range(32002, 32034):
                if pos in dados_extraidos:
                    v = (dados_extraidos[pos] >> 8)
                    if 32 <= v <= 126:
                        nome_obra_chars.append(chr(v))
                    else:
                        baixo = v & 0xFF
                        if 32 <= baixo <= 126:
                            nome_obra_chars.append(chr(baixo))

            nome_obra = "".join(nome_obra_chars).strip()
            if not nome_obra:
                nome_obra = "[Não encontrado ou vazio]"

            smart_print(f"Tipo Equipamento: {tipo_equipamento}")
            smart_print(f"Serial: {serial}")
            smart_print(f"Cód. Cliente: {cod_cliente}")
            smart_print(f"Nome Obra: {nome_obra}")

            smart_print("-" * 40)
            smart_print(" CONFIGURAÇÕES")
            smart_print("-" * 40)

            dadoPosicao = dados_extraidos.get(4, "[Não encontrado]")
            smart_print(f"Hora Inicial Leituras: {dadoPosicao}")

            dadoPosicao = dados_extraidos.get(5, "[Não encontrado]")
            smart_print(f"Hora Final Leituras: {dadoPosicao}")

            dadoPosicao = dados_extraidos.get(6, "[Não encontrado]")
            smart_print(f"Intervalo Leituras: {dadoPosicao}")

            dadoPosicao = dados_extraidos.get(7, "[Não encontrado]")
            smart_print(f"Forçar Leitura: {dadoPosicao}")

            dadoPosicao = dados_extraidos.get(8, "[Não encontrado]")
            smart_print(f"Hora Salvar Histórico das 16:00: {dadoPosicao}")

            dadoPosicao = dados_extraidos.get(30000, "[Não encontrado]")
            smart_print(f"Tipo Caixa: {dadoPosicao}")

            totalCabos_raw = dados_extraidos.get(30001, "[Não encontrado]")
            smart_print(f"Número de Cabos: {totalCabos_raw}")

            # Trava de segurança: garante que seja um número para fazer a conta do range
            totalCabos_calc = totalCabos_raw if isinstance(totalCabos_raw, int) else 0

            # --- INÍCIO DA ALTERAÇÃO DA TABELA ---
            smart_print("\n" + "-" * 30)
            smart_print(" TABELA DE CABOS E SENSORES")
            smart_print("-" * 30)

            # Formatação: Endereço (10 espaços alinhado à esquerda) | Sensores (10 espaços)
            formato_linha_cabos = "{:<10} | {:<10}"
            smart_print(formato_linha_cabos.format("Endereço", "Sensores"))
            smart_print("-" * 30)

            # cabos e seus sensores
            for pos in range(30002, 30002 + (totalCabos_calc * 2), 2):
                if pos in dados_extraidos:
                    enderecoCabo = dados_extraidos[pos]
                    numSensores = dados_extraidos.get(pos + 1, "[Erro]")

                    # Usa o mesmo formato para imprimir as variáveis (convertendo para string por segurança)
                    smart_print(formato_linha_cabos.format(str(enderecoCabo), str(numSensores)))

            smart_print("-" * 30 + "\n")
            # --- FIM DA ALTERAÇÃO ---

            # total de sensores
            dadoPosicao = dados_extraidos.get(30062, "[Não encontrado]")
            smart_print(f"Total de sensores da Caixa: {dadoPosicao}")

            # Mapear exceções para não repetir na lista geral
            posicoes_excecao = set(range(32000, 32034))
            posicoes_excecao |= set(range(4, 9))
            posicoes_excecao |= set(range(30000, 30002 + (totalCabos_calc * 2)))
            posicoes_excecao.add(30062)
            posicoes_excecao |= set(range(32034, 32041))

            smart_print("-" * 40)
            smart_print(" DADOS NÃO MAPEADOS")
            smart_print("-" * 40)

            for pos in sorted(dados_extraidos.keys()):
                if pos not in posicoes_excecao:
                    smart_print(f"Posição: {pos:5} : Valor: {dados_extraidos[pos]}")

            smart_print("\n" + "=" * 40)
            if arquivo_ok:
                smart_print(" STATUS: Arquivo OK")
            else:
                smart_print(" STATUS: Arquivo com problemas")
                smart_print(f" [!] Tamanho esperado: {tamanho_esperado} bytes | Tamanho real: {len(conteudo)} bytes")
            smart_print("=" * 40 + "\n")
        print(f"[OK] {nome_arquivo} analisado com sucesso -> {nome_log}")

    except Exception as e:
        print(f"\n[ERRO] Não foi possível salvar o arquivo de log: {e}")

#arquivo COEFVOL.SDW do TF/MW - valida o numero de unidades com o arquivo INSTALL.SDW
def analisar_coefvol(pasta, estado_unidades):
    #pasta = os.path.join(pasta, "MICROWIRELESS", "FLEX", "MICRONET", "CONFIG")
    nome_arquivo = "COEFVOL.SDW"
    #caminho_arquivo = os.path.join(pasta, nome_arquivo)
    caminho_arquivo = procurar_arquivo(pasta, nome_arquivo)
    caminho_arquivo = caminho_arquivo[0]

    pasta_do_arquivo = os.path.dirname(caminho_arquivo)
    nome_log = os.path.splitext(nome_arquivo)[0] + ".log"
    caminho_log = os.path.join(pasta_do_arquivo, nome_log)

    if not os.path.isfile(caminho_arquivo):
        print(f"[AVISO] Arquivo não encontrado: {nome_arquivo}\n")
        return

    # Proteção: Verifica se temos os dados do INSTALL.SDW para poder processar
    if not estado_unidades:
        print(
            f"[ERRO] Impossível processar {nome_arquivo}: Dados de Unidades/Cabos ausentes (INSTALL.SDW falhou ou não rodou).")
        return

    try:
        with open(caminho_arquivo, 'rb') as f:
            conteudo = f.read()
    except Exception as e:
        print(f"[ERRO] Falha ao ler o arquivo {nome_arquivo}: {e}\n")
        return

    if len(conteudo) < 1:
        print(f"[ERRO] O arquivo {nome_arquivo} está vazio..\n")
        return

    # O número de unidades é a quantidade de chaves no dicionário retornado pelo INSTALL.SDW
    num_unidades_esperadas = len(estado_unidades)
    tamanho_esperado = num_unidades_esperadas * 2
    arquivo_ok = (len(conteudo) == tamanho_esperado)

    try:
        with open(caminho_log, 'w', encoding='utf-8') as f_out:
            def smart_print(texto):
                # Escreve APENAS no log conforme solicitado
                f_out.write(texto + "\n")

            smart_print("=" * 45)
            smart_print(" INFORMAÇÕES - COEFICIENTE DE VOLUME (COEFVOL.SDW)")
            smart_print("=" * 45)
            smart_print(f" Unidades esperadas (INSTALL): {num_unidades_esperadas}")
            smart_print("-" * 45)

            # Processamos as unidades de acordo com o esperado
            for i in range(num_unidades_esperadas):
                offset = i * 2

                # Proteção caso o arquivo físico seja menor do que o INSTALL.SDW indicou
                if offset + 1 < len(conteudo):
                    val_high = conteudo[offset]
                    val_low = conteudo[offset + 1]
                    valor = (val_high << 8) | val_low
                    smart_print(f"Unidade {i + 1}: {valor}")
                else:
                    smart_print(f"Unidade {i + 1}: [ERRO - Dado ausente no arquivo]")

            smart_print("\n" + "=" * 45)
            if arquivo_ok:
                smart_print(" STATUS: Arquivo OK (Tamanho compatível com INSTALL)")
            else:
                smart_print(" STATUS: Arquivo com problemas")
                smart_print(f" [!] Tamanho esperado: {tamanho_esperado} bytes ({num_unidades_esperadas} un.)")
                smart_print(f" [!] Tamanho real:     {len(conteudo)} bytes")
            smart_print("=" * 45 + "\n")

        print(f"[OK] {nome_arquivo} analisado com sucesso -> {nome_log}")

    except Exception as e:
        print(f"\n[ERRO] Não foi possível salvar o arquivo de log para {nome_arquivo}: {e}")

#arquivo ALTCABOS.SDW
def analisar_altcabos(pasta, estado_unidades):
    #pasta = os.path.join(pasta, "MICROWIRELESS", "FLEX", "MICRONET", "CONFIG")
    nome_arquivo = "ALTCABOS.SDW"
    #caminho_arquivo = os.path.join(pasta, nome_arquivo)
    caminho_arquivo = procurar_arquivo(pasta, nome_arquivo)
    caminho_arquivo = caminho_arquivo[0]

    pasta_do_arquivo = os.path.dirname(caminho_arquivo)
    nome_log = os.path.splitext(nome_arquivo)[0] + ".log"
    caminho_log = os.path.join(pasta_do_arquivo, nome_log)

    if not os.path.isfile(caminho_arquivo):
        print(f"[AVISO] Arquivo não encontrado: {nome_arquivo}\n")
        return

    # Proteção: Verifica se temos os dados do INSTALL.SDW para poder processar
    if not estado_unidades:
        print(
            f"[ERRO] Impossível processar {nome_arquivo}: Dados de Unidades/Cabos ausentes (INSTALL.SDW falhou ou não rodou).")
        return

    try:
        with open(caminho_arquivo, 'rb') as f:
            conteudo = f.read()
    except Exception as e:
        print(f"[ERRO] Falha ao ler o arquivo {nome_arquivo}: {e}\n")
        return

    if len(conteudo) < 9:
        print(f"[ERRO] O arquivo {nome_arquivo} está vazio ou tem menos de 9 bytes.\n")
        return

    try:
        with open(caminho_log, 'w', encoding='utf-8') as f_out:
            def smart_print(texto):
                f_out.write(texto + "\n")

            smart_print("=" * 55)
            smart_print(" INFORMAÇÕES - ALTURA DE CABOS (ALTCABOS.SDW)")
            smart_print("=" * 55)

            offset = 0

            # Ordenamos as unidades (1, 2, 3...) para ler na sequência correta
            for unidade, total_cabos_unidade in sorted(estado_unidades.items()):

                # Proteção de leitura do cabeçalho da unidade (2 bytes)
                if offset + 2 > len(conteudo):
                    smart_print(f"\n[!] AVISO: Fim inesperado do arquivo antes da Unidade {unidade}.")
                    break

                # --- CABEÇALHO DA UNIDADE ---
                alt_max = (conteudo[offset] << 8) | conteudo[offset + 1]
                offset += 2

                smart_print(f"\n--- Unidade {unidade} ---")
                smart_print(f" Altura Máx. do Produto: {alt_max}")
                smart_print(f" Total Cabos Esperados:  {total_cabos_unidade}")

                # --- LEITURA DOS BLOCOS DE CABOS ---
                cabos_lidos = 0
                seq = 1  # Contador apenas para exibição

                while cabos_lidos < total_cabos_unidade:
                    # Proteção de leitura do bloco de cabos (7 bytes)
                    if offset + 7 > len(conteudo):
                        smart_print(f"\n[!] AVISO: Fim inesperado durante os blocos da Unidade {unidade}.")
                        break

                    num_cabos_seq = conteudo[offset]
                    alt_inicial = (conteudo[offset + 1] << 8) | conteudo[offset + 2]
                    alt_sensor_1 = (conteudo[offset + 3] << 8) | conteudo[offset + 4]
                    dist_sensores = (conteudo[offset + 5] << 8) | conteudo[offset + 6]
                    offset += 7

                    smart_print(f"   > Sequência {seq} ({num_cabos_seq} cabos)")
                    smart_print(f"     Altura Inicial Chão:   {alt_inicial}")
                    smart_print(f"     Altura 1º Sensor:      {alt_sensor_1}")
                    smart_print(f"     Distância Sensores:    {dist_sensores}")

                    # Incrementa os cabos processados com a quantidade desta sequência
                    cabos_lidos += num_cabos_seq
                    seq += 1

            # Validação do arquivo: o offset final deve ser exatamente o tamanho do arquivo
            arquivo_ok = (offset == len(conteudo))

            smart_print("\n" + "=" * 55)
            if arquivo_ok:
                smart_print(" STATUS: Arquivo OK (Estrutura lida integralmente)")
            else:
                smart_print(" STATUS: Arquivo com problemas")
                smart_print(f" [!] Bytes lidos: {offset} | Tamanho real: {len(conteudo)} bytes.")
                if offset < len(conteudo):
                    smart_print(f"     Existem {len(conteudo) - offset} bytes sobrando no final do arquivo.")
            smart_print("=" * 55 + "\n")

        print(f"[OK] {nome_arquivo} analisado com sucesso -> {nome_log}")

    except Exception as e:
        print(f"\n[ERRO] Não foi possível salvar o arquivo de log para {nome_arquivo}: {e}")

#arquivo INSTALL.SDW para MW e TF
def analisar_install(pasta):
    #pasta = os.path.join(pasta, "MICROWIRELESS", "FLEX", "MICRONET", "CONFIG")
    nome_arquivo = "INSTALL.SDW"
    #caminho_arquivo = os.path.join(pasta, nome_arquivo)
    caminho_arquivo = procurar_arquivo(pasta, nome_arquivo)
    caminho_arquivo = caminho_arquivo[0]

    pasta_do_arquivo = os.path.dirname(caminho_arquivo)
    nome_log = os.path.splitext(nome_arquivo)[0] + ".log"
    caminho_log = os.path.join(pasta_do_arquivo, nome_log)

    if not os.path.isfile(caminho_arquivo):
        print(f"[AVISO] Arquivo não encontrado: {nome_arquivo}\n")
        return

    # Abre o arquivo em modo leitura binária ('rb')
    try:
        with open(caminho_arquivo, 'rb') as f:
            conteudo = f.read()
    except Exception as e:
        print(f"[ERRO] Falha ao ler o arquivo {nome_arquivo}: {e}\n")
        return

    if len(conteudo) < 202:
        print(f"[ERRO] O arquivo {nome_arquivo} está vazio ou tem menos de 202 bytes.\n")
        return

    # --- Extração do Cabeçalho ---
    tipo_equipamento = (conteudo[0] << 8) | conteudo[1]
    serial = (conteudo[2] << 8) | conteudo[3]

    # Tratamento ASCII da Obra (Bytes 4 a 75 = 36 words = 72 bytes)
    nome_obra_chars = []
    for pos in range(4, 76, 2):
        high = conteudo[pos]
        low = conteudo[pos + 1]
        if 32 <= high <= 126:
            nome_obra_chars.append(chr(high))
        elif 32 <= low <= 126:
            nome_obra_chars.append(chr(low))

    nome_obra = "".join(nome_obra_chars).strip()
    if not nome_obra:
        nome_obra = "[Não encontrado ou vazio]"

    versao_sw = (conteudo[76] << 8) | conteudo[77]
    versao_fw = (conteudo[78] << 8) | conteudo[79]
    cod_cliente = (conteudo[80] << 8) | conteudo[81]

    # Validação dos bytes 198 (#CR) e 199 (#LF)
    cr_valid = (conteudo[198] == 13)
    lf_valid = (conteudo[199] == 10)

    num_cabos = (conteudo[200] << 8) | conteudo[201]

    # Validação do arquivo (202 bytes de cabeçalho + 10 bytes por cabo)
    tamanho_esperado = 202 + (num_cabos * 10)
    arquivo_ok = (len(conteudo) == tamanho_esperado)

    # --- Dicionários de Tradução ---
    MAP_TIPO_CAIXA = {0: "120s", 1: "192s", 2: "Digital", 3: "Trena"}
    MAP_TIPO_UNID = {1: "S. Plano", 2: "S. Cônico", 3: "Arm. Plano", 4: "Arm. Semi-V",
                     5: "Arm. V", 6: "Arm. W", 7: "S. Semi-V", 8: "S. Melita"}
    MAP_TIPO_CABO = {0: "Temp (+)", 1: "Temp (+-)", 2: "Temp + U"}

    # Variável de Estado para Pós-Processamento
    # Formato: { unidade_id: max_cabo }
    cabos_por_unidade = {}

    try:
        with open(caminho_log, 'w', encoding='utf-8') as f_out:
            def smart_print(texto):
                # Escreve APENAS no arquivo de log, mantendo o terminal limpo
                f_out.write(texto + "\n")

            smart_print("=" * 110)
            smart_print(" INFORMAÇÕES - INSTALAÇÃO (INSTALL.SDW)")
            smart_print("=" * 110)

            smart_print(f"Tipo Equipamento:    {tipo_equipamento}")
            smart_print(f"Número de Série:     {serial}")
            smart_print(f"Nome da Obra:        {nome_obra}")
            smart_print(f"Versão Software:     {versao_sw}")
            smart_print(f"Versão Firmware:     {versao_fw}")
            smart_print(f"Código do Cliente:   {cod_cliente}")
            smart_print(f"Validação CR/LF:     {'OK' if (cr_valid and lf_valid) else 'FALHA (Assinatura incorreta)'}")
            smart_print(f"Total de Cabos:      {num_cabos}")

            smart_print("\n" + "-" * 110)
            smart_print(" LISTAGEM DE CABOS E SENSORES")
            smart_print("-" * 110)

            # Cabeçalho da Tabela
            # Unidade | Cabo | Arco | Pos. Arco | Sensores | Caixa | Offset | Tipo Caixa | Tipo Unid. | Tipo Cabo
            formato_linha = "{:<8} | {:<5} | {:<5} | {:<10} | {:<9} | {:<6} | {:<7} | {:<11} | {:<12} | {:<10}"
            smart_print(formato_linha.format("Unidade", "Cabo", "Arco", "Pos. Arco", "Sensores", "Caixa", "Offset",
                                             "Tipo Caixa", "Tipo Unid.", "Tipo Cabo"))
            smart_print("-" * 110)

            # Processamento do laço de cabos
            offset = 202
            # Proteção: iterar até o num_cabos, mas limitando ao tamanho do arquivo para não estourar
            for _ in range(num_cabos):
                if offset + 10 > len(conteudo):
                    print(f"\n[AVISO] Arquivo {nome_arquivo} terminou antes de listar todos os cabos.")
                    break

                tipo_caixa_raw = conteudo[offset]
                caixa = conteudo[offset + 1]
                offset_sensor = conteudo[offset + 2]
                sensores = conteudo[offset + 3]
                arco = conteudo[offset + 4]
                pos_arco = conteudo[offset + 5]
                cabo = (conteudo[offset + 6] << 8) | conteudo[offset + 7]
                unidade = conteudo[offset + 8]
                byte9 = conteudo[offset + 9]

                # Deslocamento de Bits para o Byte 9
                tipo_unid_raw = byte9 & 0x0F  # Pega os 4 bits da direita (0-3)
                tipo_cabo_raw = (byte9 >> 4) & 0x0F  # Pega os 4 bits da esquerda (4-7)

                # Traduções de Texto
                str_tipo_caixa = MAP_TIPO_CAIXA.get(tipo_caixa_raw, f"Desc({tipo_caixa_raw})")
                str_tipo_unid = MAP_TIPO_UNID.get(tipo_unid_raw, f"Desc({tipo_unid_raw})")
                str_tipo_cabo = MAP_TIPO_CABO.get(tipo_cabo_raw, f"Desc({tipo_cabo_raw})")

                # Salva o estado para Pós-Processamento
                if unidade not in cabos_por_unidade:
                    cabos_por_unidade[unidade] = cabo
                elif cabo > cabos_por_unidade[unidade]:
                    cabos_por_unidade[unidade] = cabo

                # Imprime a linha da tabela
                linha = formato_linha.format(
                    unidade, cabo, arco, pos_arco, sensores, caixa, offset_sensor,
                    str_tipo_caixa, str_tipo_unid, str_tipo_cabo
                )
                smart_print(linha)

                offset += 10  # Avança 10 bytes para o próximo bloco

            smart_print("\n" + "=" * 110)
            if arquivo_ok:
                smart_print(" STATUS: Arquivo OK")
            else:
                smart_print(" STATUS: Arquivo com problemas")
                smart_print(f" [!] Tamanho esperado: {tamanho_esperado} bytes | Tamanho real: {len(conteudo)} bytes")
            smart_print("=" * 110 + "\n")

        print(f"[OK] {nome_arquivo} analisado com sucesso -> {nome_log}")

        # Retorna os dados coletados para uso futuro!
        return cabos_por_unidade

    except Exception as e:
        print(f"\n[ERRO] Não foi possível salvar o arquivo de log para {nome_arquivo}: {e}")
        return None


if __name__ == '__main__':
    # A pasta é solicitada uma vez e passada para todos os processadores
    #pasta_selecionada = carregar_pasta()
    pasta_selecionada = selecionar_pasta_raiz()

    if pasta_selecionada:
        #analisar_cxf(pasta_selecionada)
        processar_todas_as_caixas(pasta_selecionada)
        retornoUnidades = analisar_install(pasta_selecionada)
        analisar_coefvol(pasta_selecionada, retornoUnidades)
        analisar_altcabos(pasta_selecionada, retornoUnidades)

        if retornoUnidades:
            total_unid = len(retornoUnidades)
            print(f"[INFO] INSTALL.SDW informou {total_unid} unidades no total.")

        print("\nProcessamento concluído.")