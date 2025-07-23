import pdfplumber
import pandas as pd
import re
import os

def pre_processar_pdf_para_txt():
    """
    Função para processar o arquivo PDF do fornecedor4 e transformá-lo em TXT limpo.
    Este pré-processamento é executado apenas para o fornecedor4.
    """
    print("🔄 Pré-processando PDF do fornecedor4 para TXT...")
    
    # Lista de arquivos (apenas fornecedor4 neste caso)
    arquivos = ["fornecedor4.txt"]
    
    # Criar as pastas necessárias
    PASTA_ORIGEM = "processamento"
    PASTA_SAIDA = "txt_limpo"
    os.makedirs(PASTA_ORIGEM, exist_ok=True)
    os.makedirs(PASTA_SAIDA, exist_ok=True)
    
    # Primeiro, extrair o texto do PDF para um arquivo texto na pasta de processamento
    with pdfplumber.open("fornecedor4.pdf") as pdf:
        texto_completo = ""
        for pagina in pdf.pages:
            texto = pagina.extract_text() or ''
            texto_completo += texto + "\n"
            
        # Salvar o texto extraído do PDF
        with open(os.path.join(PASTA_ORIGEM, "fornecedor4.txt"), "w", encoding="utf-8") as f:
            f.write(texto_completo)
    
    # Regex para correção - apenas palavras espaçadas, NÃO números
    regex_palavras_espacadas = re.compile(r"\b(?:[A-Z]\s+){2,}[A-Z]\b")
    
    # Regex para eliminar linhas irrelevantes
    regex_irrelevante = re.compile(
        r"(CNPJ|Telefone|Fax|Endere[cç]o|CEP|Email|Transportadora|Emiss[aã]o|Pagina|Página|Validade|Subtotal|Total|Condi[cç][aã]o|Impress[aã]o|Pedido|Vendedor|Contato|Natureza|Moeda|Inscri[cç][aã]o|Origem|Orçamento|Cliente : C.)",
        re.IGNORECASE
    )
    
    # Detecção de itens válidos - mais simples
    regex_item_valido = re.compile(r'\b\d{1,4}\b.*\b[\w/-]{3,}\b.*\b\d{1,4}[.,]\d{2}\b')
    regex_linha_quebrada = re.compile(r"^\s*[\w/-]+\s*$")
    
    def limpar_linha(linha):
        linha = linha.strip()
    
        if not linha or regex_linha_quebrada.match(linha):
            return ""
    
        # Corrigir apenas palavras espaçadas (NÃO números)
        linha = regex_palavras_espacadas.sub(lambda m: m.group(0).replace(" ", ""), linha)
    
        # Adicionar espaço após 'tractorcraft' se estiver colado com número ou palavra
        linha = re.sub(r'(tractorcraft)([\w]+)', r'\1 \2', linha, flags=re.IGNORECASE)
    
        # Separar 'ItemEst.Marca', 'ItemEstMarca' ou variações coladas em 'Item Est Marca'
        linha = re.sub(r'Item\s*Est[\.]?\s*Marca', 'Item Est Marca', linha, flags=re.IGNORECASE)
    
        return linha
    
    def inserir_cabecalho_manual(linhas_originais):
        for linha in linhas_originais:
            limpa = limpar_linha(linha)
            if "Descrição" in limpa or "Seq" in limpa:
                return limpa
        return None
    
    # Processamento principal
    for nome_arquivo in arquivos:
        caminho_origem = os.path.join(PASTA_ORIGEM, nome_arquivo)
        caminho_saida = os.path.join(PASTA_SAIDA, nome_arquivo)
    
        with open(caminho_origem, "r", encoding="utf-8") as f:
            linhas = f.readlines()
    
        resultado = []
        linha_anterior = ""
        cabecalho_adicionado = False
    
        for linha in linhas:
            limpa = limpar_linha(linha)
    
            if not limpa or regex_irrelevante.search(limpa):
                continue
    
            # Detectar e adicionar o cabeçalho apenas uma vez
            if not cabecalho_adicionado and ("Descrição" in limpa or "Seq" in limpa):
                if linha_anterior:
                    resultado.append(linha_anterior)
                resultado.append(limpa)
                cabecalho_adicionado = True
                continue
    
            # Adiciona se for linha de item (mais permissivo)
            if regex_item_valido.search(limpa) or (len(limpa.split()) >= 8 and any(c.isdigit() for c in limpa)):
                resultado.append(limpa)
    
            # Atualiza a linha anterior válida
            linha_anterior = limpa if not regex_irrelevante.search(limpa) else ""
    
        # Inserir manualmente o cabeçalho, se não foi adicionado
        if not cabecalho_adicionado:
            cabecalho_manual = inserir_cabecalho_manual(linhas)
            if cabecalho_manual:
                resultado.insert(0, cabecalho_manual)
    
        # Remover a última linha se ela for muito diferente da penúltima ou antepenúltima em quantidade de palavras
        if len(resultado) >= 3:
            ultima = resultado[-1]
            penultima = resultado[-2]
            antepenultima = resultado[-3]
            len_ultima = len(ultima.split())
            len_penultima = len(penultima.split())
            len_antepenultima = len(antepenultima.split())
            # Se a última linha tiver menos da metade ou mais do dobro de palavras que as anteriores, remove
            if not (min(len_penultima, len_antepenultima) * 0.5 <= len_ultima <= max(len_penultima, len_antepenultima) * 2):
                resultado.pop()
    
        # Salvar saída
        with open(caminho_saida, "w", encoding="utf-8") as f_out:
            f_out.write("\n".join(resultado))
    
        print(f"✅ Arquivo limpo com cabeçalho + linha anterior salvo: {caminho_saida}")
    
    # Retornar o caminho do arquivo limpo
    return os.path.join(PASTA_SAIDA, "fornecedor4.txt")

def extrair_fornecedor1(linha):
    partes = linha.strip().split()
    if len(partes) < 8:
        return None
    try:
        codigo = partes[1]
        idx_ncm = next(i for i, v in enumerate(partes[2:], start=2) if re.match(r"^\d{8}$", v))
        raw = partes[2:idx_ncm]
        descricao_tokens = []
        temp = []
        for tok in raw:
            if len(tok) == 1:
                temp.append(tok)
            else:
                if temp:
                    descricao_tokens.append("".join(temp))
                    temp = []
                descricao_tokens.append(tok)
        if temp:
            descricao_tokens.append("".join(temp))
        descricao = " ".join(descricao_tokens).strip()
        valor_tok = next((t for t in partes[idx_ncm+1:] if re.match(r"^[\d\.]+,\d{2}$", t)), None)
        if not valor_tok:
            return None
        valor = float(valor_tok.replace(".", "").replace(",", "."))
        return [codigo, descricao, valor, "fornecedor1"]
    except Exception as e:
        print(f"⚠️ Erro ao processar linha (fornecedor1): {linha}\n→ {e}")
        return None

def extrair_fornecedor2(linha):
    try:
        # Extrair todos os valores monetários com vírgula
        valores = re.findall(r'\d{1,9}(?:\.\d{3})*,\d{2}', linha)
        if len(valores) < 4:
            return None

        # Valor unitário = 4º valor de trás pra frente
        valor_str = valores[-4]
        valor = float(valor_str.replace('.', '').replace(',', '.'))

        # Localiza posição do valor na string original
        match_valor = re.search(rf'{re.escape(valor_str)}', linha)
        if not match_valor:
            return None
        idx_valor = match_valor.start()

        # Tudo antes do valor é usado para extrair código e descrição
        parte_antes_valor = linha[:idx_valor].strip()

        # Dividimos os tokens antes do valor para buscar o código e descrição
        tokens = parte_antes_valor.split()

        if len(tokens) < 2:
            return None

        # O código é assumido como o último token antes do início da descrição
        codigo = tokens[2]

        # Descrição = tudo entre o código e o valor
        idx_codigo = parte_antes_valor.find(codigo)
        descricao = parte_antes_valor[idx_codigo + len(codigo):].strip()

        return [codigo.strip(), descricao.strip(), valor, "fornecedor2"]

    except Exception as e:
        print(f"⚠️ Erro ao processar linha (fornecedor2): {linha}\n→ {e}")
        return None

def extrair_fornecedor3(linha):
    try:
        # Procurar o primeiro código com pelo menos 5 dígitos (geralmente é o código do produto)
        codigo_match = re.search(r"\d{5,}", linha)
        if not codigo_match:
            return None
        codigo = codigo_match.group()

        # Buscar todos os valores monetários no padrão brasileiro
        valores = re.findall(r'\d{1,3}(?:\.\d{3})*,\d{2}', linha)
        if not valores:
            return None

        # Considerar o último valor como o valor do item
        valor_str = valores[-2]
        valor = float(valor_str.replace('.', '').replace(',', '.'))

        # Descrição = tudo entre o código e o valor
        idx_codigo = linha.find(codigo)
        idx_valor = linha.find(valor_str)
        if idx_codigo == -1 or idx_valor == -1 or idx_valor <= idx_codigo:
            return None
        descricao = linha[idx_codigo + len(codigo):idx_valor].strip()

        return [codigo.strip(), descricao.strip(), valor, "fornecedor3"]

    except Exception:
        return None

def extrair_fornecedor4(arquivo_texto='fornecedor4.txt'):
    """
    Extrai informações diretamente do arquivo fornecedor4.txt
    """
    dados = []
    try:
        with open(arquivo_texto, 'r', encoding='utf-8') as f:
            linhas = f.readlines()
            
        # Pula a primeira linha (cabeçalho)
        for linha in linhas[1:]:
            try:
                # Dividir a linha em colunas usando espaços como delimitador
                colunas = linha.strip().split()
                
                if len(colunas) < 6:
                    continue
                
                # As colunas no arquivo são: Item, Est, Marca, Código, Descrição, ...
                # Extrair marca e código diretamente das colunas corretas
                item_num = colunas[0]
                
                # Verificar se é uma linha de item válida (começa com número ou letra+número)
                if not re.match(r'^[0-9]+[A-Z]?$', item_num):
                    continue
                
                # A marca está na coluna 2 (índice 2)
                marca = colunas[2]
                
                # O código está na coluna 3 (índice 3)
                codigo = colunas[3]
                
                # Encontrar onde começa o NCM (8 dígitos)
                idx_ncm = -1
                for i, col in enumerate(colunas):
                    if re.fullmatch(r'\d{8}', col):
                        idx_ncm = i
                        break
                
                if idx_ncm == -1:
                    continue
                
                # A descrição são todas as colunas entre o código e o NCM
                descricao = ' '.join(colunas[4:idx_ncm])
                
                # Valor unitário: com padrão de vírgula para centavos
                valores = [c for c in colunas[idx_ncm:] if re.fullmatch(r'\d{1,3}(?:\.\d{3})*,\d{2}', c)]
                if not valores:
                    continue
                    
                # Geralmente é o segundo valor monetário após o NCM
                if len(valores) >= 2:
                    valor_str = valores[1]  # O segundo valor geralmente é o unitário
                else:
                    valor_str = valores[0]
                    
                valor = float(valor_str.replace('.', '').replace(',', '.'))
                
                dados.append([codigo, descricao, valor, "fornecedor4"])
                
            except Exception as e:
                print(f"⚠️ Erro ao processar linha do fornecedor4: {linha}\n→ {e}")
                continue
                
        return dados
    except Exception as e:
        print(f"⚠️ Erro ao processar arquivo do fornecedor4: {e}")
        return []

def processar_pdf(caminho_pdf, fornecedor):
    dados=[]
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            texto=pagina.extract_text() or ''
            for linha in texto.split('\n'):
                func = globals().get(f"extrair_{fornecedor}")
                if not func: continue
                res = func(linha)
                if res: dados.append(res)
    return dados

# Processar e exportar
fornecedores=[
    ('fornecedor1.pdf','fornecedor1'),
    ('fornecedor2.pdf','fornecedor2'),
    ('fornecedor3.pdf','fornecedor3')
]

all_data=[]
# Processa os PDFs normalmente para os fornecedores 1, 2 e 3
for caminho, prov in fornecedores:
    print(f'📄 Processando {caminho}...')
    all_data.extend(processar_pdf(caminho, prov))

# Para o fornecedor 4, primeiro faz o pré-processamento PDF → TXT, depois processa o TXT
print('📄 Processando fornecedor4...')
# Pré-processar o PDF para TXT limpo
caminho_txt_limpo = pre_processar_pdf_para_txt()
# Processar o TXT limpo
fornecedor4_data = extrair_fornecedor4(caminho_txt_limpo)
all_data.extend(fornecedor4_data)

# Criar DataFrame
df = pd.DataFrame(all_data, columns=['Código', 'Descrição', 'Valor', 'Fornecedor'])

# Exportar para Excel
df.to_excel('orcamentos_unificados.xlsx', index=False)
print('✅ Arquivo criado com sucesso!')
