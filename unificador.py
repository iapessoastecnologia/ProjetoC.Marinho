import pdfplumber
import pandas as pd
import re

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

def extrair_fornecedor4(linha):
    try:
        partes = linha.strip().split()
        if len(partes) < 10:
            return None

        # Encontrar o índice do NCM (8 dígitos)
        idx_ncm = next(i for i, v in enumerate(partes) if re.fullmatch(r"\d{8}", v))

        # Identificar código como o primeiro "token misto" após a marca
        codigo = next(
            (p for i, p in enumerate(partes[:idx_ncm]) if i >= 3 and re.search(r'[A-Za-z]', p) and len(p) >= 5),
            None
        )
        if not codigo:
            return None

        idx_codigo = partes.index(codigo)

        # Descrição: entre código e NCM
        descricao_tokens = partes[idx_codigo + 1:idx_ncm]
        descricao = ' '.join(descricao_tokens).strip()

        # Valor Unitário: penúltimo valor monetário real (sem %)
        valores = [p for p in partes if re.fullmatch(r'\d{1,3}(?:\.\d{3})*,\d{2}', p) and '%' not in p]
        if len(valores) < 2:
            return None
        valor_str = valores[-4]
        valor = float(valor_str.replace('.', '').replace(',', '.'))

        return [codigo.strip(), descricao, valor, "fornecedor4"]

    except Exception as e:
        print(f"⚠️ Erro ao processar linha (fornecedor4): {linha}\n→ {e}")
        return None

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
    ('fornecedor3.pdf','fornecedor3'),
    ('fornecedor4.pdf','fornecedor4')
]

all_data=[]
for caminho, prov in fornecedores:
    print(f'📄 Processando {caminho}...')
    all_data.extend(processar_pdf(caminho, prov))

df=pd.DataFrame(all_data,columns=['Código','Descrição','Valor','Fornecedor'])
df.to_excel('orcamentos_unificados.xlsx',index=False)
print('✅ Arquivo criado com sucesso!')
