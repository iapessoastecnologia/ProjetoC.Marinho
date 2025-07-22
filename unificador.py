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
        partes = linha.strip().split()
        if len(partes) < 10:
            return None

        # Verifica se a linha começa com item numérico
        if not re.match(r'^\d{2,3}[A-Z]*$', partes[0]):
            return None

        # Código está sempre na posição 2
        codigo = partes[2]

        # Encontra índice do NCM (8 dígitos)
        idx_ncm = next(i for i, p in enumerate(partes) if re.fullmatch(r'\d{8}', p))
        
        # Descrição é tudo entre código e NCM
        descricao = ' '.join(partes[3:idx_ncm]).strip()

        # R$ Unit. está logo após o NCM
        valor_bruto = partes[idx_ncm + 1]
        valor = float(valor_bruto.replace('.', '').replace(',', '.'))

        return [codigo, descricao, valor, "fornecedor2"]
    except Exception as e:
        print(f"⚠️ Erro ao processar linha (fornecedor2): {linha}\n→ {e}")
        return None


def extrair_fornecedor3(linha):
    partes = linha.strip().split()
    if len(partes) < 6 or not re.fullmatch(r"\d{3,}-\d", partes[1]): return None
    try:
        codigo = partes[1]
        val_tok = partes[3]
        valor = float(val_tok.replace('.', '').replace(',', '.'))
        idx_ncm = next(i for i, v in enumerate(partes) if re.fullmatch(r"\d{8}", v))
        start, end = idx_ncm+1, next((i for i, v in enumerate(partes[idx_ncm+1:], idx_ncm+1) if v.isdigit()), len(partes))
        descricao = ' '.join(partes[start:end]).strip()
        return [codigo, descricao, valor, "fornecedor3"]
    except: return None

def extrair_fornecedor4(linha):
    partes = linha.strip().split()
    if len(partes) < 10: return None
    try:
        codigo = partes[3]
        idx_ncm = next(i for i, v in enumerate(partes) if re.fullmatch(r"\d{8}", v))
        descricao = ' '.join(partes[4:idx_ncm]).strip()
        valor_tok = next(t for t in partes[idx_ncm+1:] if re.match(r"^[\d\.]+,\d{2}$", t))
        valor = float(valor_tok.replace('.', '').replace(',', '.'))
        return [codigo, descricao, valor, "fornecedor4"]
    except: return None

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
fornecedores=[('fornecedor1.pdf','fornecedor1'),('fornecedor2.pdf','fornecedor2'),('fornecedor3.pdf','fornecedor3'),('fornecedor4.pdf','fornecedor4')]
all_data=[]
for caminho, prov in fornecedores:
    print(f'📄 Processando {caminho}...')
    all_data.extend(processar_pdf(caminho, prov))
df=pd.DataFrame(all_data,columns=['Código','Descrição','Valor','Fornecedor'])
df.to_excel('orcamentos_unificados.xlsx',index=False)
print('✅ Arquivo criado com sucesso!')