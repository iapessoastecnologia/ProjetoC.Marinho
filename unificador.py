import pdfplumber
import pandas as pd
import re

def extrair_fornecedor1(linha):
    partes = linha.strip().split()
    if len(partes) < 8:
        return None
    try:
        # Código = posição 1 (Referência)
        codigo = partes[1]

        # Descrição = do idx 2 até antes do NCM (detectado pelo padrão de 8 dígitos)
        idx_ncm = next(i for i, v in enumerate(partes[2:], start=2) if re.match(r"^\d{8}$", v))
        descricao = " ".join(partes[2:idx_ncm])

        # Valor Unitário = campo após NCM + Qtde (logo após 2 posições)
        valor_unit_str = partes[idx_ncm + 2]
        valor = float(valor_unit_str.replace(".", "").replace(",", "."))

        return [codigo, descricao.strip(), valor, "fornecedor1"]
    except Exception as e:
        print(f"⚠️ Erro ao processar linha (fornecedor1): {linha}\n→ {e}")
        return None

def extrair_fornecedor2(linha):
    partes = linha.strip().split()
    if len(partes) < 8:
        return None
    try:
        # Código = sempre na posição 2
        codigo = partes[2]

        # Descrição = do idx 3 até antes do NCM (8 dígitos)
        idx_ncm = next(i for i, v in enumerate(partes[3:], start=3) if re.fullmatch(r"\d{8}", v))
        descricao = " ".join(partes[3:idx_ncm])

        # Valor Unitário = campo após NCM (R$ Unit.), que é idx_ncm + 1
        valor_unit_str = partes[idx_ncm + 1]
        valor = float(valor_unit_str.replace(".", "").replace(",", "."))

        return [codigo, descricao.strip(), valor, "fornecedor2"]
    except Exception as e:
        print(f"⚠️ Erro ao processar linha (fornecedor2): {linha}\n→ {e}")
        return None

def processar_pdf(caminho_pdf, fornecedor):
    dados = []
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if texto:
                for linha in texto.split("\n"):
                    if fornecedor == "fornecedor1":
                        resultado = extrair_fornecedor1(linha)
                    elif fornecedor == "fornecedor2":
                        resultado = extrair_fornecedor2(linha)
                    else:
                        resultado = None

                    if resultado:
                        dados.append(resultado)
    return dados

# Lista de arquivos e seus fornecedores
fornecedores = [
    ("fornecedor1.pdf", "fornecedor1"),
    ("fornecedor2.pdf", "fornecedor2")
]

# Processar todos e consolidar
todas_linhas = []
for caminho, nome in fornecedores:
    print(f"📄 Processando {caminho}...")
    linhas = processar_pdf(caminho, nome)
    print(f"✅ {len(linhas)} linhas extraídas de {nome}")
    todas_linhas.extend(linhas)

# Exportar Excel unificado
df = pd.DataFrame(todas_linhas, columns=["Código", "Descrição", "Valor", "Fornecedor"])
df.to_excel("orcamentos_unificados.xlsx", index=False)
print("✅ Arquivo 'orcamentos_unificados.xlsx' criado com sucesso!")
