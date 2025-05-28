import pandas as pd
from odf.opendocument import load
from odf.text import P
from odf.table import Table, TableRow, TableCell  

# 1. Carregar o arquivo .ods
caminho_planilha = "Fundos_imobiliarios.ods"  
doc = load(caminho_planilha)

# 2. Listar todas as abas
abas = doc.getElementsByType(Table)
print(f"Abas disponíveis: {[aba.getAttribute('name') for aba in abas]}")

# 3. Ler cada aba e verificar os tipos de dados
for aba in abas:
    nome_aba = aba.getAttribute('name')
    print(f"\n=== ABA: '{nome_aba}' ===")
    
    # Converter a aba para DataFrame
    dados = []
    for linha in aba.getElementsByType(TableRow):
        celulas = linha.getElementsByType(TableCell)
        dados_linha = []
        for celula in celulas:
            # Extrai o texto de cada célula
            textos = []
            for p in celula.getElementsByType(P):
                for n in p.childNodes:
                    try:
                        textos.append(n.data)
                    except:
                        continue 
            dados_linha.append(" ".join(textos) if textos else "")
        dados.append(dados_linha)
    
    df = pd.DataFrame(dados[1:], columns=dados[0])  # Assume a 1ª linha como cabeçalho
    df = df[df['ATIVO'].notna() & (df['ATIVO'].astype(str).str.strip() != '')]  # Filtra linhas válidas    
    df['QUANTIDADE'] = pd.to_numeric(df['QUANTIDADE'], errors='coerce').fillna(0).astype(int)
    # Verificar tipos de dados
    print("\nTipos de dados originais:")
    print(df.dtypes)
    
    # Tentar converter colunas específicas
    if "DATA" in df.columns:
        try:
            df["DATA"] = pd.to_datetime(df["DATA"], format="%d/%m/%y", errors="coerce")
            print("\n✅ Coluna 'DATA' convertida para datetime.")
        except:
            print("\n❌ Falha ao converter 'DATA' para datetime.")
    
    if "VALOR UNIT" in df.columns:
        try:
            # Remove "R$" e vírgulas, converte para float
            df["VALOR UNIT"] = df["VALOR UNIT"].str.replace(r"[^\d,]", "", regex=True)  # Remove tudo exceto números e vírgula
            df["VALOR UNIT"] = df["VALOR UNIT"].str.replace(",", ".").astype(float)  # Converte para float 
            print("✅ Coluna 'VALOR UNIT' convertida para numérico (float).")
        except:
            print("❌ Falha ao converter 'VALOR UNIT' para numérico.")

    if not df.empty and {'ATIVO', 'QUANTIDADE', 'VALOR UNIT'}.issubset(df.columns):
    # Calcula totais por ativo
        resumo = df.groupby('ATIVO').agg(
            Quantidade_Total=('QUANTIDADE', 'sum'),
            Valor_Total_Investido=('VALOR UNIT', lambda x: (x * df.loc[x.index, 'QUANTIDADE']).sum()),  # Valor total = preço unitário × quantidade 
            Preço_Medio=('VALOR UNIT', 'mean')).reset_index()
    
    # Formata os valores
    resumo['Valor_Total_Investido'] = resumo['Valor_Total_Investido'].map('R$ {:,.2f}'.format)
    resumo['Preço_Medio'] = resumo['Preço_Medio'].map('R$ {:,.2f}'.format)
    
    print("\n📊 RESUMO POR ATIVO:")
    print(resumo.to_string(index=False))

    # Mostrar DataFrame processado
    print("\nConteúdo da aba:")
    with pd.option_context('display.max_rows', None, 'display.max_columns', None):
        print(df)
    
    print("\n" + "=" * 50)
