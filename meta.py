import pandas as pd
import matplotlib.pyplot as plt

# Dividend Yield configuration for each asset
dy_por_ativo = {
    'HGRE11': 9.13,
    'HGBS11': 9.86,
    'VILG11': 8.89,
    'HGRU11': 9.53
}

def calcular_dividendos(caminho_planilha):
    try:
        # Read ODS file
        df = pd.read_excel(caminho_planilha, sheet_name=None, engine='odf')
    except Exception as e:
        print(f"Error reading file: {e}")
        return

    resultados = []
    
    for aba_nome, dados_aba in df.items():
        ativo = aba_nome.upper()
        if ativo not in dy_por_ativo:
            continue
        
        try:
            # Create copy and filter valid rows
            dados_aba = dados_aba.copy()
            dados_aba = dados_aba[
                dados_aba['ATIVO'].notna() & 
                (dados_aba['ATIVO'].astype(str).str.strip() != '')
            ]
            
            # Convert numeric columns
            dados_aba.loc[:, 'QUANTIDADE'] = pd.to_numeric(
                dados_aba['QUANTIDADE'], errors='coerce'
            ).fillna(0).astype(int)
            
            dados_aba.loc[:, 'VALOR UNIT'] = (
                dados_aba['VALOR UNIT'].astype(str)
                .str.replace(r'[^\d,]', '', regex=True)
                .str.replace(',', '.')
                .astype(float)
            )
            
            # Calculate totals
            total_cotas = dados_aba['QUANTIDADE'].sum()
            dy = dy_por_ativo[ativo] / 12
            dividendo_mensal = (total_cotas * dy)  # Monthly dividend
            
            resultados.append({
                'Ativo': ativo,
                'Total de Cotas': total_cotas,
                'DY Mensal': f"{dy:.2f}%",
                'Dividendo Mensal': f"R$ {dividendo_mensal:.2f}",
                'Valor Numerico': dividendo_mensal  # For plotting
            })
            
        except Exception as e:
            print(f"Error processing {ativo}: {e}")
            continue

    if resultados:
        resultados_df = pd.DataFrame(resultados)
        
        # --- PLOT CONFIGURATION ---
        plt.style.use('seaborn-v0_8') 
        fig, ax = plt.subplots(figsize=(12, 8))
        
        # Data for plot
        ativos = resultados_df['Ativo']
        dividendos = resultados_df['Valor Numerico']
        meta = 25.0
        percentuais = (dividendos / meta) * 100
        
        # Create pie chart
        wedges, texts, autotexts = ax.pie(
            dividendos,
            labels=[f"{ativo}\nR$ {div:.2f}\n({perc:.1f}% da meta)" 
                   for ativo, div, perc in zip(ativos, dividendos, percentuais)],
            autopct=lambda p: f"{p:.1f}%",
            startangle=90,
            explode=[0.05] * len(ativos),
            colors=['#FF9AA2', '#FFB7B2', '#FFDAC1', '#E2F0CB'],
            textprops={'fontsize': 10},
            pctdistance=0.85
        )
        
        # Style adjustments
        plt.setp(autotexts, size=10, weight="bold")
        ax.set_title(
            'Progresso em Relação à Meta de Dividendos Mensais\n(Meta: R$ 25,00 por ativo)',
            pad=20, fontsize=14, fontweight='bold'
        )
        
        # Add center circle for doughnut effect
        centre_circle = plt.Circle((0,0), 0.70, fc='white')
        fig.gca().add_artist(centre_circle)
        
        # Add total value in center
        total = dividendos.sum()
        ax.text(
            0, 0, 
            f"Total Mensal:\nR$ {total:.2f}",
            ha='center', va='center', 
            fontsize=12, fontweight='bold'
        )
        
        # Equal aspect ratio ensures pie is drawn as circle
        ax.axis('equal')  
        
        # Save and show
        plt.tight_layout()
        plt.savefig('progresso_dividendos.png', dpi=300, bbox_inches='tight')
        print("\n✅ Gráfico salvo como 'progresso_dividendos.png'")
        
        # Show results table
        print("\n💰 RESUMO DE DIVIDENDOS POR ATIVO 💰")
        print(resultados_df.drop(columns=['Valor Numerico']).to_string(index=False))
        
    else:
        print("No valid assets found for dividend calculation.")

# Usage
calcular_dividendos("Fundos_imobiliarios.ods")
