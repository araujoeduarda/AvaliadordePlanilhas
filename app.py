import pandas as pd
import matplotlib.pyplot as plt

# Ler arquivo Excel
def read_excel_file(file_path):
    try:
        df = pd.read_excel(file_path)
        print(f"Arquivo '{file_path}' lido com sucesso.")
        return df
    except Exception as e:
        print(f"Erro ao ler o arquivo: {e}")
        return None

# Processar os dados
def process_data(df):
    if 'Valor' not in df.columns:
        print('Coluna "Valor" não encontrada.')
        return None, None

    media_valor = df['Valor'].mean()
    resultado = [f'Média da coluna "Valor": {media_valor:.2f}']

    df_ativo = df[df['Status'] == 'Ativo']
    soma_ativo = df_ativo['Valor'].sum()
    resultado.append(f'Soma dos valores "Ativo": {soma_ativo:.2f}')
    count_ativo = df_ativo.shape[0]
    resultado.append(f'Contagem de itens "Ativo": {count_ativo}')
    maior_valor = df['Valor'].max()
    menor_valor = df['Valor'].min()
    resultado.append(f'Maior valor: {maior_valor}, Menor valor: {menor_valor}')
    mediana_valor = df['Valor'].median()
    moda_valor = df['Valor'].mode()[0]
    desvio_padrao_valor = df['Valor'].std()
    resultado.append(f'Mediana: {mediana_valor:.2f}, Moda: {moda_valor:.2f}, Desvio Padrão: {desvio_padrao_valor:.2f}')
    resultado.append(f'Total recebido ativamente: {soma_ativo:.2f}')

    return df_ativo, resultado

# Salva os dados processados em um novo arquivo Excel
def save_to_excel(df, output_path):
    df.to_excel(output_path, index=False)
    print(f'Dados salvos em {output_path}')

# Salva o resumo em um arquivo CSV
def save_summary_to_csv(summary, output_path):
    summary_df = pd.DataFrame({'Resultado': summary})
    summary_df.to_csv(output_path, index=False)
    print(f'Resumo salvo em {output_path}')

# Gera gráfico dos valores ativos
def plot_values(df):
    plt.figure(figsize=(10, 6))
    plt.bar(df['ID'], df['Valor'], color='green', alpha=0.7)
    plt.title('Valores Ativos')
    plt.xlabel('ID')
    plt.ylabel('Valor')
    plt.xticks(df['ID'])
    plt.grid(axis='y')
    plt.savefig('valores_ativos.png', format='png')  # Salva o gráfico como imagem
    plt.show()  # Exibe o gráfico

# Gera gráfico de pizza para status
def plot_status_distribution(df):
    status_count = df['Status'].value_counts()
    plt.figure(figsize=(8, 8))
    plt.pie(status_count, labels=status_count.index, autopct='%1.1f%%', startangle=140, colors=['lightblue', 'lightgreen'])
    plt.title('Distribuição de Status')
    plt.savefig('distribuicao_status.png', format='png')  # Salva o gráfico de pizza
    plt.show()  # Exibe o gráfico de pizza

# Resumo dos dados
def summarize_data(df):
    print("\nResumo dos Dados:")
    resumo = df.describe(include='all')
    print(resumo)
    return resumo

# Caminho para o arquivo de entrada e saída
input_file = 'dados_entrada.xlsx'  # O nome do arquivo Excel
output_file = 'dados_processados.xlsx'  # O nome do arquivo de saída
summary_file = 'resumo_dados.csv'  # Novo arquivo CSV para o resumo

# Execução do script
if __name__ == "__main__":
    dados = read_excel_file(input_file)
    
    if dados is not None:
        dados_processados, resultado = process_data(dados)
        
        if dados_processados is not None:
            save_to_excel(dados_processados, output_file)
            save_summary_to_csv(resultado, summary_file)  # Salva o resumo em CSV
            plot_values(dados_processados)  # Plota os valores dos dados processados
            plot_status_distribution(dados)  # Plota a distribuição de status
            resumo = summarize_data(dados)  # Mostra um resumo dos dados
