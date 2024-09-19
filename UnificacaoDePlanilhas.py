import pandas as pd
import os
from tqdm import tqdm

def juntarPlanilhas(listarArquivos, arquivoSaida):
    if not listarArquivos:
        print("Nenhum arquivo Excel foi encontrado.")
        return
    
    # Lista para armazenar os DataFrames
    planilhas = []

    # Inicializa a barra de progresso
    print("Processando as planilhas...")
    for arquivo in tqdm(listarArquivos, desc="Carregando Planilhas", unit="arquivo"):
        try:
            df = pd.read_excel(arquivo)
            planilhas.append(df)
        except Exception as e:
            print(f"Erro ao ler o arquivo {arquivo}: {e}")

    # Verificar se há DataFrames para concatenar
    if planilhas:
        funcaoFinal = pd.concat(planilhas, ignore_index=True)
        funcaoFinal.to_excel(arquivoSaida, index=False)
        print(f"Planilhas combinadas salvas em: {arquivoSaida}")
    else:
        print("Nenhuma planilha foi adicionada para combinar.")

if __name__ == "__main__":
    # Definir os arquivos que deseja juntar
    pastaArquivos = input('Local das Planilhas: ')
    
    # Lista de arquivos Excel a serem mesclados
    extensaoExcel = input('Extensão dos arquivos Excel (ex: .xlsx, .xls): ')
    listarArquivos = [os.path.join(pastaArquivos, f) for f in os.listdir(pastaArquivos) if f.endswith(extensaoExcel)]
    
    # Nome do arquivo de saída
    nomeSaida = input("Qual será o nome do novo arquivo? ")
    if not nomeSaida.endswith(".xlsx"):
        nomeSaida += ".xlsx"

        arquivoSaida = nomeSaida
    
    # Chama a função para juntar os arquivos e salvar o resultado
    juntarPlanilhas(listarArquivos, arquivoSaida)

    print(f"Planilhas combinadas salvas em: {arquivoSaida}")
