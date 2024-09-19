Juntar Planilhas Excel
Este script Python é usado para juntar diversas planilhas Excel em um único arquivo Excel. Ele processa arquivos em uma pasta específica, adiciona uma barra de progresso para mostrar o andamento e salva o resultado em um novo arquivo Excel.

Requisitos
Python 3.x

Bibliotecas: pandas, tqdm

Para instalar as bibliotecas necessárias, execute:

bash
Copiar código
pip install pandas tqdm
Como Usar
Prepare os Arquivos: Coloque todas as planilhas Excel que você deseja unir em uma pasta.

Execute o Script: Abra o terminal ou prompt de comando e navegue até o diretório onde o script está localizado. Execute o script com o comando:

bash
Copiar código
python nome_do_script.py
Forneça as Informações:

Nome do novo arquivo: O script solicitará o nome para o arquivo de saída. Certifique-se de incluir a extensão .xlsx (ou o script a adicionará automaticamente se não for fornecida).
Local das Planilhas: Informe o caminho da pasta onde as planilhas estão localizadas.
Extensão dos arquivos Excel: Informe a extensão dos arquivos que deseja processar (por exemplo, .xlsx ou .xls).
Exemplo de Uso
bash
Copiar código
python juntar_planilhas.py
Quando solicitado:

java
Copiar código
Qual será o nome do novo arquivo? relatorio_final
Local das Planilhas: C:\Users\SeuUsuario\Documents\Planilhas
Extensão dos arquivos Excel (ex: .xlsx, .xls): .xlsx
O Que o Script Faz
Solicita o nome do arquivo de saída e adiciona a extensão .xlsx se não estiver presente.
Solicita o local da pasta que contém as planilhas Excel e a extensão dos arquivos que deseja processar.
Lê todos os arquivos Excel com a extensão especificada na pasta fornecida.
Usa uma barra de progresso para mostrar o andamento da junção dos arquivos.
Concatena todas as planilhas em um único DataFrame e salva no arquivo especificado.
Saída
Um arquivo Excel com o nome fornecido que contém a combinação de todas as planilhas da pasta.
