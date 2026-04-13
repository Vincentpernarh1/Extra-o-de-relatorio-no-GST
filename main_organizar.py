import pandas as pd
import os
import sys
from tkinter import messagebox

BASE_DIR = os.getcwd()  # Get the current working directory
DOWNLOAD_DIR = os.path.join(BASE_DIR, "Downloads_Auxiliar")
CONSOLIDATED_DIR = os.path.join(BASE_DIR, "Arquivos_Consolidados")

# Check if called with --skip-confirmation argument (when called from main.py)
skip_confirmation = '--skip-confirmation' in sys.argv

if skip_confirmation:
    resposta = 'yes'
else:
    resposta = messagebox.askquestion("Confirmação", "Deseja prosseguir com o tratamento?")

if resposta == 'yes':   
    try:
        # Debug info
        print(f"BASE_DIR: {BASE_DIR}")
        print(f"DOWNLOAD_DIR: {DOWNLOAD_DIR}")
        print(f"CONSOLIDATED_DIR: {CONSOLIDATED_DIR}")
        
        # Ensure consolidated directory exists
        os.makedirs(CONSOLIDATED_DIR, exist_ok=True)
        
        # Check if there are any Excel files to process
        if not os.path.exists(DOWNLOAD_DIR):
            raise FileNotFoundError(f"Pasta {DOWNLOAD_DIR} não encontrada!")
        
        excel_files = [f for f in os.listdir(DOWNLOAD_DIR) if f.endswith('.xlsx')]
        print(f"Arquivos encontrados: {excel_files}")
        
        if len(excel_files) == 0:
            raise FileNotFoundError("Nenhum arquivo Excel encontrado em Downloads_Auxiliar!")
        
        print(f"Processando {len(excel_files)} arquivo(s) Excel...")
                
        def ler_excel_e_concatenar(caminho_pasta, x, y):
            # Lista todos os arquivos na pasta especificada com a extensão .xlsx
            arquivos_excel = [arquivo for arquivo in os.listdir(caminho_pasta) if arquivo.endswith('.xlsx')]

            # Ordena os arquivos por data de modificação (do mais recente para o mais antigo)
            arquivos_excel = sorted(arquivos_excel, key=lambda arquivo: os.path.getmtime(os.path.join(caminho_pasta, arquivo)))

            # Seleciona os últimos quatro arquivos
            arquivos_selecionados = arquivos_excel[x:y]

            # Inicializa um DataFrame vazio para armazenar os dados
            df_concatenado = pd.DataFrame()

            # Itera sobre cada arquivo CSV na pasta
            for arquivos_excel in arquivos_selecionados:
                # Constrói o caminho completo do arquivo
                caminho_arquivo = os.path.join(caminho_pasta, arquivos_excel)

                # Lê o arquivo CSV e adiciona ao DataFrame
                df_atual = pd.read_excel(caminho_arquivo)

                # Verifica se df_atual é vazio
                if df_atual.empty:
                    continue
                else:
                    df_concatenado = pd.concat([df_concatenado, df_atual], ignore_index=True)

            return df_concatenado

        # Chama a função para ler e concatenar os arquivos CSV
        df_source_package_management = ler_excel_e_concatenar(DOWNLOAD_DIR, None, 1)
        df_sourcing_management = ler_excel_e_concatenar(DOWNLOAD_DIR, 1, None)


        # Filtrando as linhas - check if columns exist
        if 'Status' in df_source_package_management.columns and 'Tipo' in df_source_package_management.columns:
            filtro = (df_source_package_management['Status'] == 'Technical Data Completed') & (df_source_package_management['Tipo'] != 'TAG')
            df_source_package_mananegement_filtred = df_source_package_management[filtro]
        elif 'Status' in df_source_package_management.columns:
            # If only Status exists
            filtro = (df_source_package_management['Status'] == 'Technical Data Completed')
            df_source_package_mananegement_filtred = df_source_package_management[filtro]
        else:
            # No filter if columns don't exist
            df_source_package_mananegement_filtred = df_source_package_management
        
        # Select columns that actually exist
        required_cols_spm = ["Source Package #", "Engineering Unit", "Descrição", "Modelo", 
                             "Version", "Tipo", "Initiator", "Nome do Comprador 1", "Status"]
        available_cols_spm = [col for col in required_cols_spm if col in df_source_package_mananegement_filtred.columns]
        
        if available_cols_spm:
            df_source_package_mananegement_filtred = df_source_package_mananegement_filtred[available_cols_spm]
        
        # Filter Sourcing Management columns
        required_cols_sm = ['Sourcing Process #', 'Modelo', 'Nome Comprador', 'Sourcing Status',
                            'LRB/LSS Status', 'Sourcing Step', 'LRB/LSS Sent to SCM On', 'Sourcing Type']
        available_cols_sm = [col for col in required_cols_sm if col in df_sourcing_management.columns]
        
        if available_cols_sm:
            df_sourcing_management = df_sourcing_management[available_cols_sm]

        # Salvando em excel
        spm_path = os.path.join(CONSOLIDATED_DIR, 'Source_Package_Management.xlsx')
        sm_path = os.path.join(CONSOLIDATED_DIR, 'Sourcing_Management.xlsx')
        
        df_source_package_mananegement_filtred.to_excel(spm_path, index=False)
        df_sourcing_management.to_excel(sm_path, index=False)



        # Importe a biblioteca pandas e crie um objeto pd.ExcelWriter para criar um arquivo Excel usando o XlsxWriter como mecanismo de escrita.
        writer = pd.ExcelWriter(spm_path, engine='xlsxwriter')

        # Salve um DataFrame chamado pivot_df na primeira aba do arquivo Excel.
        df_source_package_mananegement_filtred.to_excel(writer, sheet_name='Primeira_Aba', index=False)

        # Obtenha uma referência para o livro e a planilha no arquivo Excel.
        workbook = writer.book
        worksheet = writer.sheets['Primeira_Aba']

        left_format = workbook.add_format({'align': 'left'})
        center_format = workbook.add_format({'align': 'center'})
        right_format = workbook.add_format({'align': 'right'})

        # Defina a largura das colunas A a J.
        worksheet.set_column('A:A', 25, center_format)
        worksheet.set_column('B:B', 25)
        worksheet.set_column('C:C', 50)
        worksheet.set_column('D:D', 25, center_format)
        worksheet.set_column('E:E', 25, center_format)
        worksheet.set_column('F:F', 25, center_format)
        worksheet.set_column('G:G', 40)
        worksheet.set_column('H:H', 40, left_format)
        worksheet.set_column('I:I', 25, center_format)

        last_row = len(df_source_package_mananegement_filtred)

        # Adicione bordas internas para as células com formatação condicional.
        inner_border = workbook.add_format({"border": 4})
        conditional_range = f'A1:I{last_row}'
        worksheet.conditional_format(conditional_range, {'type': 'no_blanks', 'format': inner_border})

        # Defina um formato para o cabeçalho da tabela.
        header_format = workbook.add_format({
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#000000',
            'bold': True,
            'font_color': '#FFFFFF',
            'border': 1,
            'border_color': '#D3D3D3'
        })

        # Escreva os nomes das colunas do DataFrame com o formato de cabeçalho na linha 2.
        for col_num, value in enumerate(df_source_package_mananegement_filtred.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # Salve o arquivo Excel.
        writer.close()



        # Importe a biblioteca pandas e crie um objeto pd.ExcelWriter para criar um arquivo Excel usando o XlsxWriter como mecanismo de escrita.
        writer = pd.ExcelWriter(sm_path, engine='xlsxwriter')

        # Salve um DataFrame chamado pivot_df na primeira aba do arquivo Excel.
        df_sourcing_management.to_excel(writer, sheet_name='Primeira_Aba', index=False)

        # Obtenha uma referência para o livro e a planilha no arquivo Excel.
        workbook = writer.book
        worksheet = writer.sheets['Primeira_Aba']

        left_format = workbook.add_format({'align': 'left'})
        center_format = workbook.add_format({'align': 'center'})
        right_format = workbook.add_format({'align': 'right'})

        # Defina a largura das colunas A a J.
        worksheet.set_column('A:A', 25, center_format)
        worksheet.set_column('B:B', 25)
        worksheet.set_column('C:C', 40, left_format)
        worksheet.set_column('D:D', 25)
        worksheet.set_column('E:E', 25)
        worksheet.set_column('F:F', 25)
        worksheet.set_column('G:G', 25)
        worksheet.set_column('H:H', 25, center_format)

        last_row = len(df_sourcing_management)

        # Adicione bordas internas para as células com formatação condicional.
        inner_border = workbook.add_format({"border": 4})
        conditional_range = f'A1:H{last_row}'
        worksheet.conditional_format(conditional_range, {'type': 'no_blanks', 'format': inner_border})

        # Defina um formato para o cabeçalho da tabela.
        header_format = workbook.add_format({
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#000000',
            'bold': True,
            'font_color': '#FFFFFF',
            'border': 1,
            'border_color': '#D3D3D3'
        })

        # Escreva os nomes das colunas do DataFrame com o formato de cabeçalho na linha 2.
        for col_num, value in enumerate(df_sourcing_management.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # Salve o arquivo Excel.
        writer.close()

        messagebox.showinfo("Sucesso", "Tratamento realizado com sucesso!")

    except Exception as exc:
        print(f"Erro detalhado: {exc}")
        messagebox.showerror("Erro", f"Erro ao executar o tratamento: {exc}")
        messagebox.showinfo("Aviso", "Tente novamente! Caso o erro persista, entre em contato com o suporte.")

else:
    # Coloque aqui o código que deseja executar se o usuário cancelar
    messagebox.showwarning("Cancelado", "Operação cancelada!")