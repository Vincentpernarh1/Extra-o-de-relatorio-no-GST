import pandas as pd
import os
import sys
import tkinter as tk
from tkinter import messagebox

# When built/launched as a windowed .exe, sys.stdout/stderr are None and any
# print() raises AttributeError, killing the process before anything shows.
if getattr(sys, "frozen", False) and (sys.stdout is None or sys.stderr is None):
    _log_path = os.path.join(os.path.dirname(sys.executable), "rpa_log.txt")
    _log_file = open(_log_path, "a", encoding="utf-8", buffering=1)
    sys.stdout = _log_file
    sys.stderr = _log_file


def _dialog_root():
    """Creates an invisible-but-mapped Tk root forced to the foreground so
    message boxes don't end up hidden behind the browser or other windows.
    A withdrawn (unmapped) root can make its transient dialog fail to paint
    at all, so we keep the root "shown" via 0 alpha instead of withdraw()."""
    root = tk.Tk()
    root.geometry("1x1+0+0")
    root.attributes("-alpha", 0.0)
    root.attributes("-topmost", True)
    root.deiconify()
    root.lift()
    root.update()
    return root


def ask_question(title, message):
    root = _dialog_root()
    try:
        return messagebox.askquestion(title, message, parent=root)
    finally:
        root.destroy()


def show_info(title, message):
    root = _dialog_root()
    try:
        messagebox.showinfo(title, message, parent=root)
    finally:
        root.destroy()


def show_error(title, message):
    root = _dialog_root()
    try:
        messagebox.showerror(title, message, parent=root)
    finally:
        root.destroy()


def show_warning(title, message):
    root = _dialog_root()
    try:
        messagebox.showwarning(title, message, parent=root)
    finally:
        root.destroy()


def run(base_dir=None, log=print, progress=lambda value: None):
    if base_dir is None:
        base_dir = os.getcwd()

    DOWNLOAD_DIR = os.path.join(base_dir, "Downloads_Auxiliar")
    CONSOLIDATED_DIR = os.path.join(base_dir, "Arquivos_Consolidados")

    try:
        # Ensure consolidated directory exists
        os.makedirs(CONSOLIDATED_DIR, exist_ok=True)
        progress(5)

        # Check if there are any Excel files to process
        if not os.path.exists(DOWNLOAD_DIR):
            raise FileNotFoundError(f"Pasta {DOWNLOAD_DIR} não encontrada!")

        excel_files = [f for f in os.listdir(DOWNLOAD_DIR) if f.endswith('.xlsx')]
        log(f"Arquivos encontrados: {len(excel_files)}")

        if len(excel_files) == 0:
            raise FileNotFoundError("Nenhum arquivo Excel encontrado em Downloads_Auxiliar!")

        log(f"Processando {len(excel_files)} arquivo(s) Excel...")
        progress(20)
                
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
                             "Version", "Tipo", "Initiator", "Nome do Comprador", "Status"]
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
        progress(60)



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
        progress(95)

        log("Tratamento realizado com sucesso!")
        progress(100)
        return True

    except Exception as exc:
        log(f"Erro ao executar o tratamento: {exc}")
        return False


if __name__ == "__main__":
    skip_confirmation = '--skip-confirmation' in sys.argv
    if skip_confirmation:
        success = run()
    else:
        resposta = ask_question("Confirmação", "Deseja prosseguir com o tratamento?")
        if resposta == 'yes':
            success = run()
        else:
            show_warning("Cancelado", "Operação cancelada!")
            success = None

    if success is True:
        show_info("Sucesso", "Tratamento realizado com sucesso!")
    elif success is False:
        show_error("Erro", "Erro ao executar o tratamento! Tente novamente ou entre em contato com o suporte.")