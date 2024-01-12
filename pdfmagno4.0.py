import pandas as pd
import win32com.client as win32
import os
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import time
import pythoncom
import pymsgbox
import psutil

MAX_CONCURRENT_FILES = 1
CONVERSION_TIMEOUT = 30  # Tempo limite para conversão de um arquivo (em segundos)

MASTER_FILE_PATH = r"C:\Arquivos Lopes\CONTROLE DE VENDAS\CONTROLE DE VENDAS.xlsm"

def convert_file(input_file, output_file, master_workbook):
    pythoncom.CoInitialize()  # Inicializa o COM para evitar o erro "CoInitialize não foi chamado"

    input_path = input_file
    filename = os.path.splitext(os.path.basename(input_file))[0]

    # Renomear o arquivo PDF antes da conversão
    current_date = datetime.now().strftime('%d-%m-%Y')
    output_filename = f"{filename}_{current_date}.pdf"
    output_file = os.path.join(os.path.dirname(output_file), output_filename)

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.DisplayAlerts = False  # Desativa a exibição de alertas (como salvar alterações)

    workbook = excel.Workbooks.Open(input_path, False, True)  # Abrir em modo somente leitura, sem exibir

    try:
        start_time = time.time()  # Tempo de início da conversão
        while True:
            try:
                workbook.ExportAsFixedFormat(0, output_file)  # Exportar para PDF
                print("Arquivo convertido com sucesso para PDF:", output_file)
                return True
            except Exception as e:
                if time.time() - start_time >= CONVERSION_TIMEOUT:
                    print("Tempo limite excedido. Pulando a conversão do arquivo:", input_file)
                    return False
                else:
                    print("Erro ao converter o arquivo para PDF:", str(e))
                    time.sleep(1)

    finally:
        workbook.Close(False)  # Fechar sem salvar as alterações no arquivo XLSX
        excel.DisplayAlerts = True  # Restaura as configurações originais de exibição de alertas

def convert_xlsx_to_pdf(input_dirs, output_dirs):
    total_converted = 0  # Variável para contar o total de arquivos convertidos

    with ThreadPoolExecutor() as executor:
        futures = []

        # Abra a planilha mãe
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False  # Não exibir o Excel durante a atualização dos arquivos
        master_workbook = excel.Workbooks.Open(MASTER_FILE_PATH, False, True)  # Abrir em modo somente leitura, sem exibir

        try:
            for input_dir, output_dir in zip(input_dirs, output_dirs):
                files = [file for file in os.listdir(input_dir) if file.endswith(".xlsx")]

                for file in files:
                    input_path = os.path.join(input_dir, file)
                    filename = os.path.splitext(file)[0]
                    output_path = os.path.join(output_dir, f"{filename}.pdf")

                    if os.path.exists(input_path):  # Verifica se o arquivo de entrada existe
                        future = executor.submit(convert_file, input_path, output_path, master_workbook)
                        futures.append(future)

                        if len(futures) >= MAX_CONCURRENT_FILES:
                            # Aguardar até que pelo menos um dos processos seja concluído
                            completed = list(as_completed(futures))
                            for future in completed:
                                if future.result():
                                    total_converted += 1
                                futures.remove(future)

            # Aguardar a conclusão dos processos restantes
            for future in as_completed(futures):
                if future.result():
                    total_converted += 1

        finally:
            # Feche a planilha mãe e o Excel
            master_workbook.Close(False)
            excel.Quit()

            # Finalizar o processo do Excel
            for proc in psutil.process_iter():
                if proc.name() == "EXCEL.EXE":
                    proc.kill()

    return total_converted

def main():
    input_dirs = [
        r"C:\Arquivos Lopes\CONTROLE DE VENDAS\ENVIAR POR EMAIL\PDF\EXCEL"
    ]

    output_dirs = [
        r"C:\Arquivos Lopes\CONTROLE DE VENDAS\ENVIAR POR EMAIL\PDF\PDF"
    ]

    total_converted = convert_xlsx_to_pdf(input_dirs, output_dirs)

    popup_message = f"Olá, Analista. O total de arquivos convertidos foi: {total_converted}. Gratidão para a equipe de Desenvolvimento!!!^.^."
    pymsgbox.alert(popup_message, "Conversão Concluída")

if __name__ == "__main__":
    main()




