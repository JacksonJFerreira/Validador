import customtkinter as ctk
from tkinter import filedialog, ttk, messagebox
import openpyxl
from openpyxl.styles import Font, Border, Side
import os
from datetime import datetime
import logging

class ValidadorDeDados:
    def __init__(self):
        self.setup_logging()
        self.setup_gui()

    def setup_logging(self):
        logging.basicConfig(filename='validador.log', level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')

    def setup_gui(self):
        ctk.set_appearance_mode("Sytsem")
        ctk.set_default_color_theme("blue")

        self.janela = ctk.CTk()
        self.janela.title("Validação Nominal de Empregados")
        self.janela.geometry("500x350")

        self.diretorio_sifac = ctk.StringVar()
        self.diretorio_sispat = ctk.StringVar()
        self.mensagem_status = ctk.StringVar()

        self.criar_widgets()

    def criar_widgets(self):
        rotulo_titulo = ctk.CTkLabel(
            self.janela, text="Selecionar Pastas de Arquivos Excel", font=("Aptos", 14))
        rotulo_titulo.pack(pady=10)

        entrada_sifac = ctk.CTkEntry(self.janela, textvariable=self.diretorio_sifac, width=400)
        entrada_sifac.pack(pady=5)
        botao_selecionar_sifac = ctk.CTkButton(
            self.janela, text="Selecionar Pasta SIFAC", command=self.selecionar_pasta_sifac)
        botao_selecionar_sifac.pack(pady=5)

        entrada_sispat = ctk.CTkEntry(self.janela, textvariable=self.diretorio_sispat, width=400)
        entrada_sispat.pack(pady=5)
        botao_selecionar_sispat = ctk.CTkButton(
            self.janela, text="Selecionar Pasta SISPAT", command=self.selecionar_pasta_sispat)
        botao_selecionar_sispat.pack(pady=5)

        botao_validar = ctk.CTkButton(self.janela, text="Validar Dados", command=self.validar_dados)
        botao_validar.pack(pady=10)

        rotulo_status = ctk.CTkLabel(
            self.janela, textvariable=self.mensagem_status, font=("Aptos", 12))
        rotulo_status.pack(pady=10)

        self.progresso = ttk.Progressbar(
            self.janela, orient="horizontal", length=400, mode="determinate")
        self.progresso.pack(pady=10)

    def selecionar_pasta_sifac(self):
        self.diretorio_sifac.set(filedialog.askdirectory())

    def selecionar_pasta_sispat(self):
        self.diretorio_sispat.set(filedialog.askdirectory())

    def obter_competencia_sispat(self, caminho):
        try:
            wb = openpyxl.load_workbook(caminho, data_only=True)
            sheet = wb.active
            competencia = sheet['G1'].value
            logging.info(f"Competência SISPAT ({caminho}): {competencia}")
            return competencia
        except Exception as e:
            logging.error(f"Erro ao ler competência SISPAT ({caminho}): {e}")
            return None

    def obter_competencia_sifac(self, caminho):
        try:
            wb = openpyxl.load_workbook(caminho, data_only=True)
            sheet = wb.active
            competencia = sheet['F5'].value
            if competencia is None:
                competencia = ''
            if isinstance(competencia, str):
                if "competencia:" in competencia.lower():
                    competencia = competencia.split(":")[-1].strip()
                else:
                    competencia = competencia.strip()
            logging.info(f"Competência SIFAC ({caminho}): {competencia}")
            return competencia
        except Exception as e:
            logging.error(f"Erro ao ler competência SIFAC ({caminho}): {e}")
            return None

    def formatar_competencia(self, competencia):
        if competencia is None:
            return None
        if isinstance(competencia, datetime):
            return competencia.strftime("%Y-%m")
        elif isinstance(competencia, str):
            competencia = competencia.strip()
            if '/' in competencia:
                mes, ano = competencia.split('/')
                return f"{ano}-{mes.zfill(2)}"
            try:
                return datetime.strptime(competencia, "%Y-%m-%d").strftime("%Y-%m")
            except ValueError:
                pass
            try:
                return datetime.strptime(competencia, "%Y/%m/%d").strftime("%Y-%m")
            except ValueError:
                pass
            if len(competencia) >= 6:
                return f"{competencia[:4]}-{competencia[4:6]}"
        return str(competencia).strip()

    def validar_competencia(self):
        try:
            arquivo_sifac = next(f for f in os.listdir(self.diretorio_sifac.get()) if f.endswith(".xlsx"))
            caminho_arquivo_sifac = os.path.join(self.diretorio_sifac.get(), arquivo_sifac)
            competencia_sifac = self.formatar_competencia(self.obter_competencia_sifac(caminho_arquivo_sifac))

            arquivo_sispat = next(f for f in os.listdir(self.diretorio_sispat.get()) if f.endswith(".xlsx"))
            caminho_arquivo_sispat = os.path.join(self.diretorio_sispat.get(), arquivo_sispat)
            competencia_sispat = self.formatar_competencia(self.obter_competencia_sispat(caminho_arquivo_sispat))

            if competencia_sifac != competencia_sispat:
                messagebox.showerror("Erro de Competência", "As competências não são iguais nos arquivos das pastas SIFAC e SISPAT.")
                return False
            return True
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao verificar competências: {e}")
            return False

    def validar_dados(self):
        if not self.validar_competencia():
            return

        try:
            arquivos_sifac = [f for f in os.listdir(self.diretorio_sifac.get()) if f.endswith(".xlsx")]
            arquivos_sispat = [f for f in os.listdir(self.diretorio_sispat.get()) if f.endswith(".xlsx")]

            empregados_sispat = {}
            for arquivo in arquivos_sispat:
                caminho_arquivo = os.path.join(self.diretorio_sispat.get(), arquivo)
                wb = openpyxl.load_workbook(caminho_arquivo)
                ws = wb["Relatorio_Empregado"]
                for row in ws.iter_rows(min_row=5, min_col=2, max_col=2, values_only=True):
                    empregado = row[0]
                    if empregado:
                        empregados_sispat[empregado] = arquivo

            total_arquivos = len(arquivos_sifac)
            self.progresso["maximum"] = total_arquivos

            for i, arquivo in enumerate(arquivos_sifac):
                caminho_arquivo = os.path.join(self.diretorio_sifac.get(), arquivo)
                wb = openpyxl.load_workbook(caminho_arquivo)
                ws_origem = wb["Relatorio_Empregado"]

                if "Dados_Validados" in wb.sheetnames:
                    ws_destino = wb["Dados_Validados"]
                    wb.remove(ws_destino)
                ws_destino = wb.create_sheet("Dados_Validados")

                for row in ws_origem.iter_rows(min_row=1, max_row=1, values_only=True):
                    ws_destino.append(row)

                for row in ws_origem.iter_rows(min_row=4, values_only=True):
                    ws_destino.append(row)

                # Encontrar a última coluna usada
                ultima_coluna = ws_destino.max_column

                # Adicionar colunas STATUS, DATA E HORA, NOME DO ARQUIVO VALIDADO no cabeçalho
                ws_destino.cell(row=5, column=ultima_coluna + 1, value="STATUS")
                ws_destino.cell(row=5, column=ultima_coluna + 2, value="DATA E HORA")
                ws_destino.cell(row=5, column=ultima_coluna + 3, value="NOME DO ARQUIVO VALIDADO")

                # Definir estilo de fonte em negrito para cabeçalhos
                for col in range(1, ultima_coluna + 4):  # +4 por causa das novas colunas
                    ws_destino.cell(row=5, column=col).font = Font(bold=True)

                # Atualizar as referências das colunas para o preenchimento dos dados
                coluna_status = ultima_coluna + 1
                coluna_data_hora = ultima_coluna + 2
                coluna_arquivo_validado = ultima_coluna + 3

                # Definir o estilo de borda
                thin_border = Border(left=Side(style='thin'),
                                     right=Side(style='thin'),
                                     top=Side(style='thin'),
                                     bottom=Side(style='thin'))

                # Adicionar dados nas novas colunas e aplicar bordas
                for idx, row in enumerate(ws_destino.iter_rows(min_row=6), start=6):
                    empregado = row[3].value
                    if empregado in empregados_sispat:
                        status = "Encontrado"
                        nome_arquivo_validado = empregados_sispat[empregado]
                    else:
                        status = "Não Encontrado"
                        nome_arquivo_validado = ""

                    ws_destino.cell(row=idx, column=coluna_status, value=status)
                    ws_destino.cell(row=idx, column=coluna_data_hora, value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                    ws_destino.cell(row=idx, column=coluna_arquivo_validado, value=nome_arquivo_validado)

                    # Aplicar bordas nas novas células
                    for col in [coluna_status, coluna_data_hora, coluna_arquivo_validado]:
                        ws_destino.cell(row=idx, column=col).border = thin_border

                # Aplicar bordas nas células do cabeçalho
                for cell in ws_destino.iter_rows(min_row=5, max_row=5):
                    for c in cell:
                        c.border = thin_border

                for row in ws_destino.iter_rows(min_row=1):
                    for cell in row:
                        cell.font = Font(name="Aptos Narrow", size=10)
                        # Aplicar bordas em todas as células da linha 1
                        cell.border = thin_border

                # Remover colunas em branco
                self.remover_colunas_em_branco(ws_destino)

                # Ajustar largura das colunas automaticamente
                self.ajustar_largura_colunas(ws_destino)

                wb.save(caminho_arquivo)

                self.progresso["value"] = i + 1
                self.janela.update_idletasks()

            self.mensagem_status.set("Processamento concluído!")
        except Exception as e:
            self.mensagem_status.set(f"Erro ao processar: {e}")
            logging.error(f"Erro ao processar: {e}")

    def remover_colunas_em_branco(self, worksheet):
        """Remove colunas em branco da planilha."""
        colunas_a_remover = []
        for col in range(1, worksheet.max_column + 1):
            # Verifica se a coluna está vazia
            if all(worksheet.cell(row=row, column=col).value is None for row in range(1, worksheet.max_row + 1)):
                colunas_a_remover.append(col)

        # Remove as colunas, começando da última coluna para evitar alteração dos índices
        for col in reversed(colunas_a_remover):
            worksheet.delete_cols(col)

    def ajustar_largura_colunas(self, worksheet):
        """Ajusta a largura das colunas para que todo o texto seja visível."""
        for col in worksheet.columns:
            max_length = 0
            column = col[0].column_letter  # Obter a letra da coluna
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)  # Espaço extra
            worksheet.column_dimensions[column].width = adjusted_width

    def run(self):
        self.janela.mainloop()

if __name__ == "__main__":
    app = ValidadorDeDados()
    app.run()
