import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import serial
import time


class RegistroJornada:
    def __init__(self, credenciais_json, spreadsheet_id, nome_planilha_funcionarios, porta_serial):
        self.credenciais_json = credenciais_json
        self.spreadsheet_id = spreadsheet_id
        self.nome_planilha_funcionarios = nome_planilha_funcionarios
        self.porta_serial = porta_serial
        self.planilha_ponto = None
        self.planilha_funcionarios = None
        self.sheet_ponto = None
        self.sheet_funcionarios = None
        self.ultima_acao = {}

        # Configuração da comunicação serial
        self.configurar_serial()

        # Autenticação e acesso às planilhas
        self.conectar_google_sheets()

    def configurar_serial(self):
        # Configuração da porta serial para comunicação com o Arduino
        self.ser = serial.Serial(self.porta_serial, 9600, timeout=1)

    def conectar_google_sheets(self):
        escopo = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
                  "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
        credenciais = ServiceAccountCredentials.from_json_keyfile_name(self.credenciais_json, escopo)
        cliente = gspread.authorize(credenciais)

        # Acesso à planilha pelo ID
        self.planilha_ponto = cliente.open_by_key(self.spreadsheet_id)
        self.sheet_ponto = self.planilha_ponto.sheet1  # Primeira aba da planilha de ponto

        # Acessa a aba de funcionários
        self.planilha_funcionarios = cliente.open_by_key(self.spreadsheet_id)
        self.sheet_funcionarios = self.planilha_funcionarios.worksheet(self.nome_planilha_funcionarios)

    def buscar_nome_funcionario(self, uid):
        funcionarios = self.sheet_funcionarios.get_all_records()
        for funcionario in funcionarios:
            if funcionario['UID'] == uid:
                return funcionario['Nome']
        return None

    def registrar_ponto(self, uid):
        nome_funcionario = self.buscar_nome_funcionario(uid)
        if not nome_funcionario:
            print("UID não encontrado. Por favor, cadastre o cartão.")
            return

        data_atual = datetime.now().strftime('%Y-%m-%d')
        hora_atual = datetime.now().strftime('%H:%M:%S')

        # Localiza a linha do colaborador para atualizar o ponto
        registros = self.sheet_ponto.get_all_records()
        linha_atualizar = None

        for idx, registro in enumerate(registros, start=2):
            if registro['Data'] == data_atual and registro['UID'] == uid:
                linha_atualizar = idx
                break

        # Define o próximo evento de ponto baseado nas ações anteriores
        if linha_atualizar:
            registro_atual = registros[linha_atualizar - 2]
            if registro_atual.get('Entrada') == "":
                coluna = 'Entrada'
                mensagem = "ENTRADA REGISTRADA COM SUCESSO"
            elif registro_atual.get('Saída para Almoço') == "":
                coluna = 'Saída para Almoço'
                mensagem = "SAÍDA PARA ALMOÇO REGISTRADA COM SUCESSO"
            elif registro_atual.get('Volta do Almoço') == "":
                coluna = 'Volta do Almoço'
                mensagem = "VOLTA DO ALMOÇO REGISTRADA COM SUCESSO"
            elif registro_atual.get('Fim do Expediente') == "":
                coluna = 'Fim do Expediente'
                mensagem = "FIM DO EXPEDIENTE REGISTRADO COM SUCESSO"
            else:
                print("TODOS OS PONTOS JÁ FORAM REGISTRADOS PARA HOJE.")
                time.sleep(5)
                return
        else:
            # Se não há linha, cria uma nova linha com Entrada
            linha_atualizar = len(registros) + 2
            self.sheet_ponto.update(values=[[data_atual, nome_funcionario, uid, hora_atual, ""]],
                                    range_name=f"A{linha_atualizar}:E{linha_atualizar}")
            print("ENTRADA REGISTRADA COM SUCESSO")
            time.sleep(3)
            return

        # Atualiza o próximo evento de ponto na linha do funcionário
        colunas = {"Entrada": "D", "Saída para Almoço": "E", "Volta do Almoço": "F", "Fim do Expediente": "G"}
        self.sheet_ponto.update(values=[[hora_atual]], range_name=f"{colunas[coluna]}{linha_atualizar}")
        print(mensagem)
        time.sleep(3)

    def iniciar_monitoramento(self):
        print("APROXIME SEU CARTÃO NO LEITOR")
        while True:
            uid = self.ler_uid_arduino()
            if uid:
                self.registrar_ponto(uid)
                print("APROXIME SEU CARTÃO NO LEITOR")

    def ler_uid_arduino(self):
        if self.ser.in_waiting > 0:
            linha = self.ser.readline().decode('utf-8').strip()
            print(f"UID lido: {linha}")
            return linha
        return None


# Exemplo de uso
if __name__ == "__main__":
    registro = RegistroJornada(
        credenciais_json="C:/Users/Lucas Dias/Desktop/PROJETO IOT/controle_jornada/credenciais.json",
        spreadsheet_id="1wfVGtf0N7j6Fr0GLLsgm0yVUB19gRwrb7t6zrXDzx-s",
        nome_planilha_funcionarios="Funcionarios",
        porta_serial="COM3"
    )
    registro.iniciar_monitoramento()
