# arduino_communication.py

import serial
import time


def conectar_arduino(porta='COM3', baud_rate=9600):
    """
    Conecta ao Arduino via porta serial.

    Parâmetros:
    - porta: String com o nome da porta COM (exemplo: 'COM3' no Windows, '/dev/ttyUSB0' no Linux).
    - baud_rate: Taxa de transmissão em bauds (9600, por padrão).

    Retorna:
    - Objeto Serial se a conexão for bem-sucedida, None se houver erro.
    """
    try:
        # Inicializa a conexão serial com o Arduino
        arduino = serial.Serial(porta, baud_rate)
        time.sleep(2)  # Aguarda 2 segundos para estabilizar a conexão
        print("Conectado ao Arduino na porta", porta)
        return arduino
    except serial.SerialException as error:
        print(error)
        print("Erro ao conectar ao Arduino.")
        return None


def ler_uid(arduino):
    """
    Lê o UID do Arduino se houver dados disponíveis na porta serial.

    Parâmetros:
    - arduino: Objeto Serial conectado ao Arduino.

    Retorna:
    - String com o UID lido, ou None se não houver dados ou ocorrer um erro.
    """
    if arduino is not None:
        try:
            # Verifica se há dados disponíveis para leitura
            if arduino.in_waiting > 0:
                # Lê a linha de dados, decodifica para string e remove espaços
                uid = arduino.readline().decode().strip()
                print("UID recebido:", uid)
                return uid
        except Exception as e:
            print("Erro ao ler o UID:", e)
    return None

if __name__ == "__main__":
    arduino = conectar_arduino()
    if arduino is not None:
        while True:
            uid = ler_uid(arduino)
            if uid is not None:
                print("UID:", uid)
