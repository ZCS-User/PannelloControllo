import time
from pymodbus.client import ModbusTcpClient, ModbusSerialClient
import warnings

warnings.filterwarnings("ignore")

com_port = 'COM3'

client = ModbusSerialClient(
            port = com_port, baudrate = 9600, timeout = 5, stopbits = 1, bytesize = 8, parity = 'N')
client.connect()

while True:
    k = 1
    for j in [2, 3, 10]:
        str_SN = ''
        try:
            registri_read = [
                [0x0445, 1, 1],  # SN 1 + 2 digits
                [0x0446, 1, 1],  # SN 3 + 4 digits
                [0x0447, 1, 1],  # SN 5 + 6 digits
                [0x0448, 1, 1],  # SN 7 + 8 digits
                [0x0449, 1, 1],  # SN 9 + 10 digits
                [0x044a, 1, 1],  # SN 11 + 12 digits
                [0x044b, 1, 1],  # SN 13 + 14 digits
                [0x044c, 1, 1],  # SN 15 + 16 digits
                [0x0470, 1, 1],  # SN 17 + 18 digits
                [0x0471, 1, 1],  # SN 19 + 20 digits
            ]
            a = 0
            while a < 5:
                try:
                    for i in registri_read:
                        read = client.read_holding_registers(address=i[0],
                                                             count=i[1],
                                                             slave=j + 1)
                        data = read.registers[0]
                        sn1 = chr(
                            int(str(hex(int(data / i[2]))).split('x')[1][0] + str(hex(int(data / i[2]))).split('x')[1][1], base=16))
                        sn2 = chr(
                            int(str(hex(int(data / i[2]))).split('x')[1][2] + str(hex(int(data / i[2]))).split('x')[1][3], base=16))
                        str_SN += sn1 + sn2
                        k += 1
                        time.sleep(0.5)
                        a = 5
                except:
                    a += 1
            print('Address ' + str(j + 1) + ':\t' + str_SN)
        except:
            try:
                str_SN = ''
                registri_read = [
                    [0x0445, 1, 1],  # SN 1 + 2 digits
                    [0x0446, 1, 1],  # SN 3 + 4 digits
                    [0x0447, 1, 1],  # SN 5 + 6 digits
                    [0x0448, 1, 1],  # SN 7 + 8 digits
                    [0x0449, 1, 1],  # SN 9 + 10 digits
                    [0x044a, 1, 1],  # SN 11 + 12 digits
                    [0x044b, 1, 1],  # SN 13 + 14 digits
                ]
                a = 0
                while a < 5:
                    try:
                        for i in registri_read:
                            read = client.read_holding_registers(address=i[0],
                                                                 count=i[1],
                                                                 slave=j + 1)
                            data = read.registers[0]
                            sn1 = chr(
                                int(str(hex(int(data / i[2]))).split('x')[1][0] + str(hex(int(data / i[2]))).split('x')[1][1],
                                    base=16))
                            sn2 = chr(
                                int(str(hex(int(data / i[2]))).split('x')[1][2] + str(hex(int(data / i[2]))).split('x')[1][3],
                                    base=16))
                            str_SN += sn1 + sn2
                            k += 1
                            time.sleep(0.5)
                            a=5
                    except:
                        a += 1
                print('Address ' + str(j + 1) + ':\t' + str_SN)
            except:
                print('Address ' + str(j + 1) + ':\tnessuna risposta')