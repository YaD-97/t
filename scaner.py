import ssl
import time
import OpenSSL
from lxml import objectify
import openpyxl
import os
import datetime


# Получение сертификата с хоста по порту 443
def get_cert(hostname):
    print("")
    certificate = ssl.get_server_certificate((hostname, 443))
    print("get_server_certificate: \n", certificate)

    decode_data = OpenSSL.crypto.load_certificate(OpenSSL.crypto.FILETYPE_PEM, certificate)

    print("issuer: ", decode_data.get_issuer())
    print("notAfter: ", time.mktime(time.strptime(decode_data.get_notAfter().decode('ascii'), '%Y%m%d%H%M%SZ')))
    print("notBefore: ", time.mktime(time.strptime(decode_data.get_notBefore().decode('ascii'), '%Y%m%d%H%M%SZ')))

    return decode_data


def days_left(time_):
    time_ = time_.decode('ascii')
    time_ = time.strptime(time_, '%Y%m%d%H%M%SZ')
    time_ = time.mktime(time_)
    if (time_ - time.time()) > 0:
        days = (time_ - time.time())//86400
    else:
        days = 0
    return days


# Форматирование даты
def format_data(time_):
    time_ = time_.decode('ascii')
    time_ = time.strptime(time_, '%Y%m%d%H%M%SZ')
    # data = str(time_[2]) + "-" + str(time_[1]) + "-" + str(time_[0])
    data = datetime.datetime(time_[0], time_[1], time_[2], time_[3], time_[4], time_[5])
    return data


# Получение списка хостов из отчета masscan
def get_hostname(file):
    addresses = []
    with open(file) as fobj:
        xml = fobj.read()

    root = objectify.fromstring(xml)

    for host in root.getchildren():
        if host.tag == "host":
            if host.ports.port.attrib["portid"] == "443":
                # print("address: ", host.address.attrib["addr"])
                addresses.append(host.address.attrib["addr"])

    return addresses


def Excel(results):
    # объект
    wb = openpyxl.Workbook()

    for result in results:
        sheet_name = result[0]
        rows = result[1]

        # активный лист
        ws = wb.active

        # название страницы
        ws = wb.create_sheet(sheet_name, 0)
        ws.title = sheet_name

        # циклом записываем данные
        for row in rows:
            ws.append(row)

        # ирина ячеек
        ws.column_dimensions["A"].width = 15
        ws.column_dimensions["B"].width = 15
        ws.column_dimensions["C"].width = 15
        ws.column_dimensions["D"].width = 22
        ws.column_dimensions["E"].width = 22
        ws.column_dimensions["F"].width = 22
        ws.column_dimensions["G"].width = 100

    # сохранение файла в текущую директорию
    wb.save("results.xlsx")


# Получаем список отчетов в папке
directory = "./reports/"
files = os.listdir(directory)
files = filter(lambda x: x.endswith('.xml'), files)
results = []


# Прокручиваем отчеты один за другим
for file in files:
    addresses = get_hostname("./reports/" + file)
    result = [["Адрес", "Доступность", "Сертификат", "Действует от", "Годен до", "Осталось дней", "Примечание"]]
    # Стучимся к хостам, проверяем сертификаты и доступность
    for addr in addresses:
        print("\n\naddress: ", addr)
        try:
            decode_data = get_cert(addr)
            try:
                result.append([addr, "Доступен", "Есть",
                               format_data(decode_data.get_notBefore()),
                               format_data(decode_data.get_notAfter()),
                               days_left(decode_data.get_notAfter()),
                               str(decode_data.get_issuer())])
            except Exception as Ex:
                print(Ex)
                result.append([addr, "Доступен", "Нет", "", "", "", str(Ex)])
        except Exception as Ex:
            print(Ex)
            result.append([addr, "Не доступен", "", "", "", "", str(Ex)])

    results.append([file, result])

Excel(results)
