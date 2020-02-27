import sys
import ssl
import socket
import time
import OpenSSL
from lxml import etree, objectify
import xml.etree.cElementTree as ET

"""
hostname = "vk.com"
ctx = ssl.create_default_context()
s = ctx.wrap_socket(socket.socket(), server_hostname=hostname)
s.connect((hostname, 443))
cert = s.getpeercert()

print(cert)
print("match_hostname: ", ssl.match_hostname(cert, hostname))
print("cert_time_to_seconds: ", str(int(ssl.cert_time_to_seconds(cert['notAfter'])) - time.time()))
"""


# Получение сертификата с хоста по порту 443
def get_cert(hostname):
    certificate = ssl.get_server_certificate((hostname, 443))
    print("get_server_certificate: \n", certificate)

    decode_data = OpenSSL.crypto.load_certificate(OpenSSL.crypto.FILETYPE_PEM, certificate)

    print("issuer: ", decode_data.get_issuer())
    print("notAfter: ", time.mktime(time.strptime(decode_data.get_notAfter().decode('ascii'), '%Y%m%d%H%M%SZ')))
    print("notBefore: ", time.mktime(time.strptime(decode_data.get_notBefore().decode('ascii'), '%Y%m%d%H%M%SZ')))

    return decode_data


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


# decode_data = get_cert(hostname)


addresses = get_hostname("report.xml")

root = ET.Element("root")

# Стучимся к хостам, проверяем сертификаты и доступность
for addr in addresses:
    print("\n\naddress: ", addr)
    try:
        decode_data = get_cert(addr)
        try:
            ET.SubElement(root, "host", addr=addr, status="available", cert_readable="yes",
                          notAfter=str(
                              time.mktime(time.strptime(decode_data.get_notAfter().decode('ascii'), '%Y%m%d%H%M%SZ'))),
                          notBefore=str(
                              time.mktime(time.strptime(decode_data.get_notBefore().decode('ascii'), '%Y%m%d%H%M%SZ'))))
        except Exception as Ex:
            print(Ex)
            ET.SubElement(root, "host", addr=addr, status="available", cert_readable="no")
    except Exception as Ex:
        print(Ex)
        ET.SubElement(root, "host", addr=addr, status="not_available")

# Создаем файл с отчетом
tree = ET.ElementTree(root)
tree.write("result.xml")
