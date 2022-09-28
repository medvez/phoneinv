import argparse
import ipaddress as ip
import os.path
import pandas
import requests
import warnings
from bs4 import BeautifulSoup

BASE_DIR = os.path.dirname(os.path.realpath(__file__))

warnings.simplefilter(action='ignore')


class CiscoPhoneInventory:
    def __init__(self) -> None:
        self.phone_parameters = {}
        self.phones_network = ''

    def get_input_data(self) -> None:
        parser = argparse.ArgumentParser(description='Get serials and MACs of cisco phones')
        parser.add_argument('network', type=str, help='Input cisco phones network in format 192.168.0.0/24')
        self.phones_network = parser.parse_args().network

    def query_network(self) -> None:
        for host in ip.ip_network(self.phones_network).hosts():
            host = format(host)
            try:
                mac, serial = HTTPRequester(host).run()
                self.phone_parameters[mac] = serial
                print(f"{host:<13}{' - ':^3}{'OK':>4}")
            except requests.exceptions.ConnectTimeout:
                print(f"{host:<13}{' - ':^3}{'VOID':>4}")

    def run(self) -> None:
        self.get_input_data()
        self.query_network()
        DictToExcelConverter(input_dict=self.phone_parameters).run()


class HTTPRequester:
    def __init__(self, phone_ip_address: str) -> None:
        self.phone_url = 'https://' + phone_ip_address + '/'
        self.http_tree = None
        self.mac = ''
        self.serial = ''

    def get_http_tree(self) -> None:
        http_response = requests.get(self.phone_url, verify=False, timeout=2)
        self.http_tree = BeautifulSoup(http_response.content, 'lxml')

    def extract_phone_parameters(self) -> None:
        phone_parameters_table = self.http_tree.findAll('table')[2]
        for child in phone_parameters_table.children:
            parameter_name = child.findAll('td')[0].string.lstrip()
            parameter_value = child.findAll('td')[2].string
            if parameter_name.startswith('MAC'):
                self.mac = parameter_value
            elif parameter_name.startswith('Serial'):
                self.serial = parameter_value

    def run(self) -> (str, str):
        self.get_http_tree()
        self.extract_phone_parameters()
        return self.mac, self.serial


class DictToExcelConverter:
    def __init__(self, input_dict: dict, output_excel='serials_and_macs.xlsx'):
        self.input_dict = input_dict
        self.output_excel = os.path.join(BASE_DIR, output_excel)

    def write_excel(self) -> None:
        _frame = pandas.DataFrame.from_dict(self.input_dict, orient='index', columns=['Serial'])
        _frame.index.name = 'MAC'
        _frame.to_excel(self.output_excel)

    def run(self) -> None:
        self.write_excel()


if __name__ == '__main__':
    CiscoPhoneInventory().run()
