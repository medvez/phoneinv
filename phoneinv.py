import argparse
import ipaddress as ip
import os.path
import pandas
import requests
import warnings
from bs4 import BeautifulSoup


# Save result inventory file in the folder with program files
BASE_DIR = os.path.dirname(os.path.realpath(__file__))

warnings.simplefilter(action='ignore')


class CiscoPhoneInventory:
    """
    Main class, that combines the rest functions. It gets input data from CLI,
    queries every single IP address from the network and puts MAC and serials into
    resulting dict.
    """
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
    """
    This class requests single host IP address, parse web-portal and extracts MAC and serial.
    This data is returnes as tuple.
    """
    def __init__(self, phone_ip_address: str) -> None:
        self.phone_url = 'https://' + phone_ip_address + '/'
        self.soup = None
        self.mac = ''
        self.serial = ''

    def get_http_tree(self) -> None:
        http_response = requests.get(self.phone_url, verify=False, timeout=2)
        self.soup = BeautifulSoup(http_response.text, features='lxml')

    def extract_phone_parameters(self) -> None:
        _data_table = self.soup.body.table.find_all('div', {'align': 'center'})
        _table_rows = _data_table[0].find_all('tr')
        for tr in _table_rows:
            name, _, value = [td.text.strip() for td in tr.children]
            if name.startswith('MAC'):
                self.mac = value
            elif name.startswith('Serial'):
                self.serial = value

    def run(self) -> (str, str):
        self.get_http_tree()
        self.extract_phone_parameters()
        return self.mac, self.serial


class DictToExcelConverter:
    """
    This is for converting result dict into excel file.
    """
    def __init__(self, input_dict: dict, output_excel='serials_and_macs.xlsx'):
        """
        :param input_dict: dict, where all MACs and serials are stored
        :param output_excel: the name of excel file, where you want to see MAC-serial table
        """
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
