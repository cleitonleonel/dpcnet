import os
import time
import json
import requests
import pandas as pd

API_URL = 'https://apiecommerce.dpcnet.com.br'
BASE_DIR = os.getcwd()


def format_col_width(ws):
    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 10
    ws.column_dimensions['C'].width = 30
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 30
    ws.column_dimensions['F'].width = 10
    ws.column_dimensions['G'].width = 10
    ws.column_dimensions['H'].width = 10
    ws.column_dimensions['I'].width = 10


class Browser(object):

    def __init__(self):
        self.response = None
        self.headers = self.get_headers()
        self.session = requests.Session()

    def get_headers(self):
        self.headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
                          " AppleWebKit/537.36 (KHTML, like Gecko) "
                          "Chrome/87.0.4280.88 Safari/537.36"
        }
        return self.headers

    def send_request(self, method, url, **kwargs):
        response = self.session.request(method, url, **kwargs)
        if response.status_code == 200:
            return response
        return None


class DpcNetAPI(Browser):

    def __init__(self, username=None, password=None):
        self.username = username
        self.password = password
        self.current_token = None
        self.source_file = None
        self.result_json = None
        self.img_path = None
        self.log = None
        self.debug_activate = False
        super().__init__()

    def auth(self):
        data = {
            "usuario": self.username,
            "senha": self.password
        }
        self.response = self.send_request('POST',
                                          API_URL + '/ec/auth/login',
                                          params=data,
                                          headers=self.headers
                                          )
        if self.response:
            self.current_token = self.response.json()['token']

    def get_products_all(self):
        self.response = self.send_request('POST',
                                          API_URL + f'/ec/produto/showall?ec=1&token={self.current_token}',
                                          headers=self.headers
                                          )
        if self.response:
            return self.response.json()

    def get_product_by_ean(self, ean):
        data = {
            "cod_produto": [],
            "ean": [ean],
            "dun14": [],
            "searchterm": None,
            "searchType": "contain",
            "offset": 0,
            "limit": 4
        }

        self.response = self.send_request('POST',
                                          API_URL + f'/ec/produto/showall?ec=1&token={self.current_token}',
                                          json=data,
                                          headers=self.headers
                                          )
        if self.response:
            return self.response.json()

    def get_product_by_id(self, identificator):
        data = {
            "id": [identificator] if identificator else [],
            "cod_produto": [],
            "ean": [],
            "dun14": [],
            "searchterm": None,
            "searchType": "contain",
            "offset": 0,
            "limit": 4
        }

        self.response = self.send_request('POST',
                                          API_URL + f'/ec/produto/showall?ec=1&token={self.current_token}',
                                          json=data,
                                          headers=self.headers
                                          )
        if self.response:
            return self.response.json()

    def get_product_info(self, ean=None, identificator=None, integration_id=None):
        """
        data0 = {
            "count": True,
            "integracao_id": [],
            "ean": [ean],
            "dun14": [],
            "categoria_integracao_id": [],
            "fornecedor_id": None,
            "searchterm": None,
            "searchtype": "contain",
            "ordenacao": "ordem-1",
            "offset": 0,
            "limit": 24,
            "tabvnd_id": 34,
            "preco": {
                "pedido_id": 332645
            },
            "token": self.current_token,
            "columns": {
                "hascasadinha": {
                    "empresa": {"id": 32},
                    "cliente": {"id": 131601}
                },
                "hasprodbloq": {
                    "cliente": {
                        "id": 131601
                    }
                }
            }
        }

        data1 = {
            "integracao_id": "53",
            "preco": {
                "pedido_id": 333511,
                "cliente_id": 110303,
                "tabvnd_id": 33,
                "all_tabs": True
            },
            "tabvnd_id": 33,
            "token": self.current_token,
            "columns": {
                "hasprodbloq": {
                    "cliente": {
                        "id": 110303
                    }
                }
            }
        }
        """

        data = {
            "count": True,
            "id": [identificator] if identificator else [],
            "cod_produto": [],
            "integracao_id": [integration_id] if integration_id else [],
            "ean": [ean] if ean else [],
            "dun14": [],
            "categoria_integracao_id": [],
            "fornecedor_id": None,
            "searchterm": None,
            "searchtype": "contain",
            "ordenacao": "ordem-1",
            "offset": 0,
            "limit": 24,
            "tabvnd_id": 33,
            "preco": {
                "pedido_id": 333511
            },
            "token": self.current_token,
            "columns": {
                "hascasadinha": {
                    "empresa": {
                        "id": 34
                    },
                    "cliente": {
                        "id": 110303
                    }
                },
                "hasprodbloq": {
                    "cliente": {
                        "id": 110303
                    }
                }
            }
        }

        self.response = self.send_request('POST',
                                          API_URL + f'/ec/produto/showall?ec=1&token={self.current_token}',
                                          json=data,
                                          headers=self.headers
                                          )
        if self.response:
            self.result_json = self.response.json()
            return self.result_json

    def get_info_by_integration_id(self, integration_id):
        data = {
            "columns": {},
            "integracao_id": integration_id,
            "preco": {},
            "tabvnd_id": "",
            "token": ""
        }
        self.response = self.send_request('POST',
                                          API_URL + f'/ec/produto/show',
                                          json=data,
                                          headers=self.headers
                                          )
        if self.response:
            self.result_json = self.response.json()
            return self.result_json

    def download_img(self):
        url_img = self.result_json.get('produtos')[0].get('src')
        file_path = os.path.join(BASE_DIR, 'src/img')
        if not os.path.exists(file_path):
            os.makedirs(file_path, exist_ok=True)
        self.img_path = f"{file_path}/{os.path.basename(url_img)}"
        self.response = self.send_request('GET', url_img, headers=self.headers, stream=True)
        with open(self.img_path, 'wb') as file:
            for chunk in self.response.iter_content(chunk_size=1024):
                file.write(chunk)
        return self.img_path

    def export_to_excel(self, download_files=False):
        print('Extraindo dados...')
        df = pd.read_excel(self.source_file)
        cont = 0
        query = list(df.columns)[0]
        total_codes = len(list(df[query].values))
        for index, code in enumerate(list(df[query].values)):
            # self.get_product_info(ean=int(code))
            self.get_product_info(identificator=int(code))
            print(self.result_json if self.debug_activate else " ")
            df.loc[index, 'price'] = '0,00'
            df.loc[index, 'img'] = ''
            if self.result_json['count'] == 0:
                self.get_product_info(integration_id=int(code))
            if self.result_json['count'] > 0:
                df.loc[index, 'descricao'] = self.result_json.get('produtos')[0].get('descricao')
                df.loc[index, 'descricaodetalhada'] = self.result_json.get('produtos')[0].get('descricaodetalhada')
                df.loc[index, 'quantidade'] = self.result_json.get('produtos')[0].get('embqtddun14')
                df.loc[index, 'descricao da embalagem'] = self.result_json.get('produtos')[0].get('embeantext')
                df.loc[index, 'ean'] = self.result_json.get('produtos')[0].get('ean')
                df.loc[index, 'embqtdean'] = self.result_json.get('produtos')[0].get('embqtdean')
                df.loc[index, 'embtpdun14'] = self.result_json.get('produtos')[0].get('embtpdun14')
                df.loc[index, 'embeantext'] = self.result_json.get('produtos')[0].get('embeantext')
                if download_files:
                    self.download_img()
                if self.result_json and self.result_json.get('produtos')[0].get('possuiestoque'):
                    self.log = f"GRAVANDO PREÇO DO PRODUTO: {self.result_json.get('produtos')[0].get('descricao')} " \
                               f"PREÇO: {self.result_json.get('produtos')[0].get('vlrproduto'):.2f} R$"
                    # print(end=self.log)
                    df.loc[index, 'price'] = \
                        f"{self.result_json.get('produtos')[0].get('vlrproduto'):.2f}".replace('.', ',')
                    df.loc[index, 'vlrproduto'] = \
                        f"{self.result_json.get('produtos')[0].get('vlrproduto'):.2f}".replace('.', ',')
                if self.img_path:
                    df.loc[index, 'img'] = self.img_path
            cont += 1
            time.sleep(0.1)
            print(end=f'\r{self.log + " |" if self.log else ""} ITEM {cont} de {total_codes} ')
        df.to_excel('out.xlsx', index=False)
        with pd.ExcelWriter('out.xlsx') as writer:
            df.to_excel(writer, index=False)
            worksheet = writer.sheets['Sheet1']
            format_col_width(worksheet)
            writer.save()
