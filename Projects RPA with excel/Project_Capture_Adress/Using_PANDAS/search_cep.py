"""
Modúlo para consulta de Ceps via API ViaCEP e armazenamento 
dos resultados em uma nova planilha Excel.

Fluxo:
    1. Lê os CEPs da aba 'CEP' do arquivo Excel.
    2. Consulta cada CEP na API ViaCEP.
    3. Salva os endereços encontrados na aba 'Dados' do mesmo arquivo.


"""

import pandas as pd
from pathlib import Path
import http.client
import json
DIR = Path(__file__).parent
FILE = DIR / 'CEP.xlsx'
COLUMNS = ["cep","logradouro", "bairro", "localidade", "uf"]


def get_address_cep(cep: str) -> dict[str,str] | None:
    """
    Consulta a API ViaCEP e retorna os dados de endereço do CEP informado.

    Args:
        cep (str): CEP a ser consultado, somente numeros (ex: 01001000)

    Returns:
        dict[str,str]: Dicionário com os dados do endereço se encontrado.
        None: Se o CEP for inválido, não encontrado ou ocorrer erro na requisição.
    """
    connection = http.client.HTTPSConnection("viacep.com.br")
    try:
        # Envia requisição GET para a API ViaCEP com o CEP informado
        connection.request("GET", f"/ws/{cep}/json")
        response = connection.getresponse()
        # Retorna None se a API não responder com  status 200 OK
        if response.status != 200:
            print(f"CEP {cep}: status HTTP inesperado ({response.status})")
            return None
        # Decodifica a resposta JSON em UTF-8
        address = json.loads(response.read().decode("utf-8"))

        if "erro" in address:
            print(f"CEP {cep} não encontrado.")
            return None
        return address 
    except Exception as e:
        print(f"Aconteceu um erro durante a conexão: {e}")
        connection.close()
        return None
    finally:
        # Garante o fechamento da conexão independente do resultado.
        connection.close()
        

if __name__ =='__main__':

    # Lê os CEPs da aba 'CEP', ignorando células vazias
    sheet = pd.read_excel(FILE, sheet_name="CEP")
    ceps = sheet['CEP'].dropna()
    
    # Carrega CEPs já salvos para evitar requisiçoes desnecessárias.
    try:
        sheet_dados = pd.read_excel(FILE, sheet_name='Dados')
        ceps_salvos = set(sheet_dados['cep'].astype(str).str.replace("-","", regex=False).tolist())
    except Exception:
        sheet_dados = pd.DataFrame()
        ceps_salvos = set()


    # Filtra apenas CEPs ainda não consultados
    ceps_novos = [c for c in ceps if str(c).replace("-", "") not in ceps_salvos]
    print(f"{len(ceps_novos)} CEP(s) novo(s) para consultar.")

    print(f"CEPs já salvos: {ceps_salvos}")
    print(f"CEPs novos para consultar: {ceps_novos}")
    rows = []
    # Consulta cada CEP na API e acumula os resultados válidos.
    for cep in ceps_novos:
        address = get_address_cep(str(cep).replace("-", ""))

        if address:
            rows.append({
                'cep':        cep,
                'Logradouro': address.get('logradouro', ''),
                'Bairro':     address.get("bairro", ''),
                'Localidade': address.get("localidade", ''),
                'UF':         address.get("uf", '')
            })
     
    if rows:
        df_novo = pd.DataFrame(rows)
        df_final = pd.concat([sheet_dados, df_novo], ignore_index=True)
     
        # Salva todos os endereços encontrados na aba 'Dados' do Excel
        results = pd.DataFrame(rows)
        with pd.ExcelWriter(FILE,engine="openpyxl",mode='a',if_sheet_exists='replace') as writer:
            df_final.to_excel(writer, sheet_name='Dados', index=False)
        
        
        print(f"{len(rows)} endereço(s) salvos na aba 'Dados'.")