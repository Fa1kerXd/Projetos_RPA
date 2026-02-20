"""
Módulo para validação e consulta de CNPJs via API ReceitaWS.

Fluxo:
    1. Valida o CNPJ informado via regex e dígitos verificadores.
    2. Consulta os dados da empresa na API ReceitaWS.
    3. Retorna as informações em formato dicionário.
    4. Salva as informações da empresa em uma Planilha Excel de forma automatizada.

Limites da API ReceitaWS:
    - Máximo de 3 requisições por minuto no plano gratuito.
    - Timeout padrão retorna status 504.
"""

import http.client
import json
import re
import pandas as pd
from pathlib import Path
import time
RECEITAWS_HOST = "www.receitaws.com.br"
DIR = Path(__file__).parent
FILE = DIR / "data_cnpj.xlsx"


def validator_cnpj(cnpj: str) -> bool:
    """
    Valida um CNPJ verificando formato e dígitos verificadores.

    Args:
        cnpj (str): CNPJ no formato '00.000.000/0000-00' ou '00000000000000'.

    Returns:
        bool: True se o CNPJ for válido, False caso contrário.

    Validações realizadas:
        1. Formato via regex (com ou sem máscara).
        2. Rejeita CNPJs com todos os dígitos iguais (ex: 11111111111111).
        3. Verifica os dois dígitos verificadores pelo algoritmo oficial da Receita Federal.
    """
    pattern = r'^(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}|\d{14})$'

    if not re.fullmatch(pattern, cnpj):
        return False

    cnpj_numbers = re.sub(r'\D', '', cnpj)

    if cnpj_numbers == cnpj_numbers[0] * 14:
        return False

    def calc_digit(cnpj_partial: str) -> str:
        """
        Calcula um dígito verificador pelo algoritmo oficial da Receita Federal.
        """
        weights = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
        total = sum(
            int(d) * weights[i + len(weights) - len(cnpj_partial)]
            for i, d in enumerate(cnpj_partial)
        )
        rest = total % 11
        return '0' if rest < 2 else str(11 - rest)

    first_digit = calc_digit(cnpj_numbers[:12])
    second_digit = calc_digit(cnpj_numbers[:12] + first_digit)

    return cnpj_numbers[-2:] == first_digit + second_digit


def cnpj_format(cnpj: str) -> str:
    """
    Remove caracteres não numéricos do CNPJ, mantendo apenas dígitos.

    Args:
        cnpj (str): CNPJ com ou sem formatação (ex: '00.000.000/0000-00').

    Returns:
        str: CNPJ contendo apenas números (ex: '00000000000000').
    """
    return re.sub(r'\D', '', cnpj)




def read_cnpj_in_sheet(file=FILE) -> list[str]:
    if not file.exists():
        print(f"Arquivo {file} não encontrado. Crie a planilha com uma aba 'CNPJS' contendo os CNPJs.")
        return []
    


    try:
        # Lê os CNPJs da aba de entrada
        df_cnpj_input = pd.read_excel(file, sheet_name='CNPJS')


        try:
            # Lê os CNPJs consultados
            df_data = pd.read_excel(file,sheet_name='Dados')
            cnpjs_consultados = set(df_data['CNPJ'].astype(str).tolist())
        except Exception:
            cnpjs_consultados = set()

        # Normaliza e filtra apenas CNPJs ainda não consultados
        cnpjs_pendente = []
        for cnpj in df_cnpj_input['CNPJ'].dropna():
            cnpj_normalizado = cnpj_format(cnpj)

            if cnpj_normalizado in cnpjs_consultados:
                print(f"CNPJ {cnpj_normalizado} já foi consultado, pulando...")
                continue
            cnpjs_pendente.append(cnpj_normalizado)
            print(f"CNPJ {cnpj_normalizado} adicionado à fila de consulta.")
        return cnpjs_pendente 


    except Exception as e:
        print(f"Erro ao ler a planilha: {e}")
        return []


   


def get_cnpj_with_httpclient(cnpj: str) -> dict[str, str] | None:
    """
    Consulta os dados de uma empresa na API ReceitaWS a partir do CNPJ.

    Args:
        cnpj (str): CNPJ a ser consultado (com ou sem máscara).

    Returns:
        dict[str, str]: Dicionário com as informações da empresa.
        None: Se o CNPJ for inválido, não encontrado, timeout ou limite atingido.

    Status HTTP tratados:
        200: Sucesso — retorna os dados da empresa.
        429: Limite de requisições atingido (máx. 3 por minuto no plano gratuito).
        504: Timeout — tempo máximo de requisição excedido.
    """
    if not validator_cnpj(cnpj):
        print("CNPJ inválido.")
        return None

    cnpj = cnpj_format(cnpj)
    connection = http.client.HTTPSConnection(RECEITAWS_HOST)
    try:
        connection.request("GET", f"/v1/cnpj/{cnpj}")
        response = connection.getresponse()
        status = response.status

        if status == 429:
            print("Limite de requisições atingido. Máximo: 3 por minuto.")
            return None

        if status == 504:
            print("Timeout — tempo máximo de requisição excedido.")
            return None

        if status == 200:
            print("Requisição bem-sucedida (200 OK).")
            return json.loads(response.read().decode("utf-8"))

        print(f"Status inesperado recebido: {status}")
        return None

    except json.JSONDecodeError as e:
        print(f"Erro na decodificação do JSON: {e}")
        return None

    except Exception as e:
        print(f"Erro na requisição com a API: {e}")
        return None

    finally:
        connection.close()


def save_data_in_sheet(data: dict[str, str], file: Path = FILE) -> None:
    """
    Salva dados de CNPJ em uma planilha Excel, evitando duplicatas.

    Args:
        data (dict[str, str]): Dicionário com os dados da empresa retornados pela API.
        file (Path): Caminho do arquivo Excel de destino.

    Abas criadas:
        - 'Dados': Informações completas de todas as empresas consultadas.
        - 'CNPJS': Lista de CNPJs já salvos para evitar duplicatas.

    Returns:
        None
    """
    # Carrega ou cria o arquivo Excel
    if file.exists():
        try:
            df_data = pd.read_excel(file, sheet_name="Dados")
        except Exception:
            df_data = pd.DataFrame(columns=["CNPJ", "nome", "situacao", "atividade_principal", "cep", "email"])

        try:
            df_cnpjs = pd.read_excel(file, sheet_name="CNPJS")
        except Exception:
            df_cnpjs = pd.DataFrame(columns=['CNPJS'])

    else:
        df_data = pd.DataFrame(columns=["CNPJ", "nome", "situacao", "atividade_principal", "cep", "email"])
        df_cnpjs = pd.DataFrame(columns=['CNPJS'])

    # Normaliza o CNPJ e verifica se já foi consultado
    cnpj = cnpj_format(data.get("cnpj", ""))


    # Extrai atividade principal (é uma lista de dicts)
    atividade = "N/A"
    if data.get("atividade_principal"):
        atividade = data["atividade_principal"][0].get("text", "N/A") #type: ignore

    # Adiciona nova linha com os dados da empresa
    nova_linha = pd.DataFrame([{
        "CNPJ": data.get("cnpj", "N/A"),
        "nome": data.get("nome", "N/A"),
        "situacao": data.get("situacao", "N/A"),
        "atividade_principal": atividade,
        "cep": data.get("cep", "N/A"),
        "email": data.get("email", "N/A"),
    }])


    # Concatena com os dados existentes
    df_data_final = pd.concat([df_data, nova_linha], ignore_index=True)

    # Salva mantendo a aba CNPJ intact.
    with pd.ExcelWriter(file, engine="openpyxl") as writer:
        df_data_final.to_excel(writer, sheet_name="Dados", index=False)
        df_cnpjs.to_excel(writer, sheet_name='CNPJS', index=False)

    print(f"CNPJ {cnpj} salvo com sucesso.")


if __name__ == "__main__":
    print("=== Iniciando consulta de CNPJs ===\n")

    cnpj_pendente = read_cnpj_in_sheet()
    if not cnpj_pendente:
        print("Nenhum CNPJ pendente para consultar.")
    else:
        print(f"\n{len(cnpj_pendente)} CNPJ(s) para consultar.\n")

        for i, cnpj in enumerate(cnpj_pendente, 1):
            print(f"[{i}/{len(cnpj_pendente)}] Consultando CNPJ {cnpj}...\n")
            result = get_cnpj_with_httpclient(cnpj)

            if result:
                save_data_in_sheet(result)
                # Aguarda 20s entre requisições para respeitar o limite da API.
                if i < len(cnpj_pendente):
                    print("Aguardando 20 segundos antes da próxima consulta...\n")
                    time.sleep(20)