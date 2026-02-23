"""
Modulo utilizado para criar planilhas do Excel utilizando a biblioteca externa xlsxwriter.

"""

#Bibliotecas importadadas.
import  xlsxwriter as xl
from xlsxwriter import format
# Biblioteca OS para abrir arquivo após o workbook ser fechado. No meu caso não tenho Excel instalado, 
# então não o usei.
# import os 



# Observe que importei Path para pegar o  caminho completo do arquivo.
# Mas se não quiser, tambem pode usar o caminho completo em uma variavel. 
# (ex. 'C:\xxx\xxx\xxx\xxx\Projetos_RPA\Projects_xlsxwriter')
from pathlib import Path

# Caminho do arquivo a ser criado.
DIR = Path(__file__).parent # C:\xxx\xxx\xxx\xxx\Projetos_RPA\Projects_xlsxwriter
FILE = DIR / 'dados.xlsx' #C:\xxx\xxx\xxx\xxx\Projetos_RPA\Projects_xlsxwriter\dados.xlsx
wb = xl.Workbook(FILE)

# Nome da sheet ativa que será usada
sheet = wb.add_worksheet('Dados')




# Funções para formatar celulas.
def set_fot_color(color_font: str) -> format.Format:
    """
    Função para mudar a cor de uma fonte.
    Args:
        color_font (str): Pode ser uma cor em codigo decimal, rgba ou por extenso.
    Returns:
        out: Referencia a uma modificação de um objeto .xlsx
    """
    standart = wb.add_format()
    standart.set_font_color(color_font)
    return standart


def set_backgroud_color(color_font: str) -> format.Format:    
    """
    Função para mudar a cor de fundo de uma celula.
    Args:
        color_font (str): Pode ser uma cor em codigo decimal, rgba ou por extenso.
    Returns:
        out: Referencia a uma modificação de um objeto .xlsx
    """
    standart = wb.add_format()
    standart.set_bg_color(color_font)
    return standart





sheet.write("A1",'Nome',set_backgroud_color('yellow'))
sheet.write("B1",'IDADE', set_backgroud_color('yellow'))
sheet.write("C1", 'Origem', set_backgroud_color('yellow'))
sheet.write("A2",'Augusto',set_fot_color('blue'))
sheet.write("B2","12",set_fot_color('blue'))

wb.close()

