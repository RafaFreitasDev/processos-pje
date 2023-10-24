# entrar no site do pje "https://pje-consulta-publica.tjmg.jus.br/"
# digitar nº oab e selecionar estado
# clicar em pesquisar
# entrar em cada um dos processos
# extrair o nº do processo e a dta de distribuição
# extrair todas as últimas movimentações
# guardar tudo em uma aba da planilha excel por processo

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select

from time import sleep


from scraping import Scraping


if __name__ == "__main__":
    code = Scraping()
    code.Start()


    




