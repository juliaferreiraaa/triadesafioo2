from planilhaebot import bot as pl
import enviar as em
import cliques as cl
from botcity.core import DesktopBot
import pandas as pd
import zipfile
import os


def main():
    file_name_test = 'Tabela_de_Dados'
    ceps_list = []

    bot = DesktopBot()
    bot.browse("https://www.geradordecep.com.br/")

    pl.criar_planilha(file_name_test)

    bot.wait(500)

    while len(ceps_list) < 5:
        cl.clique_no_botao(bot, not_found)
        cl.copiar_cep(bot, not_found, ceps_list)
        pl.atualizar_tabela1(file_name_test, 'Cep', 'CEP', ceps_list)
        pl.consumir_api(file_name_test)
        ceps_list = list(set(ceps_list))   #  Remover ceps duplicados na lista
        bot.wait(500)
        print(ceps_list)


    bot.browse('https://forms.gle/UwtGAF7YyTzL52kd6')
    url_filename = pd.read_excel(f'C:/Users/julia/PycharmProjects/botcep/botdesafio/{file_name_test}.xlsx', sheet_name="Modelo")
    for i, name in enumerate(url_filename['CEP']):
        if not bot.find( "preenche_cep", matching=0.97, waiting_time=10000):
            not_found("preenche_cep")
        bot.click_relative(23, 74)
        bot.paste(url_filename.loc[i, 'CEP'])
        bot.tab()

        bot.paste(url_filename.loc[i, 'Logradouro'])
        bot.page_down()
        bot.tab()

        bot.paste(url_filename.loc[i, 'Bairro'])
        bot.tab()

        bot.paste(url_filename.loc[i, 'Localidade'])
        bot.tab()

        bot.paste(url_filename.loc[i, 'UF'])
        bot.tab()
        bot.wait(500)

        bot.paste(str(url_filename.loc[i, 'DDD']))

        bot.control_p()

        if not bot.find("clicareselecionardestino", matching=0.97, waiting_time=10000):
            not_found("clicareselecionardestino")
        bot.click_relative(286, 22)

        bot.wait(2000)

        if not bot.find("salvarcomopdf2", matching=0.97, waiting_time=10000):
            not_found("salvarcomopdf2")
            if not bot.find("salvarcomopdf", matching=0.97, waiting_time=10000):
                not_found("salvarcomopdf")
        bot.click()

        if not bot.find("clickemsalvar", matching=0.97, waiting_time=10000):
            not_found("clickemsalvar")
        bot.click()

        if not bot.find( "clique_salva", matching=0.96, waiting_time=10000):
            not_found("clique_salva")
        if not bot.find( "icone_outras_pastas", matching=0.97, waiting_time=10000):
            not_found("icone_outras_pastas")
        bot.click()
        bot.delete()
        bot.delete()
        bot.paste("C:/Users/julia/PycharmProjects/botcep/botdesafio")
        bot.enter()

        if not bot.find("nomearpdf", matching=0.97, waiting_time=10000):
            not_found("nomearpdf")
        bot.click_relative(81, 9)
        bot.delete()
        bot.paste(url_filename.loc[i, 'CEP'])
        bot.tab()
        bot.tab()
        bot.tab()
        bot.enter()

        bot.page_down()

        if not bot.find( "limpa_formulario", matching=0.97, waiting_time=10000):
            not_found("limpa_formulario")
        bot.click()

        if not bot.find( "limpa_form2", matching=0.97, waiting_time=10000):
            not_found("limpa_form2")
        bot.click()


    arquivo_zip = zipfile.ZipFile(r'C:\Users\julia\PycharmProjects\botcep\botdesafio\formularioszip', 'w')
    for pasta, subpastas, arquivos in os.walk(r'C:\\Users\julia\PycharmProjects\botcep\botdesafio'):
        for arquivo in arquivos:
            if arquivo.endswith(('.pdf', 'Tabela_de_Dados.xlsx')):
                arquivo_zip.write(os.path.join(pasta, arquivo), os.path.relpath(os.path.join(pasta, arquivo), r'C:\\Users\julia\PycharmProjects\botcep\botdesafio'), compress_type=zipfile.ZIP_DEFLATED)
    arquivo_zip.close()

    em.enviar_email()

def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
