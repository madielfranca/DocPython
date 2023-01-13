import pandas as pd
import numpy as np
import sys, os
import time
from datetime import datetime, date
from numbers import Number
import calendar

"""## Retorna os programas em exibição"""

def emExibicao(df_dados_cv, criterio_amortizacao, data_competencia):
  '''
    Verifica se os campos "Início Exibição" e
    "Fim Exibição" estão com valores de datas e 
    filtra de acordo com o critério de amortização
    fornecido.

    :param df_dadosCV: dataframe a ser analisádo
    :type df_dadosCV: pandas.core.frame.DataFrame

    :param criterio_amortizacao: dataframe a ser analisádo
    :type criterio_amortizacao: string

    :returns: resultado
    :rtype: pandas.core.frame.DataFrame
  '''
  # filtra dataframe onde inicio e fim da exibição sao datas
  serie_inicio_exibicao = df_dados_cv['Início Exibição'][df_dados_cv['Início Exibição'].apply(isinstance, args=(date,))]
  serie_fim_exibicao = df_dados_cv['Fim Exibição'][df_dados_cv['Fim Exibição'].apply(isinstance, args=(date,))]
  df_dados_datas_validas = df_dados_cv.loc[df_dados_cv['Início Exibição'].isin(serie_inicio_exibicao)]
  df_dados_datas_validas = df_dados_cv.loc[df_dados_cv['Fim Exibição'].isin(serie_fim_exibicao)]

  # Seleciona apenas os casos onde o programa está em exibição.
  return df_dados_datas_validas.loc[(df_dados_datas_validas['Início Exibição'] <= data_competencia) & 
                                  (df_dados_datas_validas['Fim Exibição'] >= data_competencia) &
                                  (df_dados_datas_validas['Critério Amortizacao'] == criterio_amortizacao)]

"""## Valida Rateio"""

def validaRateio(df_dados_cv):
  '''
    As somas da coluna RATEIO devem ser igual a 100% por projeto 
    ou (2)se o projeto não estiver com rateios definidos entre 
    as janelas, este campo poderá ficar como "A definir", 
    nesse caso só poderá existir uma linha para esse projeto, 
    ou seja, uma janela (3)Este campo deverá ter a informação 
    de Rateio preenchida com um percentual definido, uma vez que 
    a data de início de exibição seja maior que a data de 
    competência.

    :param df_dadosCV: dataframe a ser analisádo
    :type df_dadosCV: pandas.core.frame.DataFrame

    :returns: resultado
    :rtype: pandas.core.frame.DataFrame
  '''
  # Filtra os dados numéricos
  df_dados_rateio = df_dados_cv[df_dados_cv['Rateio']
                                .apply(isinstance, args=(Number,))]

  # Soma os Rateios para o mesmo projeto
  serie_rateio = df_dados_rateio[['Rateio','PROJETO']
                                 ].groupby('PROJETO').sum()

  # Retorna apenas os projetos em que o rateio apresenta diferença
  return df_dados_cv[df_dados_cv['PROJETO']
                     .isin(serie_rateio['Rateio']
                           [serie_rateio['Rateio'] != 100].keys())]

"""## Valida Total Capitulos"""

def validaTotalCapitulos(df_dados_cv, data_competencia):
  '''
    Esse campo deverá ter o preenchimento obrigatório para
    os casos de Janela com critério de amortização="Capítulos
    Exibidos", onde o Início de Exibição tenha se iniciado e
    o Fim de Exibição não tenha acabado. Início de Exibição =>
    Data de Competência e Fim de Exibição <= Data de competência.

    :param df_dadosCV: dataframe a ser analisádo
    :type df_dadosCV: pandas.core.frame.DataFrame

    :returns: resultado
    :rtype: pandas.core.frame.DataFrame
  '''
  criterio_amortizacao = 'Capítulos exibidos'
  df_dados_em_exibicao = emExibicao(df_dados_cv, criterio_amortizacao, data_competencia)
  
  return df_dados_em_exibicao[pd.to_numeric(df_dados_em_exibicao['Quantidade de episódios'], errors='coerce').isnull()]

"""## Valida Capítulos no mês"""

def validaCapitulosNoMes(df_dados_cv, data_competencia):
  '''
    Esse campo deverá ter o preenchimento obrigatório para 
    os casos de Janela com critério de amortização="Capítulos 
    Exibidos", onde o Início de Exibição tenha se iniciado e 
    o Fim de Exibição não tenha acabado. Início de Exibição => 
    Data de Competência e Fim de Exibição <= Data de competência.

    :param df_dadosCV: dataframe a ser analisádo
    :type df_dadosCV: pandas.core.frame.DataFrame

    :returns: resultado
    :rtype: pandas.core.frame.DataFrame
  '''
  criterio_amortizacao = 'Capítulos exibidos'
  df_dados_em_exibicao = emExibicao(df_dados_cv, criterio_amortizacao, data_competencia)

  return df_dados_em_exibicao[pd.to_numeric(df_dados_em_exibicao['Eps.  exibidos no mês'], errors='coerce').isnull()]

"""## Valida Período de Exclusividade"""

def validaPeriodoExclusividade(df_dados_cv, data_competencia):
  '''
    Esse campo deverá ter o preenchimento obrigatório para os 
    casos de Janela com critério de amortização="Meses de 
    Exclusividade", Início de Exibição => Data de Competência 
    e Fim de Exibição <= Data de competência.

    :param df_dadosCV: dataframe a ser analisádo
    :type df_dadosCV: pandas.core.frame.DataFrame

    :returns: resultado
    :rtype: pandas.core.frame.DataFrame
  '''
  criterio_amortizacao = 'Meses de exclusividade'
  df_dados_em_exibicao = emExibicao(df_dados_cv, criterio_amortizacao, data_competencia)

  return df_dados_em_exibicao[pd.to_numeric(df_dados_em_exibicao['Tempo de Exclusividade Globoplay'], errors='coerce').isnull()]

"""## Valida Janela"""

def validaJanela(df_dados_cv):
  '''
    Este campo deverá ter a informação de Janela 
    preenchida diferente de "A definir", uma vez 
    que a data de início de exibição seja maior 
    que a data de competência.

    :param df_dadosCV: dataframe a ser analisádo
    :type df_dadosCV: pandas.core.frame.DataFrame

    :returns: resultado
    :rtype: pandas.core.frame.DataFrame
  '''
  serie_inicio_exibicao = df_dados_cv['Início Exibição'][df_dados_cv['Início Exibição'].apply(isinstance, args=(date,))]
  df_dados_datas_inicio = df_dados_cv.loc[df_dados_cv['Início Exibição'].isin(serie_inicio_exibicao)]
  df_dados_janela = df_dados_datas_inicio.loc[(df_dados_datas_inicio['Início Exibição'] < datetime.now())]

  return df_dados_janela.loc[df_dados_janela['Cód. Centro de Resultado (Janela)'] == 'A definir']

"""## Main"""

def valida_cv(path_cv, path_resultado, data_competencia_str):
  # importa dados cliclo de vida

  try:
    
    data_competencia = datetime.strptime(data_competencia_str, '%d/%m/%Y')

    df_dados_cv = pd.read_excel(path_cv, sheet_name="CICLO DE VIDA")

    # limpa primeira coluna vazia
    df_dados_cv = df_dados_cv.drop(df_dados_cv.columns[0], axis=1)

    # verifica o tipo de projeto, splitando o código do projeto.
    #f = lambda item: item.split('.', maxsplit=2)[1]
    f = lambda item: item[:4]
    df_dados_cv['tipo_projeto'] = df_dados_cv['PROJETO'].apply(f)

    # Valida critério de amortização
    condicoes = [
      (df_dados_cv['tipo_projeto'].str.upper()!='PRLT') & (df_dados_cv['Descrição do Centro de Resultado'].str.upper() =='GLOBOPLAY'),
      (df_dados_cv['Estoque/Giro'].str.upper()=='GIRO')          
    ]
    escolhas = ['Meses de exclusividade', 'Custo mensal total']
    df_dados_cv['Critério Amortizacao'] = np.select(condicoes, escolhas, default='Capítulos exibidos')

    # Verifica se planilha é válida
    valida_rateio = validaRateio(df_dados_cv)
    valida_total_capitulos = validaTotalCapitulos(df_dados_cv, data_competencia)
    valida_capitulos_mes = validaCapitulosNoMes(df_dados_cv, data_competencia)
    valida_periodo_exclusividade = validaPeriodoExclusividade(df_dados_cv, data_competencia)
    valida_janela = validaJanela(df_dados_cv)

    result_validation = ("OK", "OK")

    if not valida_rateio.empty:
      result_validation = ("NOK", f"Existe um problema no arquivo de Ciclo de Vida. A área de negócios deve verificar. Um arquivo com as inconsistências está em {path_resultado}")
    if not valida_total_capitulos.empty:
      result_validation = ("NOK", f"Existe um problema no arquivo de Ciclo de Vida. A área de negócios deve verificar. Um arquivo com as inconsistências está em {path_resultado}")
    if not valida_capitulos_mes.empty:
      result_validation = ("NOK", f"Existe um problema no arquivo de Ciclo de Vida. A área de negócios deve verificar. Um arquivo com as inconsistências está em {path_resultado}")
    if not valida_periodo_exclusividade.empty:
      result_validation = ("NOK", f"Existe um problema no arquivo de Ciclo de Vida. A área de negócios deve verificar. Um arquivo com as inconsistências está em {path_resultado}")
    if not valida_janela.empty:
      result_validation = ("NOK", f"Existe um problema no arquivo de Ciclo de Vida. A área de negócios deve verificar. Um arquivo com as inconsistências está em {path_resultado}")

    if result_validation[0] == "NOK":

      # Imprime planilha
      writer = pd.ExcelWriter(path_resultado, engine='openpyxl')
      df_dados_cv.to_excel(writer, 'Ciclo_de_vida', index=False)
      valida_rateio.to_excel(writer, 'Rateio', index=False)
      valida_total_capitulos.to_excel(writer, 'Total de Capitulos', index=False)
      valida_capitulos_mes.to_excel(writer, 'Capitulos por mes', index=False)
      valida_periodo_exclusividade.to_excel(writer, 'Periodo Exclusividade', index=False)
      valida_janela.to_excel(writer, 'Janela', index=False)
      writer.save()

      return result_validation



    return result_validation 
        
  except Exception as e:
    print(f"Falha ao gerar relatório")
    print('-----------------------------------------------------------------------')
    print('-----------------------------------------------------------------------')
    print(f"Tipo do erro: {e}")

    exc_type, exc_obj, exc_tb = sys.exc_info()

    print('-----------------------------------------------------------------------')
    print('-----------------------------------------------------------------------')
    print('-----------------------------------------------------------------------')
    print(f'O erro se encontra na linha: {exc_tb.tb_lineno} ')

    print('-----------------------------------------------------------------------')
    print('-----------------------------------------------------------------------')
    print('-----------------------------------------------------------------------')

    breakpoint()


try:

  data_e_hora_atuais = datetime.now()
  dia = data_e_hora_atuais.strftime("%d")
  mes = data_e_hora_atuais.strftime("%m")
  ano = data_e_hora_atuais.strftime("%Y")
  ano = int(ano)
  mes = int(mes) -1

  if int(mes) == 0:

    ano = int(ano) -1
    mes = 12

  dia = int(dia)
  test_date = datetime(ano, mes, dia)
  res = calendar.monthrange(test_date.year, test_date.month)[1]
  data = f'{res}/{mes}/{ano}'

except Exception as e:
  print(f"Falha ao gerar relatório")
  print('-----------------------------------------------------------------------')
  print('-----------------------------------------------------------------------')
  print(f"Tipo do erro: {e}")

  exc_type, exc_obj, exc_tb = sys.exc_info()

  print('-----------------------------------------------------------------------')
  print('-----------------------------------------------------------------------')
  print('-----------------------------------------------------------------------')
  print(f'O erro se encontra na linha: {exc_tb.tb_lineno} ')

  print('-----------------------------------------------------------------------')
  print('-----------------------------------------------------------------------')
  print('-----------------------------------------------------------------------')

  breakpoint()


if __name__ == '__main__':

    path_cv = 'Ciclo_de_vida.xlsx'
    path_resultado = r'Resultado.xlsx'
    data_competencia_str = data

df = pd.read_excel(path_cv, sheet_name="CICLO DE VIDA")

for linha in range(len(df)):
  if 'A definir' in str(df['Início Exibição'].iloc[linha]) and 'A definir' != str(df['Fim Exibição'].iloc[linha]):
    print('data de inicio incorreta')
    print(f'O erro se encontra na linha: {linha+2} ')
    print(f'Projeto:')
    print(df['PROJETO'].iloc[linha])
    print('------------------------')
    print('------------------------')





valida_cv(path_cv, path_resultado, data_competencia_str)
print('-----------------------------------------------------------------------')
print('-----------------------------------------------------------------------')
print('-----------------------------------------------------------------------')
print('processo executado com sucesso')
time.sleep(10)






