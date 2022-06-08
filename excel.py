from my_imports import *
from globals import _total_atendimento, _total_deslocamento, _total_gaps, _total_suporte


def conversor_planilha():

    db = dbConnection()

    ##### PRIMEIRA PARTE, LEITURA DO PORTFOLIO PARA PEGAR SOL MÃE #####
    df = pd.read_excel('Portfolio.xlsx')
    #seto a linha dos titulos como as colunas do dataframe
    df.columns = df.iloc[5]
    #remove linhas inuteis do dataframe
    df.drop(df.index[[0,1,2,3,4]], inplace=True)
    #pego sol agrupadora
    # sol_agrupadora = df.iloc[9]['Solicitação Agrupadora']
    sol_agrupadora = 167458
    # sols_filhas = db.query(f""" select ss.nr_solicitacao
    #                             from solicservico_v2 ss
    #                             start with ss.nr_solicitacao = {sol_agrupadora} 
    #                             connect by prior ss.nr_solicitacao = ss.nr_solicitacao_pai """)
    sols_filhas = db.query(f""" select * from solicservico_v2 where nr_solicitacao_pai = {sol_agrupadora} """)
    
    sols_filhas = [e[0] for e in sols_filhas.values]

    print('sol pai', sol_agrupadora)

    sols_ocupadas=  classifica(sol_agrupadora, sols_filhas, db)

    #### SEGUNDA PARTE, PEGO A STATUS REPORT ONDE PREENCHEREMOS OS DADOS ####
    planilha =  load_workbook('statusreport.xlsx', data_only=False)
    ws = planilha['Status Report Período']
    ws[f'B1'].value = "Status Report do Projeto: Projeto Cooperalfa"
    ws = planilha['Acompanhamento Solics. Projeto']
    start, count, horas_total = preenche_atendimento(ws, planilha, sols_filhas, db, sols_ocupadas)
    _total_atendimento = faz_total(count, start, horas_total, ws)
    start, count, horas_total = preenche_deslocamento(ws, planilha, sols_filhas, db)
    _total_deslocamento = faz_total(count, start, horas_total, ws)
    start, count, horas_total = preenche_gaps(ws, planilha, sols_filhas, sol_agrupadora, db)
    _total_gaps = faz_total(count, start, horas_total, ws)
    start, count, horas_total = preenche_suporte(ws, planilha, sols_filhas, db, sols_ocupadas, sol_agrupadora)
    _total_suporte = faz_total(count, start, horas_total, ws)

    ### TERCEIRA PARTE
    # ws = planilha['Status Report Período']
    # proj_id = get_proj_id(sol_agrupadora)
    # result = perc_expec(proj_id)
    # #percentual
    # ws[f'C8'].value = result[0]/100
    # #expectaviva
    # ws[f'C9'].value = result[1]/100
    # #Horas utilizadas do projeto de Implantação
    # ws[f'C16'].value = _total_atendimento.value
    # #Horas Utilizadas em Deslocamentos
    # ws[f'C22'].value = _total_deslocamento.value


    ## QUARTA PARTE
    ws = planilha['Acompanhamento Financeiro']

    start, count, horas_total = acompanhamento_financeiro(ws, planilha, sols_filhas, sol_agrupadora, db)
    _total_acompanhamento_financeiro = faz_total(count, start, horas_total, ws)

    planilha.save('teste.xlsx')

conversor_planilha()