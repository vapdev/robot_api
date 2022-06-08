from functions.utils import *
def preenche_atendimento(ws, planilha, sols_filhas, db, sols_ocupadas):
    horas_total = 0 
    count = 0
    for i, nr_solic in enumerate(sols_filhas):
        query_sol = f"select tp_solicitacao, nr_solicitacao, ds_assunto, st_solic, ds_status from solicservico_v2 a, statussolic b where a.nr_solicitacao = {nr_solic} and a.st_solic = b.cd_status"
        query = db.query(query_sol)
        
        # atribuo as variaveis que vieram da query do banco
        nr_solicitacao = query['nr_solicitacao'].values[0]
        ds_assunto = query['ds_assunto'].values[0]
        ds_status = query['ds_status'].values[0]
        tp_solicitacao = query['tp_solicitacao'].values[0]
        horas_realizadas = formata_horas(db.query(f'SELECT retorna_min_realizados_solic({nr_solicitacao})/60 minutos FROM DUAL')['minutos'].values[0])
        
        query_atendimento = db.query(
                            f"""
                             select retorna_usupapelsolicitacao({167458}, 18) from dual
                             """
        ).values[0]
        start = acha_start('atendimento', ws)
        if query_atendimento == 'FFF' and "deslocamento" not in ds_assunto.lower() and "mestra" not in ds_assunto.lower() and tp_solicitacao not in [2,3,10,12]: 
            count +=1
            horas_total += db.query(f'SELECT retorna_min_realizados_solic({nr_solicitacao})/60 minutos FROM DUAL')['minutos'].values[0]
            # verifico se alguma das células que estou preenchendo, foi erroneamente mergeada com as coluna E ou F, e então desfaço o merge
            
            e_merged = f'B{start+count-1}:E{start+count-1}'
            f_merged = f'B{start+count-1}:F{start+count-1}'
            lista = str(ws.merged_cells).split(' ')
            if e_merged in lista:
                ws.unmerge_cells(e_merged)
            if f_merged in lista:
                ws.unmerge_cells(f_merged)
            
            # insiro uma nova linha
            ws.insert_rows(start+count-1)
            
            preenche_linha(ds_assunto, 'B', start+count-1, 'left', 'no_right', ws, planilha) 
            preenche_linha(None, 'C', start+count-1, 'left', 'no_left', ws, planilha) 
            preenche_linha(nr_solicitacao, 'D', start+count-1, 'center', 'total', ws, planilha) 
            preenche_linha(ds_status, 'E', start+count-1, 'center', 'total', ws, planilha)  
            preenche_linha(1, 'F', start+count-1, 'center', 'total', ws, planilha) 
            ws[f'G{start+1}'].number_format = '[h]:mm:ss'
            preenche_linha('Maxicon', 'H', start+count-1, 'center', 'total', ws, planilha) 
            preenche_linha(horas_realizadas, 'G', start+count-1, 'center', 'total', ws, planilha)
            
    return (start, count, horas_total)