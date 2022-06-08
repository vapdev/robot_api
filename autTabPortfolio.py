import pandas as pd
import requests

def perc_expec(proj_id):

  url = "https://maxicon.teamwork.com/projects.json"
  request = requests.get(url, headers={'Authorization':'Basic dHdwX0g5ZEtmbTBqckdPM2lYWWEzWTFreXg5a0ZzVE06eHh4'}) 
  projetos = request.json()

  for d in projetos['projects']: 
      count = 0
      completed = 0
      request = requests.get("https://maxicon.teamwork.com/projects/api/v3/projects/"+d.get('id')+"/summary.json", 
                              headers={'Authorization':'Basic dHdwX0g5ZEtmbTBqckdPM2lYWWEzWTFreXg5a0ZzVE06eHh4'}) 
      summaryProject = request.json()
      for i in summaryProject['columns']['data']:
          cards = i.get('cards')
          count = count + cards['count']
          completed = completed + cards['completed']

      atividadesAtrasadas = summaryProject['tasks']['everyone']['late']
      if count > 0:
        percCompletudo = completed * 100 / count
        percAtividadesAtrasadas = (completed + atividadesAtrasadas) * 100 / count
        if percAtividadesAtrasadas > 100:
          percAtividadesAtrasadas = 100
  
      else:
        percCompletudo = 0
        percEstimado = 0
        percAtividadesAtrasadas = 0
      
      if (str(d.get('id'))) == proj_id:
        print (d.get('name')+': '+str(round(percCompletudo))+' --> '+str(round(percAtividadesAtrasadas)))
        percentual = percCompletudo
        expectativa = percAtividadesAtrasadas

        return (percentual,expectativa)
      
