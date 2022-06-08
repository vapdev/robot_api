from fastapi import FastAPI
from db import dbConnection
import pandas as pd
from excel import conversor_planilha

app = FastAPI()
db = dbConnection()

@app.get("/")
async def root():
    return {"message": "API de ferramentas do chatbot Maxicon"}

@app.get("/apps/{app_id}")
async def app_action(app_id):
    
    if app_id == "cv_planilha":
        conversor_planilha()
    if app_id == "maxicon":
        pass
    if app_id == "maxicon":
        pass
    
    return {"app_id": app_id}

