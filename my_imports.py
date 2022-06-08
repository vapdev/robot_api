from doctest import DocFileSuite
from urllib.request import AbstractBasicAuthHandler
import pandas as pd
from db import dbConnection
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, NamedStyle
from functions.utils import *
from autTabPortfolio import *
from functions.preencher.atendimento import *
from functions.preencher.deslocamento import *
from functions.preencher.gaps import *
from functions.preencher.suporte import *
from functions.preencher.acompanhamento_financeiro  import *
import time