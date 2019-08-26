import re
import os
import numpy as np
import pandas as pd
import datetime
import json
from argparse import ArgumentParser
from dateutil import parser, relativedelta
from gooey import Gooey, GooeyParser


def readData(peoday, msg, vac, dis):
    msg = pd.read_excel(msg)
    peoday = pd.read_csv(peoday, low_memory=False)
    vacdata = pd.read_csv(vac)
    disdata = pd.read_excel(dis, sheet_name= "綠點派遣-在職名單")
    return peoday, msg, vacdata, disdata


peoday = r"./data/people_day.csv"
vac = r"./data/已批准休假.csv"
msg = r"./data/error.xlsx"
dis = r"./data/派遣新進離職資訊-2018-new.xlsx"
outdir = r"./result/"

peoday, msg, vacdata, disdata = readData(msg = msg,
            peoday = peoday,
            vac = vac,
            dis = dis)