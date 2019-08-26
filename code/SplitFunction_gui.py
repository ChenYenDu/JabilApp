import re
import os
import numpy as np
import pandas as pd
import datetime
import json
from argparse import ArgumentParser
from dateutil import parser, relativedelta
from gooey import Gooey, GooeyParser

@Gooey(program_name= "Dispatch Attendance List")
def parse_args():
    '''
    Use ArgParser to build up the arguments we will use in our script
    Save the arguments in a default json file so that we can retrieve
    every time we run the script
    '''
    stored_args = {}
    # get the script name without the extension & use it to build up
    # the json filename
    script_name = os.path.splitext(os.path.basename(__file__))[0]
    args_file = "{}-args.json".format(script_name)
    # Read the prior arguments as a dictionary
    if os.path.isfile(args_file):
        with open(args_file) as data_file:
            stored_args = json.load(data_file)
    parser = GooeyParser(description = 'Create Dispatch Attendance Report')

    parser.add_argument('peoday',
                        metavar = 'People Day',
                        action = 'store',
                        widget = 'FileChooser',
                        default = stored_args.get('peoday'),
                        help = "Select the People-Day data from PLT")
    parser.add_argument('msg',
                        metavar = 'Error Message',
                        action = 'store',
                        widget = 'FileChooser',
                        default = stored_args.get('msg'),
                        help = 'Select the abnoral message data from PLT')
    parser.add_argument('vac',
                        metavar = "Vacation Data",
                        action = 'store',
                        widget = 'FileChooser',
                        default = stored_args.get('vac'),
                        help = 'Select the approved vacation data from PLT')
    parser.add_argument('dis',
                        metavar = "Dispatch List",
                        action = 'store',
                        widget = 'FileChooser',
                        default = stored_args.get('dis'),
                        help = 'Select dispatch employee information data in shared folder')
    parser.add_argument('outdir',
                        metavar = "Output Directory",
                        action = 'store',
                        widget = 'DirChooser',
                        default = stored_args.get('output_directory'),
                        help = 'Select the folder to save files')
    
    parser.add_argument('-needBenefit',
                        metavar = '產生全勤獎金名單',
                        action = 'store_true',
                        help = "Select to create benefit list"
                        )
    parser.add_argument('-paidCount',
                        metavar = "有薪假算全勤",
                        action = "store_true",
                        help = 'Select to count paid as part of benefit')
    parser.add_argument('-monthStart',
                        metavar = "Month Start Date",
                        action = 'store',
                        widget = 'DateChooser',
                        help = 'Select the start date of month')
    
    #parser.add_argument('-cutdispatchComp',
    #                    metavar = '拆派遣公司檔案',
    #                    action = 'store_true',
    #                    help = 'Select if to cut by dispatch company')
    
    args = parser.parse_args()

    # Store the values of the arguments so we have them next time we run
    with open(args_file, 'w', encoding = 'UTF-8' ) as data_file:
        # Using vars(args) returns the data as a dictionary
        json.dump(vars(args), data_file)
    return args

def readData(peoday, msg, vac, dis):
    msg = pd.read_excel(msg)
    peoday = pd.read_csv(peoday, low_memory=False)
    vacdata = pd.read_csv(vac)
    disdata = pd.read_excel(dis, sheet_name= "綠點派遣-在職名單")
    return peoday, msg, vacdata, disdata

def manipulatePeoday():
    global peoday, disdata, outdir
    #####
    # manipulate dispatch data
    disdata = disdata[["SAP工號", "姓名", "單位", "派遣公司"]]
    disdata['SAP工號'] = disdata['SAP工號'].astype('category')

    #####
    # manipulate peoday data
    peoday['emp_id'] = peoday['emp_id'].astype('category')
    peoday['att_dt'] = peoday['att_dt'].astype('datetime64[ns]')

    # select XA employees
    peoday = peoday.loc[peoday.emp_subgroup == "XA",:]

    #####
    # merge poeday and disdata
    peoday = pd.merge(peoday, disdata, how= "left", left_on= "emp_id", right_on = 'SAP工號')

def manipulateMsg():
    global msg, needBenefit
    #####
    # manipulate msg data
    msg = msg[['員工編碼', '日期', '備註']]
    msg = msg.loc[msg.員工編碼.isin(disdata['SAP工號']),]
    msg.columns = ["emp_id", "att_dt","備註"]
    msg['emp_id'] = msg.emp_id.astype('category')

    # get arrive late mins and leave early mins
    msg['遲到'] = msg['備註'].str.extract(r'(遲.[0-9]+.[0-9]+)')
    msg['遲到'] = msg['遲到'].str.extract(r'([0-9]+.[0-9]+)').astype(float)

    msg['早退'] = msg['備註'].str.extract(r'(早.[0-9]+.[0-9]+)')
    msg['早退'] = msg['早退'].str.extract(r'([0-9]+.[0-9]+)').astype(float)
    
    if needBenefit == True:
        global msg_uq
        msg_uq = msg.copy()
        # count total minuts of late and early 
        msg_uq['LE'] = msg_uq[["遲到","早退"]].sum(axis= 1, skipna= True)
        # assign value according to LE
        msg_uq['LEC'] = 0
        isLess5 = (msg_uq.LE > 0) & (msg_uq.LE <=5)
        isMore5 = msg_uq.LE > 5
        msg_uq.loc[isLess5, 'LEC'] = 1
        msg_uq.loc[isMore5, 'LEC'] = 4
        # count total LEC for each employees
        msg_uq = msg_uq.groupby(["emp_id"])[["LEC"]].sum()
        msg_uq = msg_uq.reset_index()
        msg_uq = msg_uq.loc[msg_uq.LEC > 3,]
        msg_uq.to_excel("".join([outdir, r"/遲到早退排除名單.xlsx"]))
       
def manipulateVac():
    global vacdata, outdir
    if needBenefit == True:
        global vac_uq
        vac_uq = vacdata.copy()
        # do paid vac time data
        paid_leave = ['特休','補休','喪假','公假']
        isPaid = vac_uq.leave_type.isin(paid_leave)
        vac_uq = vac_uq[~isPaid] # select vacdata which is not paid
        vac_uq = vac_uq.groupby(['emp_id','leave_type'])[['leave_hours']].sum()
        vac_uq = vac_uq.dropna().reset_index() # remove na value and reset index to column
        vac_uq.to_excel("".join([outdir, r"/扣薪假明細.xlsx"]), header = True, index = None)
        
    #####
    ### manipulate vacdata
    # filter dispatch empoloyees
    vacdata['emp_id'] = vacdata['emp_id'].astype('category')
    vacdata = vacdata[vacdata['emp_id'].isin(disdata['SAP工號'])]

    # change start_date & start_time to datetime
    vacdata['start_date'] = pd.to_datetime(vacdata['start_date'], format= '%Y-%m-%d')
    vacdata['start_time'] = pd.to_datetime(vacdata['start_time'], format= '%H:%M').dt.time

    # identify the att_dt of vacdata
    # create a filter mask to filter data for start_time before 7:50
    isBefore0750 = vacdata['start_time'] < datetime.time(hour= 7, minute= 50)
    vacdata['att_dt'] = vacdata['start_date']

    # minus 1 day for start_time before 7:50
    vacdata.loc[isBefore0750, 'att_dt'] = vacdata['att_dt'] - datetime.timedelta(days = 1)

    # reshape vacdata to add columns for each vacation type
    vacdata = vacdata.pivot_table(index = ['emp_id', 'att_dt'],
                                columns = 'leave_type',
                                values = 'leave_hours')

    vacdata = vacdata.rename_axis(None, axis = 1).reset_index()
    
def mergeData():
    global peoday, msg, vacdata, outdir
    #####
    # merge all datas
    peoday['emp_id'] = peoday.emp_id.astype('category')
    peoday['att_dt'] = peoday.att_dt.astype('datetime64[ns]')

    global att_final
    att_final = pd.merge(peoday, msg, how= 'left', on= ['emp_id', 'att_dt'])
    att_final['emp_id'] = att_final.emp_id.astype('category')
    att_final = pd.merge(att_final, vacdata, how= 'left', on= ['emp_id', 'att_dt'])

    # drop unused columns and rename others
    att_final = att_final.drop(
        ["emp_code", "enjoy_offday_allowance", "non_standard_shift_type",
        "non_standard_shift_hours", "meal_times", "tea_times", "NT200Times",
        "NT500Times","SAP工號", "姓名"], axis = 1)
    
    old_cols = ['emp_id', 'name_cn','emp_subgroup', 'att_dt',
        'first_clock_in_time', 'last_clock_out_time', 'absent_hours',
        'working_hours', 'shift_code', 'shift_code_on_offday',
        'shift_startdatetime', 'shift_enddatetime', 'shift_type', 'ot_date',
        'ot_starttime', 'ot_endtime', 'ot_type_cn', 'ot_hours']

    new_cols = ["工號","姓名","ESG","出勤日期","上班卡點","下班卡點","缺勤時數",
        "應出勤時數", "班別", "假日班別", "班別開始","班別結束","平假日", "加班日期",
        "加班開始", "加班結束", "加班類型", "加班時數"]

    #rename columns
    att_final.rename(
        columns=dict(zip(att_final.columns[:18], new_cols)),inplace=True)
    att_final.rename(
        columns={"備註":"系統異常"},inplace = True)

    # save files
    fileName = "".join([outdir, r"/All_出勤明細.xlsx"])
    att_final.to_excel(fileName, header= True, index= None)

def cutData():
    global att_final, outdir
    unit_list = att_final.單位.unique() # distinct unit list
    if not os.path.exists("".join([outdir,r'/attList'])):# create folder if attList not exist 
        os.makedirs("".join([outdir, r'/attList']))
    for unit in unit_list:
        temp = att_final[att_final.單位 == unit]
        fileName = "".join([outdir, "/attList/", str(unit), r"_出勤明細.xlsx"])
        temp.to_excel(fileName, header= True, index= None)    


def benefitList():
    global att_final, vac_uq, outdir, monthStart, paidCount
    
    ## drop employees how come late and leave early
    att_final = att_final.loc[~att_final.工號.isin(msg_uq.emp_id),]

    ## filter employees who do not work for a full year
    monthStart = parser.parse(monthStart)
    monthEnd = monthStart + relativedelta.relativedelta(months= 1, days= -1)
    
    ## make month start end table
    ms = att_final.groupby("工號")[['出勤日期']].min().reset_index()
    ms.columns = ["emp_id","start"]
    me = att_final.groupby("工號")[['出勤日期']].max().reset_index()
    me.columns = ["emp_id","end"]
    mse = pd.merge(me, ms)
    isFullmonth = (mse.start == monthStart) & (mse.end == monthEnd)
    mse = mse.loc[isFullmonth,:]
    
    # drop employees who did not work full month
    att_final = att_final.loc[att_final.工號.isin(mse.emp_id),]

    # drop employees who have absent time
    absent_emp = att_final.loc[att_final.缺勤時數 > 0, "工號"]
    att_final = att_final.loc[~(att_final.工號.isin(absent_emp)),]

    if paidCount == True:
        att_final = att_final.loc[~att_final.工號.isin(vac_uq.emp_id),]
    else:
        leave_emp = att_final.loc[att_final.leave_hours != 0, "工號"]
        att_final = att_final.loc[~(att_final.工號.isin(leave_emp)),]

    balance_list = att_final[["工號","姓名","派遣公司","單位"]].drop_duplicates()
    balance_list.to_excel("".join([outdir, "/All_激勵獎金總名單.xlsx"]),header = True, index = None)

    unit_list = balance_list.單位.unique()
    if not os.path.exists("".join([outdir,r'/benefit'])):# create folder if attList not exist 
        os.makedirs("".join([outdir, r'/benefit']))
    for unit in unit_list:
        temp = balance_list[balance_list.單位 == unit]
        fileName = "".join([outdir, "/benefit/", str(unit), r"_激勵獎金總名單.xlsx"])
        temp.to_excel(fileName, header= True, index= None)  
        

if __name__ == '__main__':
    conf = parse_args()
    needBenefit = conf.needBenefit
    paidCount = conf.paidCount
    outdir = conf.outdir

    peoday, msg, vacdata, disdata = readData(msg = conf.msg,
            peoday = conf.peoday,
            vac = conf.vac,
            dis = conf.dis)
    print("Read data")

    manipulatePeoday() # merge peoday and disdata
    print("Manipulate Peopleday data")
    
    manipulateMsg()  # wrangle msg data
    print('Manipulate attendance error data')

    manipulateVac() # wrangle vac data
    print('Manipulate vacation data')

    mergeData() # merge peoday, msg, vac
    print('Merge all datas')
    
    cutData() # cut by unit as save
    print('Cut and save attendance data')

    # create benefit available list if necessary 
    if conf.needBenefit == True:
        monthStart = conf.monthStart
        benefitList() 
        print('Benefit list is create!')
    print('Done!')
