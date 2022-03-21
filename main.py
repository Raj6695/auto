from flask import Flask

import os

import pandas as pd
import openpyxl
import random
from pandas import DataFrame
import numpy as np
from openpyxl import Workbook, load_workbook
import string
import re
import datetime 
import dateutil.parser as dparser
import math
import os
import itertools
from datetime import timedelta

import warnings

def work_allocation():
    print("begun")
 
    warnings.filterwarnings("ignore")

    x_path             = "//10.9.57.54/Recepting-&-Closure/work_allocation/"

    files              = os.listdir(x_path)
    pending_sr         = pd.read_excel(f"{x_path}/PENDING_SR.xlsx")


    cdnc               = pd.read_excel(f"{x_path}CDNC.xlsx")
    error              = pd.read_excel(f"{x_path}ERROR.xlsx")
    gecl               = pd.read_excel(f"{x_path}GECL_LINK.xlsx", sheet_name = "GECL_Loan")
    link               = pd.read_excel(f"{x_path}GECL_LINK.xlsx", sheet_name = "Link_Loan")
    overall            = pd.read_excel(f"{x_path}OVERALL_HOLD.xlsx")
    #Settlement         = pd.read_excel(f"{x_path/Settlement not updated night.xlsx")
    tamil_nadu         = pd.read_excel(f"{x_path}TAMIL_NADU_ONLINE.xlsx")
    user_id            = pd.read_excel(f"{x_path}USER_ID_MASTER.xlsx")
    vf_branch          = pd.read_excel(f"{x_path}VF_BRANCH.xlsx")
    hyper              = pd.read_excel(f"{x_path}HYPER_BRANCH.xlsx")


    #################################################################################################################################

    pending_sr.columns = pending_sr.columns.str.upper()
    gecl.columns       = gecl.columns.str.upper()
    link.columns       = link.columns.str.upper()
    cdnc.columns       = cdnc.columns.str.upper()
    overall.columns    = overall.columns.str.upper()
    tamil_nadu.columns = tamil_nadu.columns.str.upper()
    error.columns      = error.columns.str.upper()
    hyper.columns      = hyper.columns.str.upper()

    list(pending_sr['REQUEST CATEGORY'].drop_duplicates())



    pending_sr               = pending_sr.rename(columns       = {'BRANCH NAME ':'BRANCH NAME'})
    pending_sr               = pending_sr.rename(columns       = {'AGREEMENT STATUS ':'AGREEMENT STATUS'})
    pending_sr               = pending_sr.rename(columns       = {'DEPOSIT STATUS FLAG ':'DEPOSIT STATUS FLAG'})

    req_list =['Express NOC','Refinance Closure only','VF Closure & NOC','VF Closure only','VF NOC Only']

    pending_sr = pending_sr[pending_sr["REQUEST CATEGORY"].isin(req_list)]
    pending_sr

    pending_sr.loc[pending_sr["STATUS"]=='Work In Progress','STATUS'] = 'Initiated'

    pending_sr = pending_sr[pending_sr["AGREEMENT STATUS"]=="A"]

    pending_sr

    list(pending_sr['STATUS'].drop_duplicates())

    list(pending_sr['AGREEMENT STATUS'].drop_duplicates())

    list(pending_sr)

    # pending_sr[pending_sr["STATUS"]!='C']

    list(pending_sr["STATUS"].drop_duplicates())

    list(pending_sr["AGREEMENT STATUS"].drop_duplicates())

    pending_sr = pending_sr[pending_sr["DEPOSIT STATUS FLAG"]!="Y"]
    pending_sr

    pending_sr["CREATION DATE"] = pending_sr["CREATION DATE"].astype(str)
    pending_sr["CREATION TIME"] = pending_sr["CREATION TIME"].astype(str)
    pending_sr["CREATION_date_time"] = pending_sr["CREATION DATE"] +" "+ pending_sr["CREATION TIME"]
    pending_sr["CREATION_date_time"]

    pending_sr["CREATION_date_time"] = pd.to_datetime(pending_sr["CREATION_date_time"], format='%Y-%m-%d %H:%M:%S') 

    pending_sr

    VF_CLOSURE = pending_sr[pending_sr["REQUEST CATEGORY"].isin(['VF Closure only','Refinance Closure only',"Closure Only"])]
    OTHERS     = pending_sr[~pending_sr["REQUEST CATEGORY"].isin(['VF Closure only','Refinance Closure only',"Closure Only"])]

    VF_CLOSURE
    VF_CLOSURE["EXPIRY"] = VF_CLOSURE["CREATION_date_time"] + np.timedelta64(4,'h')

    OTHERS
    OTHERS["EXPIRY"] = OTHERS["CREATION_date_time"] + np.timedelta64(8,'h')

    # zz = OTHERS[["CREATION_date_time","EXPIRY"]

    # datetimes = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    # datetimes = pd.to_datetime(datetimes)
    # datetimes

    # zz["current"] = datetimes

    # zz["tat"] =    (zz["EXPIRY"]-datetimes).astype('timedelta64[h]')

    # import pandas
    # df = pandas.DataFrame(columns=['to','fr','ans'])
    # df.to = [pandas.Timestamp('2014-01-24 13:03:12.050000'), pandas.Timestamp('2014-01-27 11:57:18.240000'), pandas.Timestamp('2014-01-23 10:07:47.660000')]
    # df.fr = [pandas.Timestamp('2014-01-26 23:41:21.870000'), pandas.Timestamp('2014-01-27 15:38:22.540000'), pandas.Timestamp('2014-01-23 18:50:41.420000')]
    # (df.fr-df.to).astype('timedelta64[h]')


    pending_sr = pd.concat([VF_CLOSURE,OTHERS])



    # zz.tail(60)

    # WORK_ALLOTED["DUE TIME"] =    (WORK_ALLOTED["EXPIRY"] -datetime.datetime.now()).astype('timedelta64[h]')


    # WORK_ALLOTED


    pending_sr

    list(pending_sr)

    # pending_sr =pending_sr.drop(columns = "CREATION_date_time" )

    pending_sr[["CREATION DATE","CREATION TIME","EXPIRY"]]


    gecl               = gecl.rename(columns       = {'AGREEEMENT NO':'AGREEMENT NO.'})
    cdnc               = cdnc.rename(columns       = {'AGREEMENT NUM':'AGREEMENT NO.'})
    overall            = overall.rename(columns    = {'AGREEMENTNO':'AGREEMENT NO.'})
    tamil_nadu         = tamil_nadu.rename(columns = {'ROW LABELS':'BRANCH NAME'})
    tamil_nadu

    pending_sr = pd.concat([pending_sr,pd.DataFrame(columns=["ALLOCATED","GECL","CDNC","LINK","OVERALL","TAMIL", "ALLOCATED_TO"])])

    PEND       = pending_sr
    PEND

    list(tamil_nadu)

    tamil_nadu["ROW LABELS"]  = tamil_nadu['BRANCH NAME']
    hyper['BRANCH NAME']      =  hyper["AREA BRANCHES"]
    tamil_nadu = tamil_nadu[["BRANCH NAME","ROW LABELS"]]
    list(tamil_nadu)

    PEND = pd.merge( PEND, tamil_nadu, on ="BRANCH NAME", how = "left" )
    TAMIL_NADU = PEND[PEND["ROW LABELS"].notnull()]
    PEND["TAMIL"] = PEND["ROW LABELS"]
    # PEND

    PEND  = pd.merge( PEND, hyper[['BRANCH NAME',"AREA BRANCHES"]], on ="BRANCH NAME", how = "left" )
    HYPER = PEND[PEND["AREA BRANCHES"].notnull()]

    PEND  = PEND.drop_duplicates()
    PEND1 = PEND

    ERROR      = pd.merge(PEND["AGREEMENT NO."] , error["AGREEMENT NO."], on ="AGREEMENT NO.", how = "inner")
    ERROR       = PEND[PEND["AGREEMENT NO."].isin(list(ERROR["AGREEMENT NO."]))]
    ERROR["ERROR"] = ERROR["AGREEMENT NO."]
    ERROR["ALLOCATED"] = "ERROR"
    ERROR

    PEND       = PEND[~PEND["AGREEMENT NO."].isin(list(ERROR["AGREEMENT NO."]))]
    GECL       = pd.merge(PEND["AGREEMENT NO."] , gecl["AGREEMENT NO."], on ="AGREEMENT NO.", how = "inner")
    GECL       = PEND[PEND["AGREEMENT NO."].isin(list(GECL["AGREEMENT NO."]))]
    GECL["GECL"] = GECL["AGREEMENT NO."]
    GECL["ALLOCATED"] = "GECL"
    GECL

    PEND       = PEND[~PEND["AGREEMENT NO."].isin(list(GECL["AGREEMENT NO."]))]
    CDNC       = pd.merge(PEND["AGREEMENT NO."] , cdnc["AGREEMENT NO."], on ="AGREEMENT NO.", how = "inner")
    CDNC       = PEND[PEND["AGREEMENT NO."].isin(list(CDNC["AGREEMENT NO."]))]
    CDNC["CDNC"] = CDNC["AGREEMENT NO."]
    CDNC["ALLOCATED"] = "CDNC"
    CDNC

    PEND       = PEND[~PEND["AGREEMENT NO."].isin(list(CDNC["AGREEMENT NO."]))]
    LINK       = pd.merge(PEND["AGREEMENT NO."] , link["AGREEMENT NO."], on ="AGREEMENT NO.", how = "inner")
    LINK       = PEND[PEND["AGREEMENT NO."].isin(list(LINK["AGREEMENT NO."]))]
    LINK["LINK"] = LINK["AGREEMENT NO."]
    LINK["ALLOCATED"] = "LINK"
    # LINK

    PEND       = PEND[~PEND["AGREEMENT NO."].isin(list(LINK["AGREEMENT NO."]))]
    OVERALL    = pd.merge(PEND["AGREEMENT NO."], overall["AGREEMENT NO."], on ="AGREEMENT NO.", how = "inner")
    OVERALL    = PEND[PEND["AGREEMENT NO."].isin(list(OVERALL["AGREEMENT NO."]))]
    OVERALL["OVERALL"] = OVERALL["AGREEMENT NO."]
    OVERALL["ALLOCATED"] = "OVERALL"
    # OVERALL

    GECL_CDNC_LINK_OVERAL= pd.concat([ERROR,GECL,CDNC,LINK,OVERALL])
    PEND       = PEND[~PEND["AGREEMENT NO."].isin(list(GECL_CDNC_LINK_OVERAL["AGREEMENT NO."]))]
    PEND       = pd.concat([PEND,GECL_CDNC_LINK_OVERAL,ERROR])
    PEND       = PEND.drop_duplicates()

    PEND[PEND["ERROR"].notnull()]

    vf_branch

    vf_branch_PLOT = vf_branch[["MIS Branch","Branch","Area", "Region","Zone"]]
    vf_branch_PLOT = vf_branch_PLOT.rename(columns={'MIS Branch':"BRANCH NAME"})
    PEND           = pd.merge(PEND , vf_branch_PLOT, on ="BRANCH NAME", how = "left")

    list(PEND)

    #### Remove and sort overall and cdnc cases

    cdnc = PEND[PEND["CDNC"].notnull()]
    PEND = PEND[PEND["CDNC"].isnull()]
    cdnc["ALLOCATED_TO"] = "CDNC"
    PEND = PEND[PEND["OVERALL"].isnull()]


    PEND[PEND["ERROR"].notnull()]

    ###### ALLOCATE AG 

    PEND = PEND[PEND["ALLOCATED"]!="ERROR"]
    PEND



    PEND_initiated = PEND[PEND["STATUS"]=="Initiated"]
    PEND_opened    = PEND[PEND["STATUS"]=="Open"]
    # PEND_initiated["ALLOCATED_TO"] = None

    PEND_initiated
    PEND_opened



    # select random employees
    user_id = user_id[user_id["Shift"]=="DAY"]

    # user_id



    ### PEND INITIATED

    rem = abs(len(PEND_initiated)%len(user_id))
    div = round((len(PEND_initiated)-rem)/len(user_id))

    div

    PEND_initiated

    if rem>0:
        selected_PEND_initiated = PEND_initiated[:-rem]
    else:  
        selected_PEND_initiated = PEND_initiated

    selected_PEND_initiated    

    if rem>0:
        leftouts_ini = selected_PEND_initiated[-rem:]
    else:  
         leftouts_ini = pd.DataFrame(columns= list(selected_PEND_initiated))

    leftouts_ini

    # distribute it equally to the employees by mapping df and emp 
    PEND_initiated_list = []
    for index, name in enumerate(list(user_id["Employee Name"])):
        df = selected_PEND_initiated[:div]
        PEND_initiated_list.append([name,df])
        selected_PEND_initiated = selected_PEND_initiated.drop(df.index[:div])

    selected_PEND_initiated

    PEND_initiated_list

    ### PEND_opened

    rem1 = abs(len(PEND_opened)%len(user_id))

    div1 = round((len(PEND_opened)-rem)/len(user_id))

    PEND_opened

    div1

    rem1

    if rem1>0:
        selected_PEND_opened = PEND_opened[:-rem1]
    else:  
        selected_PEND_opened = PEND_opened
    selected_PEND_opened    

    if rem1>0:
        leftouts_op = selected_PEND_opened[-rem1:]
    else:  
         leftouts_op = pd.DataFrame(columns= list(selected_PEND_opened))

    leftouts_op

    # selected_PEND_opened = PEND_opened[:-rem1]
    # selected_PEND_opened
    # leftouts_op = selected_PEND_opened[-rem1:]

    # distribute it equally to the employees by mapping df and emp 
    PEND_opened_list = []
    for index, name in enumerate(list(user_id["Employee Name"])):
        df = selected_PEND_opened[:div1]
        PEND_opened_list.append(df)
        selected_PEND_opened = selected_PEND_opened.drop(df.index[:div1])
    # selected_PEND_opened

    ### LEFTOVER

    all_leftouts = pd.concat([leftouts_ini,leftouts_op])
    rand_user = random.choice(list(user_id["Employee Name"]))
    all_leftouts["ALLOCATED_TO"] = rand_user

    ## Merging all dataframes in list to one large dataframe

    ini_op =[]
    for ini, op in zip(PEND_initiated_list, PEND_opened_list):
        inif = pd.concat([ini[1],op])
        inif["ALLOCATED_TO"] = ini[0]
        ini_op.append(inif)

    ini_op = pd.concat(ini_op)

    #### Merging leftouts with large dataframe

    WORK_ALLOTED = pd.concat([ini_op, all_leftouts])

    WORK_ALLOTED.loc[WORK_ALLOTED["LINK"].notnull(),'REQUEST CATEGORY'] ="Closure Only"
    WORK_ALLOTED.loc[WORK_ALLOTED["CDNC"].notnull(),'REQUEST CATEGORY'] ="Closure Only"

    WORK_ALLOTED[WORK_ALLOTED["LINK"].notnull()]

    WORK_ALLOTED = pd.concat([WORK_ALLOTED,cdnc])

    WORK_ALLOTED.to_csv(f"{x_path}work_allocate_output.csv")

    WORK_ALLOTED["DUE TIME"] =    (WORK_ALLOTED["EXPIRY"] - datetime.datetime.now()).astype('timedelta64[h]')

    # WORK_ALLOTED["Creation"]

    WORK_ALLOTED.to_csv("//10.9.57.54/Recepting-&-Closure/work_allocation/work_allocate_output.csv")
    
    print("allocated")




from flask import Flask, request, render_template_string
import datetime

app = Flask(__name__)

@app.route('/')
def index():
    return render_template_string('''
<button onclick="my_function()">Get Time</button>
<span id="time">Press Button to see current time on server.</span>
<script>
span_time = document.querySelector("#time");
function my_function(){
   fetch('/r_n_c')
}
</script>
''')

@app.route('/get_time')
def time():
    print(datetime.datetime.now().strftime('%Y.%m.%d %H:%M.%S'))
    
@app.route('/add')
def add():
    z=1+2
    print(z)
@app.route('/r_n_c')
def r_n_c():
    return work_allocation()
    
if __name__ == '__main__':
    #app.debug = True
    app.run() #debug=True 

def add():
    z=1+2
    print(z)
