import pandas as pd
from datetime import date, datetime
from tkinter import *
import tkinter.filedialog
import numpy as np

def get_sheet(filepath = None):
    if not filepath:
        filepath = tkinter.filedialog.askopenfilename()
    sheet = pd.read_excel(filepath, na_filter = False)
    sheet_np = sheet.to_numpy()
    return filepath, sheet_np

path, bounds_sheet = get_sheet(filepath = 'bounds.xlsx')

class Shot:
    all_shots = []
    name_strs = []
    def __init__(self, name, minage, maxages, specialminage, specialminage_shotnum):
        self.name = name
        self.minage = minage
        self.maxages = maxages
        self.specialminage = specialminage
        self.specialminage_shotnum = specialminage_shotnum
        
    @classmethod
    def init_all_shot_bounds(cls, bounds_sheet):
    
        shots = []
        names = []

        for row in bounds_sheet:
            def list_str_to_list_int(list_str):
                lst = list_str.replace('[', '').replace(']', '').replace(' ', '').split(',')
                if lst == ['']:
                    return None
                lst = [int(s) for s in lst]
                return lst
            def process_specialminage(specialminage_str):
                shotnum = int(specialminage_str[0])
                stuffafter = specialminage_str[1:]
                stuffafter = stuffafter.replace(':', '').replace(' ', '')
                minage = list_str_to_list_int(stuffafter)
                return shotnum, minage
            '''
            ix: content
            0: vaccine name
            1: min age
            2: # shots
            3: shot 1 max age
            4: shot 2 max age
            5: shot 3 max age
            6: shot 4 max age
            7: special shot min age
            '''
            name, minage_str, nshots_str, maxages_str, specialminage_str = row[0], row[1], row[2], row[3:7], row[7]
            minage = list_str_to_list_int(minage_str)
            nshots = int(nshots_str)
            maxages = [list_str_to_list_int(maxage_str) for maxage_str in maxages_str]
            specialshotnum = None
            specialminage = None
            
            if specialminage_str != '':
                specialshotnum, specialminage = process_specialminage(specialminage_str)
                
            shots.append(Shot(name, minage, maxages, specialminage, specialshotnum))
            names.append(name)

        cls.all_shots = shots
        cls.name_strs = names

Shot.init_all_shot_bounds(bounds_sheet)

def get_parameters(shot_date_raw, bday_raw, product_type, additional_info, shot_description):
    '''
    Returns encoded shot_date, bday & shotnum
    '''
    def get_shot(product_description):
        for shot in Shot.all_shots:
            if shot.name in product_description:
                return shot

    def get_shot_num(product_str, additional_info):
        if '第一针' in product_str + additional_info or '单针' in product_str + additional_info:
            return 1
        elif '第二针' in product_str + additional_info:
            return 2
        elif '第三针' in product_str + additional_info:
            return 3
        elif '第四针' in product_str + additional_info:
            # Special case like reinforcement shot (aka 4th shot for PREV13)
            return 4
        else:
            return 1

    shot_date = datetime.strptime(shot_date_raw, '%Y-%m-%d')
    bday = datetime.strptime(bday_raw, '%Y/%m/%d')
    shot_num = get_shot_num(product_type, additional_info)
    shot = get_shot(shot_description)

    return shot_date, bday, shot_num, shot

def check_shot_age(shot, shotnum, shot_date, bday):

    def get_limit_date(bday, limit, upper = False, lower = False):
        if lower:
            return bday + pd.DateOffset(years=limit[0], months=limit[1], days=limit[2] - 1)
        if upper:
            last_nonzero_ix = np.max(np.nonzero(limit))
            inclusive_limit = limit[:]
            inclusive_limit[last_nonzero_ix] += 1

            return bday + pd.DateOffset(years=inclusive_limit[0], months=inclusive_limit[1], days=inclusive_limit[2] - 1)

    def does_satisfy_age_limit(shotdate, bday, limit, check_lowerbound = False):
        if check_lowerbound:
            lowerlimit_date = get_limit_date(bday, limit, lower = True)
            return not shotdate.date() < lowerlimit_date
        else:
            upperlimit_date = get_limit_date(bday, limit, upper = True)
            return not shotdate.date() > upperlimit_date

    upperlimit = -1
    lowerlimit = -1

    # Shot type is never None.
    # Kinda gambling here, betting that noone's gonna make a mistake and cause index out of bound error.
    upperlimit = shot.maxages[shotnum-1]
    lowerlimit = shot.minage

    # i.e. Prev13 special case. Has different min age for 4th shot.
    if shot.specialminage and shotnum == shot.specialminage_shotnum:
        lowerlimit = shot.specialminage

    # Check lower boundaries first
    if not does_satisfy_age_limit(shot_date, bday, lowerlimit, check_lowerbound = True):
        return '！警告：年龄过小。出生日期：' + str(bday) +'，下限：' + str(get_limit_date(bday, lowerlimit, lower = True)) + '\n'

    # Check strict & lenient upper bounds
    if not does_satisfy_age_limit(shot_date, bday, upperlimit):
        return '！警告：年龄过大。出生日期：' + str(bday) +'，上限：' + str(get_limit_date(bday, upperlimit, upper = True)) + '\n'

    return '通过'

def check_sheet_vacc_ages(sheet):

    to_display = ''
    unknown_types = []
    errors = ''

    i = 4
    for row in sheet[3:]: # First 3 rows excluded (备注 counted as label for pandas)
        i += 1
        # *Col index 1 is shot date
        # *Col index 7 is patient bday
        # Col index 13 is status. Exclude 取消 status data.
        # Col index 16 is additional info (look here if shot # not indicated in col ix 17)
        # Col index 17 is 就诊类型（疫苗 -- vaccine shot num appears here
        # Col index 18 is 就诊原因（套餐/券 -- vaccine type appears here
        # * = Required. Skip if missing.

        status = row[13]
        product = row[18]

        if product is not np.nan and status != '取消' and row[1] != '' and row[7] != '':

            shot_date_raw = row[1][:10] # Only need date; exclude hours and mins
            bday_raw = row[7]
            product_type = row[17]
            additional_info = row[16]

            shot_date, bday, shot_num, shot = get_parameters(shot_date_raw, bday_raw, product_type, additional_info, product)

            if shot is None:
                if product not in ['',' ', '  ', '\n', '\r']:
                    unknown_types.append(product)
                continue

            name = row[3]

            row_str = 'Excel第' + str(i) + '行，姓名：' + name + '，预约时间：' + shot_date_raw + '，类型：' + product + '，第' + str(shot_num) + '针 ' + '，出生日期：' + bday_raw

            try:
                check_result = check_shot_age(shot, shot_num, shot_date, bday)
            except:
                print(row)
            if check_result != '通过':
                errors += (row_str + '\n' + check_result + '\n')

    unknown_types = set(unknown_types)
    if len(unknown_types) > 0:
        to_display += '发现了未知的套餐类型：' + '\n\n'
        for uktype in unknown_types:
            to_display += str(uktype) + '\n'
        to_display += '\n'

    if errors == '':
        to_display += '年龄检查全部通过'
    else:
        to_display += '发现有超龄的人员' + '\n\n'
        to_display += errors

    return to_display

def choose_file():
    global path
    global np_sheet
    path, np_sheet = get_sheet()
    lbl.configure(text='目前文件: '+ path)

def analyze():
    txt['state'] = 'normal'
    txt.delete(1.0, END)
    if path is None:
        return
    results = check_sheet_vacc_ages(np_sheet)
    txt.insert(1.0, results)

window = Tk()
window.title("自动检查疫苗年龄要求程序:D")
window.geometry('800x500')

path, np_sheet = None, None

lbl = Label(window, text="请先选择一个文件", anchor = 'center', wraplength = 300, font = ('Songti SC', 15))
lbl.grid(column = 0, row = 0)

choosebtn = Button(window, text="选择文件", command = choose_file)
choosebtn.grid(column = 0, row = 1)

analyzebtn = Button(window, text='开始检查', command = analyze)
analyzebtn.grid(column = 0, row = 2)

scroll = Scrollbar(window)
scroll.grid(column = 1, row = 3, sticky = 'ns')

txt = Text(window, height=15, width=95)
txt.grid(column = 0, row = 3)

scroll.config(command=txt.yview)
txt.config(yscrollcommand=scroll.set)

txt.insert(END, '结果会在这里显示')
txt.grid_propagate(False)
txt.config(font = ('Songti SC', 15))

window.mainloop()


