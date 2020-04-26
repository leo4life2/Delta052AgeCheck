import pandas as pd
from datetime import date, datetime
from tkinter import *
import tkinter.filedialog
import numpy as np

class ShotType:
    strict_margin = 7
    # In Days
    # Boundaries are *inclusive*
    # Bounds in format [Year, Month, Day]. Year, month and days will be summed.
    class HPV9:
        min_age = [16, 0, 0] # So age has to be larger than 15Y

        age_upper_bounds = [
            [26, 8, 0], # so age has to be smaller than 26Y+9M
            [26, 9, 0], # so age has to be smaller than 26Y+10M
            [26, 0, 0]  # so age has to be smaller than 27Y
        ]

    class HPV4:
        min_age = [20, 0, 0]

        age_upper_bounds = [
            [45, 6, 0],
            [45, 8, 0],
            [45, 0, 0]
        ]

    class Prev13:
        min_age = [0, 0, 6 * 7]

        age_upper_bounds = [
            [0, 6, 0],
            [0, 6, 0],
            [0, 6, 0],
            [1, 3, 0]
        ]

        reinforcement_minage = [1, 0, 0]

    class Five_Vacc:
        min_age = [0, 2, 0]

        age_upper_bounds = [
            [5, 0, 0],
            [5, 0, 0],
            [5, 0, 0],
            [5, 0, 0]
        ]

    class Infant_Flu:
        min_age = [0, 6, 0]

        age_upper_bounds = [
            [3, 0, 0], # Upper bound is 35 months, which is equivalent to <3 years.
            [3, 0, 0] # ! 科兴 one's upper limit is 3 Years--not sure if it means 3岁 or <3 years. !
        ]

    class Adult_Flu:
        min_age = [3, 0, 0]

        age_upper_bounds = [
            [99, 0, 0]
        ]

    class Rotavirus:
        min_age = [0, 0, 6 * 7]

        age_upper_bounds = [
            [2, 8, 0],
            [2, 8, 0],
            [2, 8, 0]
        ]

    class Chickenpox:
        min_age = [1, 0, 0]

        age_upper_bounds = [
            [99, 0, 0],
            [99, 0, 0],
            [99, 0, 0]
        ]

    class HandFootMouth:
        min_age = [0, 6, 0]

        age_upper_bounds = [
            [5, 0, 0]
        ]

    class JapEnceph:
        min_age = [0, 6, 0]

        age_upper_bounds = [
            [10, 0, 0],
            [10, 0, 0],
            [10, 0, 0],
            [10, 0, 0]
        ]

    class HepaA:
        min_age = [0, 12, 0]

        age_upper_bounds = [
            [17, 0, 0],
            [17, 0, 0],
            [17, 0, 0]
        ]

    class HepaB:
        min_age = [0, 0, 0]

        age_upper_bounds = [
            [99, 0, 0],
            [99, 0, 0],
            [99, 0, 0]
        ]

    class Meningitis:
        min_age = [0, 6, 0]

        age_upper_bounds = [
            [15, 0, 0],
            [15, 0, 0]
        ]

    class Measles:
        min_age = [0, 8, 0]

        age_upper_bounds = [
            [99, 0, 0],
            [99, 0, 0]
        ]

    name_strs = [
        'HPV四价',
        'HPV九价',
        '13价',
        '5联',
        '小儿流感',
        '成人流感',
        '五价轮状',
        '水痘',
        '手足口',
        '乙脑',
        '甲肝',
        '乙肝',
        '流脑'
    ]

    str_to_type = {
        'HPV四价': HPV4,
        'HPV九价': HPV9,
        '13价': Prev13,
        '5联': Five_Vacc,
        '小儿流感': Infant_Flu,
        '成人流感': Adult_Flu,
        '五价轮状': Rotavirus,
        '水痘': Chickenpox,
        '手足口': HandFootMouth,
        '乙脑': JapEnceph,
        '甲肝': HepaA,
        '乙肝': HepaB,
        '流脑': Meningitis
    }


def get_sheet():
    filepath = tkinter.filedialog.askopenfilename()
    sheet = pd.read_excel(filepath, na_filter = False)
    sheet_np = sheet.to_numpy()
    return filepath, sheet_np

def get_parameters(shot_date_raw, bday_raw, product_type, additional_info, shot_description):
    '''
    Returns encoded shot_date, bday & shotnum
    '''
    def get_shot_type(product_description):
        for shot_str in ShotType.name_strs:
            if shot_str in product_description:
                return ShotType.str_to_type[shot_str]

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
    shot_type = get_shot_type(shot_description)

    return shot_date, bday, shot_num, shot_type

def check_shot_age(shottype, shotnum, shot_date, bday):

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
    upperlimit = shottype.age_upper_bounds[shotnum-1]
    lowerlimit = shottype.min_age

    # Prev13 special case. Has different min age for 4th shot.
    if shotnum == 4 and shottype is ShotType.Prev13:
        lowerlimit = shottype.reinforcement_minage

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

            shot_date, bday, shot_num, shot_type = get_parameters(shot_date_raw, bday_raw, product_type, additional_info, product)

            if shot_type is None:
                if product not in ['',' ', '  ', '\n', '\r']:
                    unknown_types.append(product)
                continue

            name = row[3]

            row_str = 'Excel第' + str(i) + '行，姓名：' + name + '，预约时间：' + shot_date_raw + '，类型：' + product + '，第' + str(shot_num) + '针 ' + '，出生日期：' + bday_raw

            try:
                check_result = check_shot_age(shot_type, shot_num, shot_date, bday)
            except:
                print(row)
            if check_result != '通过':
                errors += (row_str + '\n' + check_result + '\n')

    unknown_types = set(unknown_types)
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

window = Tk()
window.title("自动检查疫苗年龄要求程序:D")
window.geometry('800x500')

path, np_sheet = None, None
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
#     txt['state'] = 'disabled'

lbl = Label(window, text="请先选择一个文件", anchor = 'center', wraplength = 300, font = ('Songti SC', 15))
lbl.grid(column = 0, row = 0)

choosebtn = Button(window, text="选择文件", command=choose_file)
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
