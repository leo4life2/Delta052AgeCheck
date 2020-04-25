from pandas import read_excel
from datetime import datetime
from tkinter import *
from tkinter import scrolledtext
import tkinter.filedialog
import numpy as np

def get_sheet():
    filepath = tkinter.filedialog.askopenfilename()
    sheet = read_excel(filepath, na_filter = False)
    sheet_np = sheet.to_numpy()
    return filepath, sheet_np

class ShotType:
    # In Days
    # Boundaries are *inclusive* so there are '-1's after the year #.
    class HPV9:
        min_age = 16 * 365.25

        threeshots_age_upper_bounds_lenient = [
            (27 * 365.25 - 1) - (4 * 30),
            (27 * 365.25 - 1) - (3 * 30),
            (27 * 365.25 - 1)
        ]
        threeshots_age_upper_bounds_strict = [bound - 7 for bound in threeshots_age_upper_bounds_lenient]

    class HPV4:
        min_age = 20 * 365.25

        threeshots_age_upper_bounds_lenient = [
            (46 * 365.25 - 1) - (6 * 30),
            (46 * 365.25 - 1) - (4 * 30),
            (46 * 365.25 - 1)
        ]
        threeshots_age_upper_bounds_strict = [bound - 7 for bound in threeshots_age_upper_bounds_lenient]

    class PREV13:
        min_age = 6 * 7

        threeshots_age_upper_bounds_lenient = [
            (7 * 30),
            (7 * 30),
            (7 * 30)
        ]
        threeshots_age_upper_bounds_strict = [bound - 7 for bound in threeshots_age_upper_bounds_lenient]

        reinforcementshot_minage = (12 * 30)
        reinforcementshot_maxage_lenient = (16 * 30)
        reinforcementshot_maxage_strict = reinforcementshot_maxage_lenient - 7

def get_parameters(shot_date_raw, bday_raw, product_type, additional_info, shot_description):
    '''
    Returns encoded shot_date, bday & shotnum
    '''
    def get_shot_type(product_description):
        name_strs = [
            'HPV四价',
            'HPV九价',
            '13价'
        ]

        str_to_type = {
            'HPV四价': ShotType.HPV4,
            'HPV九价': ShotType.HPV9,
            '13价': ShotType.PREV13
        }

        for shot_str in name_strs:
            if shot_str in product_description:
                return str_to_type[shot_str]

    def get_shot_num(product_str, additional_info):
        if '第一针' in product_str or '第一针' in additional_info:
            return 1
        elif '第二针' in product_str or '第二针' in additional_info:
            return 2
        elif '第三针' in product_str or '第三针' in additional_info:
            return 3
        else:
            # Special case like reinforcement shot (aka 4th shot for PREV13)
            return 4

    shot_date = datetime.strptime(shot_date_raw, '%Y-%m-%d')
    bday = datetime.strptime(bday_raw, '%Y/%m/%d')
    shot_num = get_shot_num(product_type, additional_info)
    shot_type = get_shot_type(shot_description)

    return shot_date, bday, shot_num, shot_type

def check_shot_age(shottype, shotnum, shot_date, bday):

    def datediff(d1, d2):
        return (d2 - d1).days

    def does_satisfy_age_limit(shotdate, bday, age_boundary, check_lowerbound = False):
        if check_lowerbound:
            return not datediff(bday, shotdate) <= age_boundary
        else:
            return not datediff(bday, shotdate) >= age_boundary

    upperlimit_strict = -1
    upperlimit_lenient = -1

    # lowerlimit is SAME for both Strict & Lenient standards.
    lowerlimit = -1

    if shottype is ShotType.HPV9 or shottype is ShotType.HPV4 or shottype is ShotType.PREV13 and shotnum in range(1,4):
        # regular 3-shot vaccine
        upperlimit_strict = shottype.threeshots_age_upper_bounds_strict[shotnum-1]
        upperlimit_lenient = shottype.threeshots_age_upper_bounds_lenient[shotnum-1]
        lowerlimit = shottype.min_age
    elif shottype is ShotType.PREV13 and shotnum not in range(3):
        # reinforcement shots for PREV13
        upperlimit_strict = shottype.reinforcementshot_maxage_strict
        upperlimit_lenient = shottype.reinforcementshot_maxage_lenient
        lowerlimit = shottype.reinforcementshot_minage

    # Check lower boundaries first
    if not does_satisfy_age_limit(shot_date, bday, lowerlimit, check_lowerbound = True):
        return '！警告：年龄过小。年龄（天）：' + str(datediff(bday, shot_date)) +'，下限：' + str(lowerlimit)

    # Check strict & lenient upper bounds
    if not does_satisfy_age_limit(shot_date, bday, upperlimit_strict):
        return '！警告：年龄过大。年龄（天）：' + str(datediff(bday, shot_date)) +'，上限：' + str(upperlimit_strict)
    if not does_satisfy_age_limit(shot_date, bday, upperlimit_lenient):
        return '提醒：即将超龄。年龄（天）：' + str(datediff(bday, shot_date)) +'，上限：' + str(upperlimit_strict) # Using strict bound for display

    return '通过'

def check_sheet_vacc_ages(sheet):

    to_display = ''
    unknown_types = []
    errors = ''

    for row in sheet[3:]: # First 3 rows excluded (备注 counted as label for pandas)
        # Col index 1 is shot date
        # Col index 7 is patient bday
        # Col index 13 is status. Exclude 取消 status data.
        # Col index 16 is additional info (look here if shot # not indicated in col ix 17)
        # Col index 17 is 就诊类型（疫苗 -- vaccine shot num appears here
        # Col index 18 is 就诊原因（套餐/券 -- vaccine type appears here
        status = row[13]
        product = row[18]

        if product is not np.nan and status != '取消':

            shot_date_raw = row[1][:10] # Only need date; exclude hours and mins
            bday_raw = row[7]
            product_type = row[17]
            additional_info = row[16]

            shot_date, bday, shot_num, shot_type = get_parameters(shot_date_raw, bday_raw, product_type, additional_info, product)

            if shot_type is None and product != '':
                unknown_types.append(product)
                continue

            name = row[3]

            row_str = '姓名：' + name + '，预约时间：' + shot_date_raw + '，类型：' + product + '，第' + str(shot_num) + '针 ' + '，出生日期：' + bday_raw

            check_result = check_shot_age(shot_type, shot_num, shot_date, bday)
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
        to_display += '发现有超龄/即将超龄的人员' + '\n\n'
        to_display += errors

    return to_display

window = Tk()
window.title("自动检查疫苗年龄要求程序:D")
window.geometry('400x500')

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
    txt['state'] = 'disabled'


lbl = Label(window, text="请先选择一个文件", anchor = 'center', wraplength = 300, font = ('Songti SC', 15))
lbl.grid(column = 0, row = 0)

choosebtn = Button(window, text="选择文件", command=choose_file)
choosebtn.grid(column = 0, row = 1)

analyzebtn = Button(window, text='开始检查', command = analyze)
analyzebtn.grid(column = 0, row = 2)

scroll = Scrollbar(window)
scroll.grid(column = 1, row = 3, sticky = 'ns')

txt = Text(window, height=15, width=45)
txt.grid(column = 0, row = 3)

scroll.config(command=txt.yview)
txt.config(yscrollcommand=scroll.set)

txt.insert(END, '结果会在这里显示')
txt.grid_propagate(False)
txt.config(font = ('Songti SC', 15))

window.mainloop()
