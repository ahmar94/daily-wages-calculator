# Start
import numpy as np
import matplotlib.pyplot as plt
import math as ma
import xlwings as xw
import time
from pathlib import Path



def wages_write():
    # Open the existing workbook
    current_folder = Path(__file__).resolve().parent
    workbook_path = current_folder / "Daily Wages Calculator.xlsm"
    wb = xw.Book(workbook_path)
    sheet = wb.sheets[0]
    sheet1 = wb.sheets[2]
    sr_num  = sheet.range('B14:B31').value
    st_time = sheet.range('G14:G31').value
    en_time = sheet.range('J14:J31').value
    st_date = sheet.range('H14:H31').value
    en_date = sheet.range('K14:K31').value
    st_day  = sheet.range('I14:I31').value
    en_day  = sheet.range('L14:L31').value
    wages   = sheet.range('X14:X31').value
    TtEO    = sheet.range('S14:S31').value
    TtEB    = sheet.range('U14:U31').value
    sr_num_clean  =  [int(x) for x in sr_num if str(x).strip() != ""                 ]
    st_clean_time =  [x for x in st_time if x is not None                            ]
    en_clean_time =  [x for x in en_time if x is not None                            ]   
    st_clean_dates = [f"{x.day}/{x.month}/{x.year}" for x in st_date if x is not None]
    en_clean_dates = [f"{x.day}/{x.month}/{x.year}" for x in en_date if x is not None]
    st_clean_day =   [str(x) for x in st_day if str(x).strip() != ""                 ]
    en_clean_day =   [str(x) for x in en_day if str(x).strip() != ""                 ]
    TtEO_clean =     [str(x) for x in TtEO if str(x).strip() != ""                   ]
    TtEB_clean =     [str(x) for x in TtEB if str(x).strip() != ""                   ]
    wages_clean =    [x for x in wages if x is not None                              ]
    print(sr_num_clean)
    print(st_clean_dates) 
    print(en_clean_dates)
    print(st_clean_day)
    print(en_clean_day)
    START_TIME = []
    END_TIME = []
    S_P = []
    E_P = []
    for i in range(len(st_clean_time)):
        st_Time = float(np.round(st_clean_time[i] * 24,2))
        if st_Time >= 12:
            st_Time_p = f'{np.round(st_Time - 12, 2)} PM'
        else:
            st_Time_p = f'{np.round(st_Time, 2)} AM' 
        S_P.append(st_Time_p)
        en_Time = float(np.round(en_clean_time[i] * 24, 2))
        if en_Time >= 12:
            en_Time_p = f'{np.round(en_Time - 12, 2)} PM'
        else:
            en_Time_p = f'{np.round(en_Time, 2)} AM'
        E_P.append(en_Time_p)
        START_TIME.append(st_Time)
        END_TIME.append(en_Time)
    print(START_TIME)
    print(END_TIME)
    print(S_P)
    print(E_P)
    
    TOTAL_TIME = []
    if st_clean_dates[0] == en_clean_dates[0]:
        print("Yes")
    elif st_clean_dates[0] < en_clean_dates[0]:
        print("No")
    for i in range(len(sr_num_clean)):
        if st_clean_dates[i] == en_clean_dates[i]:
            END_TIME[i] = END_TIME[i]
            total_time = END_TIME[i] - START_TIME[i]
            if START_TIME[i] >=0 and END_TIME[i] <= 8:
                t_time = f'{total_time} hrs'
                if TtEB_clean[i] == "Yes":
                    dehari = wages_clean[i] * (total_time - 1)
                else:
                    dehari = wages_clean[i] * total_time
        elif st_clean_dates[i] < en_clean_dates[i] and END_TIME[i] >= START_TIME[i]:
            END_TIME[i] = END_TIME[i] + 24
            total_time = END_TIME[i] - START_TIME[i]
        elif st_clean_dates[i] < en_clean_dates[i] and START_TIME[i] > END_TIME[i]:
            total_time = 24- (START_TIME[i] - END_TIME[i])               
        TOTAL_TIME.append(total_time)
    print(TOTAL_TIME)
    Mon_s = 'F'
    Mon_e = 'G'
    Mon_w = 'H'
    Tue_s = 'I'
    Tue_e = 'J'
    Tue_w = 'K'
    Wed_s = 'L'
    Wed_e = 'M'
    Wed_w = 'N'
    Thu_s = 'O'
    Thu_e = 'P'
    Thu_w = 'Q'
    Fri_s = 'R'
    Fri_e = 'S'
    Fri_w = 'T'
    Sat_s = 'U'
    Sat_e = 'V'
    Sat_w = 'W'    
    Sun_s = 'X'
    Sun_e = 'Y'
    Sun_w = 'Z'
    cleared_mon = False
    cleared_tue = False
    cleared_wed = False
    cleared_thu = False
    cleared_fri = False
    cleared_sat = False
    cleared_sun = False
    for i in range(len(sr_num)):
        sheet1.range(f'C{i + 10}').value = sr_num[i]
        if st_time[i] is None:
            st_time[i] = 0
        total_minutes = round(st_time[i] * 24 * 60)

        hours = total_minutes // 60
        minutes = total_minutes % 60

        ap = "AM" if hours < 12 else "PM"

        display_hour = hours if hours <= 12 else hours - 12
        if display_hour == 0:
            display_hour = 12
        ts = f'{display_hour}:{minutes:02d} {ap}'
        if en_time[i] is None:
            en_time[i] = 0
        total_minutes = round(en_time[i] * 24 * 60)

        hours = total_minutes // 60
        minutes = total_minutes % 60

        ap = "AM" if hours < 12 else "PM"

        display_hour = hours if hours <= 12 else hours - 12
        if display_hour == 0:
            display_hour = 12

        te = f'{display_hour}:{minutes:02d} {ap}'
        if st_day[i] == "Monday":
            if not cleared_mon:
                sheet1.range(f'{Mon_s}10:{Mon_s}27').value = ""
                sheet1.range(f'{Mon_w}10:{Mon_w}27').value = ""
                sheet1.range(f'{Mon_e}10:{Mon_e}27').value = ""
                cleared_mon = True
            sheet1.range(f'{Mon_s}{i + 10}').value = ts
            sheet1.range(f'{Mon_w}{i + 10}').value = wages[i]
            sheet1.range(f'{Mon_e}{i + 10}').value = te
        elif st_day[i] == "Tuesday":
            if not cleared_tue:
                sheet1.range(f'{Tue_s}10:{Tue_s}27').value = ""
                sheet1.range(f'{Tue_w}10:{Tue_w}27').value = ""
                sheet1.range(f'{Tue_e}10:{Tue_e}27').value = ""
                cleared_tue = True
            sheet1.range(f'{Tue_s}{i + 10}').value = ts
            sheet1.range(f'{Tue_w}{i + 10}').value = wages[i]
            sheet1.range(f'{Tue_e}{i + 10}').value = te
        elif st_day[i] == "Wednesday":
            if not cleared_wed:
                sheet1.range(f'{Wed_s}10:{Wed_s}27').value = ""
                sheet1.range(f'{Wed_w}10:{Wed_w}27').value = ""
                sheet1.range(f'{Wed_e}10:{Wed_e}27').value = ""
                cleared_wed = True
            sheet1.range(f'{Wed_s}{i + 10}').value = ts
            sheet1.range(f'{Wed_w}{i + 10}').value = wages[i]
            sheet1.range(f'{Wed_e}{i + 10}').value = te
        elif st_day[i] == "Thursday":
            if not cleared_thu:
                sheet1.range(f'{Thu_s}10:{Thu_s}27').value = ""
                sheet1.range(f'{Thu_w}10:{Thu_w}27').value = ""
                sheet1.range(f'{Thu_e}10:{Thu_e}27').value = ""
                cleared_thu = True
            sheet1.range(f'{Thu_s}{i + 10}').value = ts
            sheet1.range(f'{Thu_w}{i + 10}').value = wages[i]
            sheet1.range(f'{Thu_e}{i + 10}').value = te
        elif st_day[i] == "Friday":
            if not cleared_fri:
                sheet1.range(f'{Fri_s}10:{Fri_s}27').value = ""
                sheet1.range(f'{Fri_w}10:{Fri_w}27').value = ""
                sheet1.range(f'{Fri_e}10:{Fri_e}27').value = ""
                cleared_fri = True
            sheet1.range(f'{Fri_s}{i + 10}').value = ts
            sheet1.range(f'{Fri_w}{i + 10}').value = wages[i]
            sheet1.range(f'{Fri_e}{i + 10}').value = te
        elif st_day[i] == "Saturday":
            if not cleared_sat:
                sheet1.range(f'{Sat_s}10:{Sat_s}27').value = ""
                sheet1.range(f'{Sat_w}10:{Sat_w}27').value = ""
                sheet1.range(f'{Sat_e}10:{Sat_e}27').value = ""
                cleared_sat = True
            sheet1.range(f'{Sat_s}{i + 10}').value = ts
            sheet1.range(f'{Sat_w}{i + 10}').value = wages[i]
            sheet1.range(f'{Sat_e}{i + 10}').value = te
        elif st_day[i] == "Sunday":
            if not cleared_sun:
                sheet1.range(f'{Sun_s}10:{Sun_s}27').value = ""
                sheet1.range(f'{Sun_w}10:{Sun_w}27').value = ""
                sheet1.range(f'{Sun_e}10:{Sun_e}27').value = ""
                cleared_sun = True
            sheet1.range(f'{Sun_s}{i + 10}').value = ts
            sheet1.range(f'{Sun_w}{i + 10}').value = wages[i]
            sheet1.range(f'{Sun_e}{i + 10}').value = te
wages_write()
    
