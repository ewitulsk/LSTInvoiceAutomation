import openpyxl
import datetime as dt
import readchar
import os

class TimeRecord:
    date: dt.date
    employee: str
    hours_worked: float
    rate: float
    work_type: str
    
    def __init__(self,date,employee,hours_worked,rate,work_type):
        self.date = date
        self.employee = employee
        self.hours_worked = hours_worked
        self.rate = rate
        self.work_type = work_type

    def __str__(self) -> str:
        return f"\n({self.date},{self.employee},{self.hours_worked},{self.rate},{self.work_type})"
    def __repr__(self) -> str:
        return f"\n({self.date},{self.employee},{self.hours_worked},{self.rate},{self.work_type})"

class MilesRecord:
    date: dt.date
    miles: float
    miles_rate: float

    def __repr__(self) -> str:
        return f"\n({self.date},{self.miles},{self.miles_rate})"

    def __str__(self) -> str:
        return f"\n({self.date},{self.miles},{self.miles_rate})"

    def __init__(self,date,miles,miles_rate):
        self.date = date
        self.miles = miles
        self.miles_rate = miles_rate

class GPSRecord:
    date: dt.date
    gps: float
    gps_rate: float

    def __init__(self,date,gps,gps_rate) -> None:
        self.date = date
        self.gps = gps
        self.gps_rate = gps_rate
    
    def __repr__(self) -> str:
        return f"\n({self.date},{self.gps},{self.gps_rate})"
    
    def __str__(self) -> str:
        return f"\n({self.date},{self.gps},{self.gps_rate})"
    
class SokkiaRecord:
    amount: float
    rate: float

    def __init__(self,amount,rate) -> None:
        self.amount = amount
        self.rate = rate
    
    def __repr__(self) -> str:
        return f"\n({self.amount},{self.rate})"
    
    def __str__(self) -> str:
        return f"\n({self.amount},{self.rate})"

class MiscSuppliesRecord:
    rebar: float
    rebar_rate: float
    lsrm: float
    lsrm_rate: float
    lath: float
    lath_rate: float
    spikes: float
    spikes_rate: float
    tpost: float
    tpost_rate: float
    
    def __init__(self,rebar,rebar_rate,lsrm,lsrm_rate,lath,lath_rate,spikes,spikes_rate,tpost,tpost_rate) -> None:
        self.rebar = rebar
        self.rebar_rate = rebar_rate
        self.lsrm = lsrm
        self.lsrm_rate = lsrm_rate
        self.lath = lath
        self.lath_rate = lath_rate
        self.spikes = spikes
        self.spikes_rate = spikes_rate
        self.tpost = tpost
        self.tpost_rate = tpost_rate

    def __str__(self) -> str:
        return f"Rebar: ({self.rebar}: {self.rebar_rate})\nLSRM: ({self.lsrm}: {self.lsrm_rate})\nLath: ({self.lath}: {self.lath_rate})\nSpikes: ({self.spikes}: {self.spikes_rate})\nTPost: ({self.tpost}: {self.tpost_rate})"

    def __repr__(self) -> str:
        return f"Rebar: ({self.rebar}: {self.rebar_rate})\nLSRM: ({self.lsrm}: {self.lsrm_rate})\nLath: ({self.lath}: {self.lath_rate})\nSpikes: ({self.spikes}: {self.spikes_rate})\nTPost: ({self.tpost}: {self.tpost_rate})"

#Not really a daily record as each line has its own date.
#This naming convention matches with the sheet
class DailyRecord:
    job_name: str
    time_records: [TimeRecord]
    miles_records: [MilesRecord]
    gps_records: [GPSRecord]
    sokkia_records: [SokkiaRecord]
    misc_record: MiscSuppliesRecord
    opex: float

    def __init__(self,job_name,time_records,miles_records,gps_records,sokkia_records,misc_record,opex=2.8) -> None:
        self.job_name = job_name
        self.time_records = time_records
        self.miles_records = miles_records
        self.gps_records = gps_records
        self.sokkia_records = sokkia_records
        self.misc_record = misc_record
        self.opex = opex

    def __repr__(self) -> str:
        return f"{self.job_name}"


class Heading:
    name: str
    row: int
    col: str
    def loc(self): 
        return f"{self.col, self.row}"
    
    def __str__(self) -> str:
        return f"{self.name} ({self.col}{self.row})"
    
    def __repr__(self) -> str:
        return f"{self.name} ({self.col}{self.row})"
    
    def __init__(self,name,col,row):
        self.name = name
        self.col = col
        self.row = row


def find_section_heading(ws):
    locations = dict()
    valid_headings = ["Daily County Record", "Daily Supplies Record"]
    for row in ws.iter_rows(min_row=1, max_col=ws.max_column, max_row=ws.max_row):
        for cell in row:
            if cell.value in valid_headings:
                locations[cell.value] = Heading(cell.value, cell.column_letter, cell.row)
    return locations


def cell_arr_is_empty(cell_arr):
    for cell in cell_arr:
        if cell.value != None:
            return False
    return True


def get_inc_dict(val, dict):
    if val not in dict:
        dict[val] = 1
    else:
        dict[val] += 1
    return dict[val]


def get_last_section_row(ws, section_start_row):
    #occurrence of stop words in section
    stop_words_occ = dict()
    #If a stop word is encountered x number of times, that row is where the section ends
    stop_words = {"Sub Total":1 , "Op/Exp":1, "Total":2}
    for row in ws.iter_rows(min_row=section_start_row, max_col=ws.max_column, min_col=1):
        for cell in row:
            if cell.value in stop_words.keys():
                stop_word = cell.value
                occ = get_inc_dict(cell.value, stop_words_occ)
                if occ >= stop_words[stop_word]:
                    return cell.row-1
    return None


#potential_cols is array of all columns with empty names
def find_time_record_date_col(ws, potential_cols, section_start_row, section_end_row):
    for col_name in potential_cols:
        col_tups = ws[f"{col_name}{section_start_row}":f"{col_name}{section_end_row}"]
        col = [x[0] for x in col_tups]
        if not cell_arr_is_empty(col):
            return col_name
    return None


def get_section_title_cols(ws, section_start_row, section_end_row):
    title_col_dict = dict()
    empty_names = []
    row = ws[section_start_row]
    for cell in row:
        if cell.value == None:
            empty_names.append(cell.column_letter)
        else:
            title_col_dict[cell.value] = cell.column_letter
    if "Date" not in title_col_dict.keys() and "date" not in title_col_dict.keys():
        date_col = find_time_record_date_col(ws, empty_names, section_start_row, section_end_row)
        title_col_dict["Date"] = date_col
        title_col_dict["date"] = date_col
    return title_col_dict

#Key words are words that are likely to indicate that tables start
def find_table_start(ws, key_words, tolerance=.66):
    for row in ws.iter_rows(min_row=1, max_col=ws.max_column, min_col=1):
        cell_vals = [cell.value for cell in row if cell.value != None]
        matched_keywords = 0
        for word in key_words:
            if word in cell_vals:
                matched_keywords += 1
        if matched_keywords/len(key_words) > tolerance:
            return row[0].row
    return None


def get_job_name(ws, key_words):
    start = find_table_start(ws, key_words)
    row = ws[start]
    cell_vals = [cell.value for cell in row if cell.value != None]
    cell_vals_cpy = cell_vals.copy()
    for word in key_words:
        if word in cell_vals_cpy:
            cell_vals_cpy.remove(word)
    if len(cell_vals_cpy) > 1:
        print("ERROR: MORE THAN 1 POSSIBLE JOB NAME")
        exit()
    return cell_vals_cpy[0]
    


#Row is an array of cells
#Return TimeRecord obj
def read_county_line(row, title_col_dict):
    if "Date" in title_col_dict.keys():
        date_col = title_col_dict["Date"]
        date_num = ord(date_col.lower())-97
        date = row[date_num].value
    else:
        date = None
    
    if date is None:
        return None
    
    if "Name" in title_col_dict.keys():
        employee_col = title_col_dict["Name"]
        employee_num = ord(employee_col.lower())-97
        employee = row[employee_num].value
    else:
        employee = None

    if "Hours worked" in title_col_dict.keys():
        hours_col = title_col_dict["Hours worked"]
        hours_num = ord(hours_col.lower())-97
        hours_worked = row[hours_num].value
    else:
        hours_worked = None

    
    if "Rate" in title_col_dict.keys():
        rate_col = title_col_dict["Rate"]
        rate_num = ord(rate_col.lower())-97
        rate = row[rate_num].value
    else:
        rate = None

    if "Type" in title_col_dict.keys():
        type_col = title_col_dict["Type"]
        type_num = ord(type_col.lower())-97
        work_type = row[type_num].value
    else:
        work_type = None
    
    time_record = TimeRecord(date, employee, hours_worked, rate, work_type)
    return time_record


def read_county_record(ws):
    job_name = get_job_name(ws, ["Name", "Hours worked", "Rate", "Total", "Type", "Date"])

    time_records = []

    section_start = find_table_start(ws, ["Name", "Hours worked"])
    if section_start == None:
        return None
    section_last_row =  get_last_section_row(ws, section_start)
    title_col_dict = get_section_title_cols(ws, section_start, section_last_row)
    for row in ws.iter_rows(min_row=section_start+1, max_col=ws.max_column, max_row=section_last_row):
        time_record = read_county_line(row, title_col_dict)
        if time_record != None:
            time_records.append(time_record)

    return time_records, job_name
    

def read_miles_line(row, title_col_dict):  
    if "Date" in title_col_dict.keys():
        date_col = title_col_dict["Date"]
        date_num = ord(date_col.lower())-97
        date = row[date_num].value
    else:
        date = None
    
    if date is None:
        return None
    
    if "Miles 2-1704" in title_col_dict.keys():
        miles_col = title_col_dict["Miles 2-1704"]
        miles_num = ord(miles_col.lower())-97
        miles = row[miles_num].value
    else:
        miles = None

    if "Rate" in title_col_dict.keys():
        rate_col = title_col_dict["Rate"]
        rate_num = ord(rate_col.lower())-97
        rate = row[rate_num].value
    else:
        rate = None
    
    miles_record = MilesRecord(date, miles, rate)
    return miles_record


def read_miles_record(ws):
    miles_records = []
    section_start = find_table_start(ws, ["Miles 2-1704"])
    if section_start == None:
        return None
    section_last_row =  get_last_section_row(ws, section_start)
    title_col_dict = get_section_title_cols(ws, section_start, section_last_row)
    for row in ws.iter_rows(min_row=section_start+1, max_col=ws.max_column, max_row=section_last_row):
        miles = read_miles_line(row, title_col_dict)
        if miles != None:
            miles_records.append(miles)
    return miles_records


def read_gps_line(row, title_col_dict):
    
    if "Date" in title_col_dict.keys():
        date_col = title_col_dict["Date"]
        date_num = ord(date_col.lower())-97
        date = row[date_num].value
    else:
        date = None
    
    if date is None:
        return None
    
    
    if "GPS 2-2500" in title_col_dict.keys():
        gps_col = title_col_dict["GPS 2-2500"]
        gps_num = ord(gps_col.lower())-97
        gps = row[gps_num].value
    else:
        gps = None

    
    if "Rate" in title_col_dict.keys():
        rate_col = title_col_dict["Rate"]
        rate_num = ord(rate_col.lower())-97
        rate = row[rate_num].value
    else:
        rate = None
    
    gps_record = GPSRecord(date, gps, rate)
    return gps_record


def read_gps_record(ws):
    gps_records = []
    section_start = find_table_start(ws, ["GPS 2-2500"])
    if section_start == None:
        return None
    section_last_row =  get_last_section_row(ws, section_start)
    title_col_dict = get_section_title_cols(ws, section_start, section_last_row)
    for row in ws.iter_rows(min_row=section_start+1, max_col=ws.max_column, max_row=section_last_row):
        gps = read_gps_line(row, title_col_dict)
        if gps != None:
            gps_records.append(gps)
    return gps_records


def read_sokkia_line(row, title_col_dict):
    if "SOKKIA  2-2500" in title_col_dict.keys():
        amount_col = title_col_dict["SOKKIA  2-2500"]
        amount_num = ord(amount_col.lower())-97
        amount = row[amount_num].value
    else:
        amount = None
    
    if amount is None:
        return None
    
    
    if "Rate" in title_col_dict.keys():
        rate_col = title_col_dict["Rate"]
        rate_num = ord(rate_col.lower())-97
        rate = row[rate_num].value
    else:
        rate = None
    
    return SokkiaRecord(amount, rate)

def read_sokkia_record(ws):
    sokkia_records = []
    section_start = find_table_start(ws, ["SOKKIA  2-2500"])
    if section_start == None:
        return None
    section_end = get_last_section_row(ws, section_start)
    title_col_dict = get_section_title_cols(ws, section_start, section_end)
    for row in ws.iter_rows(min_row=section_start+1, max_col=ws.max_column, max_row=section_end):
        sokkia_record = read_sokkia_line(row, title_col_dict)
        if sokkia_record != None:
            sokkia_records.append(sokkia_record)
    return sokkia_records


def read_misc_line(row, title_col_dict, item):
    if item in title_col_dict.keys():
        amount_col = title_col_dict[item]
        amount_num = ord(amount_col.lower())-97
        amount = row[amount_num].value
    else:
        amount = None
    
    if amount is None:
        return None
    
    if "Rate" in title_col_dict.keys():
        rate_col = title_col_dict["Rate"]
        rate_num = ord(rate_col.lower())-97
        rate = row[rate_num].value
    else:
        rate = None
    
    return amount, rate

def read_misc_record(ws, item):
    section_start = find_table_start(ws, [item])
    if section_start == None:
        return None
    title_col_dict = get_section_title_cols(ws, section_start, section_start+1)
    row = ws[section_start+1]
    return read_misc_line(row, title_col_dict, item)





def read_sheet(ws):
    time_records_maybe = read_county_record(ws)
    if time_records_maybe != None:
        time_records, job_name = time_records_maybe
    else:
        time_records, job_name = None, None
    
    miles_records = read_miles_record(ws)
    gps_records = read_gps_record(ws)
    sokkia_records = read_sokkia_record(ws)


    rebar_record = read_misc_record(ws, "Rebar 3-0306")
    lsrm_record = read_misc_record(ws, "LS/RM not AL")
    spikes_record = read_misc_record(ws, "Spikes 3-0306")
    lath_record = read_misc_record(ws, "Lath 3-0306")
    tpost_record = read_misc_record(ws, "T-Post 3-0306")
    
    if rebar_record is not None:
        rebar_amount, rebar_rate = rebar_record
    else:
        rebar_amount, rebar_rate = None, None
    if lsrm_record is not None:
        lsrm_amount, lsrm_rate = lsrm_record
    else:
        lsrm_amount, lsrm_rate = None, None
    if spikes_record is not None:
        spikes_amount, spikes_rate = spikes_record
    else:
        spikes_amount, spikes_rate = None, None
    if lath_record is not None:
        lath_amount, lath_rate = lath_record
    else:
        lath_amount, lath_rate = None, None
    if tpost_record is not None:
        tpost_amount, tpost_rate = tpost_record
    else:
        tpost_amount, tpost_rate = None, None

    misc = MiscSuppliesRecord(rebar_amount, rebar_rate, lsrm_amount, lsrm_rate, lath_amount, lath_rate, spikes_amount, spikes_rate, tpost_amount, tpost_rate)

    return DailyRecord(job_name, time_records, miles_records, gps_records, sokkia_records, misc)


def print_sheet(sheet):
    print(sheet)

    print("Time Records\n")
    print(sheet.time_records)
    print(f"Op/Ex: {sheet.opex}")
    print("Miles Record\n")
    print(sheet.miles_records)
    print("GPS Record\n")
    print(sheet.gps_records)
    print("SOKKIA Record\n")
    print(sheet.sokkia_records)
    print("Misc Records\n")
    print(sheet.misc_record)


def sum_sheets(sheets):
    total = 0

def clear():
    os.system('cls||clear')


def display_sheets(sheets):
    cur_sheet = 0
    is_printed = False

    while(True):
        clear()
        print_sheet(sheets[cur_sheet])
        print(f"{cur_sheet+1}/{len(sheets)}")

        key = readchar.readkey()

        if key == readchar.key.RIGHT:
            cur_sheet += 1
        if key == readchar.key.LEFT:
            cur_sheet -= 1
        
        if cur_sheet > len(sheets)-1:
            cur_sheet = 0
            is_printed = False
        
        if cur_sheet < 0:
            cur_sheet = len(sheets)-1
            is_printed = False
            
        


def edit_op_ex():
    clear()
    print("Editing OP/EX")


def setup_sheets():
    #Dictionary of sheet data frames
    file = "sheets/County Time and Supplies Record.xlsx"

    wb = openpyxl.load_workbook(file)

    sheets = []

    for ws in wb.worksheets:
        sheet = read_sheet(ws)
        sheets.append(sheet)
    
    return file, sheets

def intro_screen(file, sheets):
    clear()
    print("LST Invoice Automation")
    print(f"Reading File: {file}")

    print(f"View/Edit Op/Exp Rate: o")
    print(f"View Spreadsheet Data: v")

    key = readchar.readkey()

    match key:
        case "e":
            edit_op_ex()
        case "v":
            display_sheets(sheets)

def main():
    file, sheets = setup_sheets()
    intro_screen(file, sheets)

    


    


if __name__ == "__main__":
    main()
    

    








