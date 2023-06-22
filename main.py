import pandas as pd
import gspread
import calendar
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, date

mapping = {
    'JET WASH': 'Car|Wash',
    'BP-KRAKOWIAK': 'Car|Petrol',
    'ZABKA': 'Food|Zabka',
    'BIEDRONKA': 'Food|Biedronka',
    'AUCHAN': 'Food|Auchan',
    'TRANSGOURMET': 'Food|Selgros',
    'ANASTASIIA ZAKHAROVA': 'Nastya|Cash',
    'Anastasiia Zakharova': 'Nastya|Cash',
    'PIEKARNIA BUCZEK': 'Cafes|Coffee',
    'CUKIERNIA': 'Cafes|Coffee',
    'NAKIELNY': 'Cafes|Coffee',
    'KABOOM': 'Food|Bazar',
    'JAMON4YOU': 'Food|Bazar',
    'KFC': 'Food delivery|KFC',
    'Bistro w Archeturze ': 'Cafes|Lunch',
    'BOLT.EU': 'Service|Taxi',
    'KYIVSTAR': 'Service|Mobile phone',
    'Apteka': 'Health|Pharmacy',
    'Studio Klaru': 'Health|Manic',
    'Internetia ': 'Service|Internet',
    'SCBSA Centrala': 'Loan|iPhone',
    'Alior centrala': 'Business|Loan',
    'Wspolnota Mieszkaniowa': 'Service|Komunalka',
}

def map_to_category(input_str):
    matching_keys = [key for key in mapping.keys() if key in input_str]

    # Map the input string to the corresponding pre-defined string
    if matching_keys:
        # Use the first matching key found
        mapped_str = mapping[matching_keys[0]]
    else:
        mapped_str = 'default|value'

    # Print the mapped value
    print("Mapped value:", mapped_str)
    return mapped_str

def convert_sum(bank, sum):
    match bank:
        case 'Mil':
            return sum * -1
        case 'Pekao':
            return float(sum.replace(',','.').replace(' ',''))*-1
        case 'ING':
            return float(sum.replace(',','.').replace(' ',''))*-1
        case _:
            return 0

def process_bank(bank):
    match bank:
        case 'Mil':
            csv = 'Historia.csv'
            description_header = 'Opis'
            sum_header = 'Obciążenia'
            date_header = 'Data transakcji'
            date_format = '%Y-%m-%d'
            unmapped_sheet = 'Mil_unmapped'
            delimiter = ','
            skiprows = 0
            skipfooter = 0
            encoding = 'utf-8'
        case 'Pekao':
            csv = 'Lista_pekao.csv'
            description_header = 'Nadawca / Odbiorca'
            sum_header = 'Kwota operacji'
            date_header = 'Data księgowania'
            date_format = '%d.%m.%Y'
            unmapped_sheet = 'Pekao_unmapped'
            delimiter = ';'
            skiprows = 0
            skipfooter = 0
            encoding = 'utf-8'
        case 'ING':
            csv = 'Lista_transakcji_ING.csv'
            description_header = 'Dane kontrahenta'
            sum_header = 'Kwota transakcji (waluta rachunku)'
            date_header = 'Data transakcji'
            date_format = '%Y-%m-%d'
            unmapped_sheet = 'ING_unmapped'
            delimiter = ';'
            skiprows = 21
            skipfooter = 1
            encoding = 'cp1252'
        case _:
            return

    global date
    worksheet_unmapped = spreadsheet.worksheet(unmapped_sheet)
    # Read and process the CSV file using pandas
    dataframe = pd.read_csv(csv, keep_default_na=False, skiprows=skiprows, skipfooter=skipfooter, on_bad_lines='skip', delimiter=delimiter, encoding=encoding)
    dataframe.fillna('')
    # Iterate through every row in CSV and put data into GSP
    dataframe = dataframe.reset_index();
    for i, row in dataframe.iterrows():
        opis = row[description_header]
        print(opis, row[sum_header])
        date = row[date_header]
        strptime = datetime.strptime(date, date_format)
        res = calendar.monthrange(strptime.year, strptime.month)
        day = res[1]
        lastDayOfMonth = strptime.replace(day=day)

        category = map_to_category(opis).split('|')
        if category[0]=='default':
            worksheet_unmapped.append_row([
                strptime.date().__str__(),
                lastDayOfMonth.__str__(),
                category[0],
                category[1],
                convert_sum(bank, row[sum_header]),
                # row[sum_header] * -1,
                # float(row[sum_header].replace(',','.').replace(' ',''))*-1,
                opis
            ],
                'RAW', 'INSERT_ROWS', 'A1')
        else:
            worksheet_expenses.append_row([
                strptime.date().__str__(),
                lastDayOfMonth.__str__(),
                category[0],
                category[1],
                convert_sum(bank, row[sum_header]),
                # row[sum_header] * -1,
                # float(row[sum_header].replace(',','.').replace(' ',''))*-1,
                opis
            ],
                'RAW', 'INSERT_ROWS', 'A1')

# Set the scope and credentials
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('budgetautomation.json', scope)

# Authenticate and create a client
client = gspread.authorize(credentials)

# Open the Google Spreadsheet
spreadsheet = client.open_by_key('1M_ZXujZNZNsTM973WjxVM5P7-tSunDJu-6akWsiCOjw')
worksheet_expenses = spreadsheet.worksheet('Expenses')

# Delete the last backup, create latest copy
worksheets = spreadsheet.worksheets()
spreadsheet.del_worksheet_by_id(worksheets[worksheets.__len__()-1].id)
worksheet_expenses.duplicate(new_sheet_name='Backup_Expenses_'+date.today().strftime('%d-%m'),insert_sheet_index=99);

# process_bank('Mil')
process_bank('Pekao')
# process_bank('ING')




# Clear existing data in the worksheet (optional)
# worksheet.clear()

# output = dataframe["Opis"].values
# print (output)
# print ([dataframe.columns.values.tolist()] + dataframe.values.tolist())

# Insert the processed data into the Google Spreadsheet
# worksheet.update([dataframe.columns.values.tolist()] + dataframe.values.tolist())

print("Data imported and processed successfully!")