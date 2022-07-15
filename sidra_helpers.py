import datetime
import xlsxwriter


longest_series_size = 0
date_format = None
number_format = None


def get_period(START_DATE : str) -> str:
    
    today = datetime.date.today()
    start_date = datetime.date.fromisoformat(START_DATE)
    return f"{start_date.year}{start_date.month:02d}-{today.year}{today.month:02d}"


# Gets api data and returns it as a list of dicts
def api_to_list(series_list: list[list]) -> list[list]:

    global longest_series_size
    out_list = []
    longest_series_size = 0
    
    # Picks longest series to use as basis to save dates
    for i in range(len(series_list)):

        if len(series_list[i]) > longest_series_size:
            longest_series_size = len(series_list[i])
            longest_series = i

    # Saves dates in the first list
    dates = []
    for item in series_list[longest_series]:
        date = item['D2C'][:4] + '-' + item['D2C'][4:] + '-01'
        dates.append(datetime.date.fromisoformat(date))
    
    out_list.append(dates)

    # Gets data for each series and saves it into a list
    for series in series_list:
        new_series = []
        for item in series:
            new_series.append(float(item['V']))
        out_list.append(new_series)

    return out_list


def make_excel(filename : str, series_list : list[list], headers : list, index_chart=False) -> tuple[xlsxwriter.Workbook, xlsxwriter.Workbook.worksheet_class]:

    global date_format, number_format

    skipped_lines = 0
    if index_chart == True:
        skipped_lines = 4

    filename = filename + f" {datetime.date.today().isoformat()}.xlsx"
    
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet('Dados')
    date_format = workbook.add_format({'num_format': 'mmm/yy'})
    number_format = workbook.add_format({'num_format': '##0.0'})

    # Writes headers
    for i in range(len(headers)):
        worksheet.write(0 + skipped_lines, i, headers[i])

    # Writes data
    for j in range(len(series_list)):

        for i in range(len(series_list[j])):

            # Writes dates
            if j == 0:
                worksheet.write_datetime(skipped_lines + 1 + i, j, series_list[j][i], date_format)
            # Writes numeric data
            else:
                worksheet.write(skipped_lines + 1 + i, j, series_list[j][i])
    
    return workbook, worksheet


# Used for index charts
def write_index_formulas(workbook : xlsxwriter.Workbook, worksheet : xlsxwriter.Workbook.worksheet_class, headers):

    global longest_series_size, date_format, number_format
    
    # Determines first column to be used for the corrected values
    first_column = len(headers) + 1
    merge_format = workbook.add_format({'align': 'center'})

    if len(headers) - 1 > 1:
        worksheet.merge_range(0, first_column, 0, first_column, first_column + len(headers) - 1, 'Valores de correção', merge_format)
        worksheet.merge_range(3, 1, 3, len(headers), 'Inalterados', merge_format)
        worksheet.merge_range(3, first_column, 3, first_column + len(headers) - 1, 'Corrigidos', merge_format)

    else:
        worksheet.write(0, first_column, 'Valores de correção')
        worksheet.write(3, 1, 'Inalterados')
        worksheet.write(3, first_column, 'Corrigidos')

    # Writes headers and correction values
    for i in range(1, len(headers)):
        worksheet.write(4, first_column + i - 1, headers[i])

        # Finds correction values
        first_cell = xlsxwriter.utility.xl_rowcol_to_cell(5, i)
        last_cell = xlsxwriter.utility.xl_rowcol_to_cell(5 + longest_series_size, i)
        correction_value_cell = xlsxwriter.utility.xl_rowcol_to_cell(1, first_column + i - 1)
        worksheet.write_formula(f'{correction_value_cell}', f'=LARGE({first_cell}:{last_cell},1)', number_format)
        
        # Finds dates for the correction values
        first_cell_date = xlsxwriter.utility.xl_rowcol_to_cell(5, 0)
        last_cell_date = xlsxwriter.utility.xl_rowcol_to_cell(5 + longest_series_size, 0)
        worksheet.write_formula(2, first_column + i - 1, f'=INDEX({first_cell_date}:{last_cell_date},MATCH({correction_value_cell},{first_cell}:{last_cell},0))', date_format)

    # Adds column with "100" in every row
    for i in range(longest_series_size + 1):
        worksheet.write(4 + i, first_column + len(headers) - 1, 100)

    # Writes corrected values
    for i in range(1, len(headers)):

        for j in range(longest_series_size):

            original_value = xlsxwriter.utility.xl_rowcol_to_cell(5 + j, i)
            correction_value = xlsxwriter.utility.xl_rowcol_to_cell(1, first_column + i - 1)
            worksheet.write_formula(5 + j, first_column + i - 1, f'{original_value}*100/{correction_value}')

