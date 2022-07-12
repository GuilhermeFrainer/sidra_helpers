import datetime
import xlsxwriter


def get_period(START_DATE : str) -> str:
    
    today = datetime.date.today()
    start_date = datetime.date.fromisoformat(START_DATE)
    return f"{start_date.year}{start_date.month:02d}-{today.year}{today.month:02d}"


# Gets api data and returns it as a list of dicts
def api_to_list(list: list[list]) -> list[list]:
    
    out_list = []
    longest_list_size = 0

    # Gets dates from longest list
    for i in range(len(list)):
        if len(list[i]) > longest_list_size:
            longest_list_size = len(list[i])
            longest_list = i

    dates = []

    for i in range(longest_list_size):
        date = list[longest_list][i]["D2C"][:4] + "-" + list[i]["D2C"][4:] + "-01"
        date = datetime.date.fromisoformat(date)
        dates.append(date)

    out_list.append(dates)

    # Creates a list for each series. Skips
    for i in range(len(list)):
        new_list = []

        for j in range(longest_list_size):
            new_list.append(float(list[i][j]['V']))

        out_list.append(new_list)

    return out_list


def make_excel(filename : str, list : list[list[dict]], headers : list) -> tuple[xlsxwriter.Workbook, xlsxwriter.Workbook.worksheet_class]:

    filename = filename + f" {datetime.date.today().isoformat()}.xlsx"
    
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet('Dados')

    # Writes headers
    for i in range(len(headers)):
        worksheet.write(0, i, headers[i])

    # Writes data
    for i in range(len(list)):

        for j in range(len(list[i])):

            worksheet.write(1 + i, j, list[i][j])
    
    return workbook, worksheet

