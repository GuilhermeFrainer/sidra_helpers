import datetime


def get_period(START_DATE : str) -> str:
    
    today = datetime.date.today()
    start_date = datetime.date.fromisoformat(START_DATE)
    return f"{start_date.year}{start_date.month:02d}-{today.year}{today.month:02d}"


# Gets api data and returns it as a list of dicts
def api_to_list(list: list) -> list:
    
    out_list = []

    for i in range(len(list)):
        # Converts dates to datetime objects
        date = list[i]["D2C"][:4] + "-" + list[i]["D2C"][4:] + "-01"
        date = datetime.date.fromisoformat(date)
        out_list.append({"date": date, "value": float(list[i]["V"])})

    return out_list