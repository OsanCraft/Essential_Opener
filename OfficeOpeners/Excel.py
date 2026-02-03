import webbrowser

EXCEL_URL = "https://excel.cloud.microsoft/en-us/?wdOrigin=OFFICECOM-WEB.REDIRECT"


def open_excel():
    return webbrowser.open_new_tab(EXCEL_URL)
