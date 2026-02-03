import webbrowser

POWERPOINT_URL = "https://powerpoint.cloud.microsoft/?wdOrigin=OFFICECOM-WEB.REDIRECT"

def open_powerpoint():
    return webbrowser.open_new_tab(POWERPOINT_URL)
