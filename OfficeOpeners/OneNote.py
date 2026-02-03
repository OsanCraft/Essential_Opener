import webbrowser

ONENOTE_URL = "https://onenote.cloud.microsoft/en-us/?wdOrigin=OFFICECOM-WEB.REDIRECT"


def open_onenote():
    return webbrowser.open_new_tab(ONENOTE_URL)
