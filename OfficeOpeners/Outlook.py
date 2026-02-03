import webbrowser

OUTLOOK_URL = "https://onenote.cloud.microsoft/?wdOrigin=OFFICECOM-WEB.REDIRECT"


def open_outlook():
    return webbrowser.open_new_tab(OUTLOOK_URL)
