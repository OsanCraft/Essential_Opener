import webbrowser

WORD_URL = "https://word.cloud.microsoft/en-us/?wdOrigin=OFFICECOM-WEB.REDIRECT"

def open_word():
    return webbrowser.open_new_tab(WORD_URL)