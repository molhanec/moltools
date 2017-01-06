#!python3.4
#encoding: UTF-8

# Assorted various utility functions. Public domain.

def setup_locale():
    "Use locale.strxfrm() for strings"
    import locale
    locale.setlocale(locale.LC_ALL, '')


def time_function(name, module='__main__'):
    from timeit import timeit
    return timeit(name + "()", setup="from %s import %s" % (module, name), number=1)
    

def open_csv(inputfilename):
    import csv  
    with open(str(inputfilename), "r", newline="", encoding="cp1250") as f:
        reader = csv.reader(f, delimiter=';')
        for row in reader:
            yield row
            
def find_newest_folder(info, start_path='.'):
    " Searches for the newest input filename. "

    from collections import namedtuple
    from datetime import date
    from pathlib import Path

    NewestFolder = namedtuple('NewestFolder', 'date path basename fullpath')
    result = None
    info("Looking for an input file")
    p = Path(start_path)
    newest_month = "00"
    newest_year = "0000"
    for f in p.glob("????-??"):
        parts = f.stem.split("-")
        if len(parts) == 2:
            year, month = parts
            if year >= newest_year or (year == newest_year and month > newest_month):
                info("Found newer source: %s", f)
                newest_month = month
                newest_year = year
                result = NewestFolder(
                    date(int(year), int(month), 1),
                    f,
                    f.stem,
                    str(f),
                )
        else:
            info("Skipping >%s<", f)
    return result


def next_month(date):
    return date.replace(month=date.month+1 if date.month<12 else 1, year=date.year if date.month<12 else date.year+1)

    
def date_as_a_month_word_and_year(date):
    return date.strftime("%B %Y")

    
def compose_email(to, carbon_copy, subject, body, attachements):
    "Attachements should be list of Path"
    import win32com.client as win32
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = to
    if carbon_copy:
        if not isinstance(carbon_copy, str):
            carbon_copy = ';'.join(carbon_copy)
        mail.CC = carbon_copy
    mail.Subject = subject
    mail.Body = body 
    for attachment in attachements:
        mail.Attachments.Add(str(attachment.resolve()))
    mail.Display()

    
# https://github.com/welbornprod/easysettings
# https://github.com/ActiveState/appdirs
# pip install easysettings
# pip install appdirs
def load_or_create_app_config(appname, configname="config.ini", config_class=None):
    import appdirs
    import os

    if config_class is None:
        from easysettings import EasySettings
        config_class = EasySettings

    if appname[0].islower():
        appname = appname.capitalize()

    path = appdirs.user_config_dir(appname, appauthor=False) # Don't make author specific subdirectory
    os.makedirs(path, exist_ok=True)
    path = os.path.join(path, configname)
    settings = config_class(path)
    settings.set_and_save = settings.setsave
    return settings
