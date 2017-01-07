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


from configparser import ConfigParser
class MolConfig:

    IMPLICIT_SECTION = "Settings"

    def __init__(self, path):
        self.path = path
        self.config_obj = ConfigParser()
        if path.exists():
            with path.open(encoding="UTF-8") as f:
                self.config_obj.read_file(f)
        if MolConfig.IMPLICIT_SECTION not in self.config_obj:
            self.config_obj.add_section(MolConfig.IMPLICIT_SECTION)

    def save(self):
        with open(str(self.path), "wt", encoding="UTF-8") as f:
            self.config_obj.write(f)

    def get(self, name, default=""):
        return self.config_obj.get(MolConfig.IMPLICIT_SECTION, name, fallback=default)

    def set(self, name, value):
        return self.config_obj.set(MolConfig.IMPLICIT_SECTION, name, str(value))

    def get_bool(self, name, default=False):
        return self.config_obj.getboolean(MolConfig.IMPLICIT_SECTION, name, fallback=default)

    def get_int(self, name, default=0):
        return self.config_obj.getint(MolConfig.IMPLICIT_SECTION, name, fallback=default)

    def get_list(self, name):
        list_str = self.get(name)
        return list(filter(lambda string: string != "", list_str.split("|")))

    def set_list(self, name, list):
        self.set(name, "|".join(list))


# https://github.com/ActiveState/appdirs
# pip install appdirs
def load_or_create_app_config(appname, configname="config.ini", config_class=None):
    import appdirs
    from pathlib import Path

    if config_class is None:
        config_class = MolConfig

    if appname[0].islower():
        appname = appname.capitalize()

    path = appdirs.user_config_dir(appname, appauthor=False) # Don't make author specific subdirectory
    path = Path(path)
    path.mkdir(parents=True, exist_ok=True)
    path /= configname
    config = config_class(path)
    return config
