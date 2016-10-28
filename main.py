from openpyxl import load_workbook
import tkinter as tk
from tkinter import filedialog


def load_from_file(file):
    wb = load_workbook(filename=file, read_only=True)
    ws = wb.active
    routes = {}
    rows = list(ws.rows)
    for row in rows[1:]:
        route = row[0].value
        if route not in routes:
            routes[route] = {}
            print(route)
        routes[route][row[1].value] = row[2].value
    print(routes)
    return routes


def write_preamble(f):
    f.write("use aewdcs\n\n")

    f.write("DELETE FROM SVRS_PollingPlaceUser Where PPLID NOT IN(593215)\n")
    f.write("DELETE FROM SurveyResponseComments\n")
    f.write("DELETE FROM PollingPlacePhoto\n")
    f.write("DELETE FROM SupplyGrant\n")
    f.write("DELETE FROM SurveyResponseHistory\n")
    f.write("DELETE FROM SurveyResponse\n")
    f.write("DELETE FROM ActivityLog\n")
    f.write("DELETE FROM FactSurveyResponse\n\n")


def write_to_file(file, name, route):
    with open(file, "w") as f:
        write_preamble(f)
        f.write("/************************************ Route ")
        f.write(name)
        f.write(" *********************************************************/\n")
        f.write("use AEWDCS\n\n")
        f.write("INSERT INTO SVRS_PollingPlaceUser\n")
        f.write("(PPLID,UserId)\n")
        f.write("VALUES \n\n")

        for location, id in route.items():
            f.write("-- " + location + "\n")
            f.write("(" + str(id) + ",3992)\n\n")


def get_file():
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(defaultextension="xlsx")


def main():
    routes = load_from_file(get_file())
    for name, route in routes.items():
        write_to_file(str(name) + ".sql", str(name), route)


if __name__ == '__main__':
    main()
