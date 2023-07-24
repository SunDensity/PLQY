from tkinter.filedialog import askopenfile
import pandas as pd


def file_select():
    file = askopenfile()
    title = file.name.split("/")[-1]
    df = pd.DataFrame()
    if file.name.split(".")[1] == "csv":
        df = pd.read_csv(file.name)
    elif file.name.split(".")[1] == "xlsx":
        df = pd.read_excel(file.name)

    if not df.empty:
        return title.split(".")[0], df
