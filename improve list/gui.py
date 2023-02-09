import os
import sys
import json
import tkinter
import customtkinter
from tkinter import filedialog

customtkinter.set_appearance_mode("System")  # Modes: system (default), light, dark
customtkinter.set_default_color_theme(
    "green"
)  # Themes: blue (default), dark-blue, green

app = customtkinter.CTk()  # create CTk window like you do with the Tk window
app.geometry("800x500")
app.title("Excel Auto Formatter")

main_path = os.path.dirname(sys.argv[0])


def button_function():
    # file explorer window
    root = tkinter.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    root.update()
    file_path = filedialog.askopenfilename()
    root.destroy()
    print(file_path)
    return file_path


with open(f"{main_path}/settings.json", "r") as _settings:
    settings = json.load(_settings)
    inventory_path = settings["inventory_excel"]
    subject_path = settings["subject_folder"]
    output_path = settings["output_folder"]

dir = {
    "在庫表のExcelファイル": inventory_path,
    "加工前の手配リストの入っているフォルダー": subject_path,
    "加工後の手配リストを入れるフォルダー": output_path,
}

for i in range(len(dir)):
    label = customtkinter.CTkLabel(master=app, text=list(dir.keys())[i])
    label.place(relx=0.1, rely=0.1 + 0.2 * i, anchor=tkinter.W)

    frame = customtkinter.CTkFrame(master=app, width=700, height=45)
    frame.grid(row=0, column=0, padx=20)
    frame.place(
        relx=0.5,
        rely=0.2 * (i + 1),
        anchor=tkinter.CENTER,
    )

    path = customtkinter.CTkLabel(master=frame, text=list(dir.values())[i])
    path.place(relx=0.05, rely=0.5, anchor=tkinter.W)

inventory_path = ""

if inventory_path != "" and subject_path != "" and output_path != "":
    button = customtkinter.CTkButton(
        master=app, text="Process", command=button_function
    )
    button.place(relx=0.5, rely=0.8, anchor=tkinter.CENTER)
else:
    avoid = customtkinter.CTkLabel(master=frame, text="すべての設定を完了してください")
    avoid.place(relx=5, rely=0.8, anchor=tkinter.CENTER)


def button_func(inventory_path):
    inventory_path = "a"
    print(inventory_path)
    app.update()


button = customtkinter.CTkButton(
    master=app, text="Pa", command=button_func(inventory_path)
)
button.place(relx=0.5, rely=0.9, anchor=tkinter.CENTER)

app.mainloop()
