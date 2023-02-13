import json
import os
import sys
import tkinter
import platform
from tkinter import filedialog

import customtkinter
from main import main
from PIL import Image

customtkinter.set_appearance_mode("System")  # Modes: system (default), light, dark
customtkinter.set_default_color_theme(
    "green"
)  # Themes: blue (default), dark-blue, green

system = platform.system()

app = customtkinter.CTk()  # create CTk window like you do with the Tk window
main_width = 880
main_height = 550
app.geometry(f"{main_width}x{main_height}")
app.title("Excel Auto Formatter")

main_path = os.path.dirname(sys.argv[0])

with open(f"{main_path}/settings.json", "r", encoding="utf-8") as _settings:
    global inventory_path, permanent_path, permanent_db, subject_path, output_path
    settings = json.load(_settings)
    inventory_path = settings["inventory_excel"]
    permanent_path = settings["permanent_excel"]
    permanent_db = settings["permanent_db"]
    subject_path = settings["subject_folder"]
    output_path = settings["output_folder"]

dic = {
    "手配品,先行品のExcelファイル": inventory_path,
    "常置品のExcelファイル": permanent_path,
    "常置品のDBファイル": permanent_db,
    "加工前の手配リストの入っているフォルダー": subject_path,
    "加工後の手配リストを入れるフォルダー": output_path,
}

dic_name = [
    "inventory_excel",
    "permanent_excel",
    "permanent_db",
    "subject_folder",
    "output_folder",
]


class PathEntry(customtkinter.CTkFrame):
    index = 0
    # CustomTkinterのフレームを継承したクラス
    def __init__(self, master, width, height, path_text, index):
        # super()は親クラスのメソッドを呼び出す（使えるようにする）
        super().__init__(master, width, height)
        self.index = index
        self.path_disp = customtkinter.CTkEntry(
            master=self.master,
            width=width - 65,
            height=height - 5,
            font=("Arial", 11),
        )
        self.path_disp.insert(0, path_text)
        # 挿入されてからreadonlyにしないと文字が表示されない
        self.path_disp.configure(state="readonly")
        self.path_disp.place(relx=0.02, rely=0.5, anchor=tkinter.W)
        self.button = customtkinter.CTkButton(
            master=self.master,
            image=icon,
            command=self.open_folder,
            text="",
            width=24,
            height=24,
        )
        self.button.place(relx=0.99, rely=0.5, anchor=tkinter.E)

    def open_folder(self):
        if self.index <= 2:
            if system == "Darwin":
                new_path = filedialog.askopenfilename(
                    initialdir="/",
                    filetypes=(("Excel .xlsx", "Excel .xls"),),
                )
            else:
                new_path = filedialog.askopenfilename(
                    initialdir="/",
                    filetypes=(
                        ("Excel", ".xlsx .xls"),
                        ("ExcelMacro .xlsm"),
                    ),
                )
                print(new_path)
                new_path = new_path.replace("/", "\\")
        else:
            # フォルダーを選択する
            new_path = filedialog.askdirectory()
            if system == "Windows":
                new_path = new_path.replace("/", "\\")
        # まずは表示を変更，そうしないと値が変更できず表示されない
        self.path_disp.configure(state="normal")
        self.path_disp.delete(0, tkinter.END)
        self.path_disp.insert(0, new_path)
        self.path_disp.configure(state="readonly")
        dic[list(dic.keys())[self.index]] = new_path


period = 0.15
width = main_width * 0.96
height = 40

icon = customtkinter.CTkImage(
    light_image=Image.open(f"{main_path}/assets/icons8-folder-48-light.png"),
    dark_image=Image.open(f"{main_path}/assets/icons8-folder-48-dark.png"),
    size=(22, 22),
)

for i in range(len(dic)):
    label = list(dic.keys())[i]
    label_disp = customtkinter.CTkLabel(
        master=app, text=list(dic.keys())[i], font=("Arial", 14, "bold")
    )
    label_disp.place(relx=0.05, rely=0.05 + period * i, anchor=tkinter.W)

    frame = customtkinter.CTkFrame(
        master=app,
        width=width,
        height=height,
    )
    frame.grid(row=0, column=0, padx=20)
    frame.place(
        relx=0.5,
        rely=period * i + 0.12,
        anchor=tkinter.CENTER,
    )

    path = list(dic.values())[i]
    PathEntry(frame, width, height, path, i)


def Process():
    toplevel = customtkinter.CTkToplevel()
    toplevel.geometry("300x200")
    # dicの名前を変更(日本語で扱っていたため)
    new_dic = {}
    for i in range(len(dic)):
        new_dic[dic_name[i]] = list(dic.values())[i]
    with open(f"{main_path}/settings.json", "w", encoding="utf-8") as _settings:
        json.dump(new_dic, _settings, indent=4, ensure_ascii=False)
    isFilled = True
    for i in range(len(dic)):
        if list(dic.values())[i] == "":
            isFilled = False
    ProcessLookupError = customtkinter.CTkLabel(
        master=toplevel, text="", font=("Arial", 16, "bold")
    )
    ProcessLookupError.place(relx=0.5, rely=0.5, anchor=tkinter.CENTER)
    if isFilled == True:
        toplevel.title("Processing...")
        ProcessLookupError.configure(text="処理中です...")
        main()
        ProcessLookupError.configure(text="処理が完了しました")
        ProcessLookupError.place(relx=0.5, rely=0.4, anchor=tkinter.CENTER)
        fin = customtkinter.CTkButton(master=toplevel, text="終了する", command=app.destroy)
        fin.place(relx=0.5, rely=0.8, anchor=tkinter.CENTER)
    else:
        toplevel.title("Error")
        ProcessLookupError.configure(text="すべての項目を入力してください")


button = customtkinter.CTkButton(master=app, text="Process", command=Process)
button.place(relx=0.5, rely=0.9, anchor=tkinter.CENTER)


app.mainloop()
