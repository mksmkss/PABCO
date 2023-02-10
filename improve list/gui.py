import os
import sys
import json
import tkinter
import customtkinter
from tkinter import filedialog
from PIL import Image

# from main import main

customtkinter.set_appearance_mode("System")  # Modes: system (default), light, dark
customtkinter.set_default_color_theme(
    "green"
)  # Themes: blue (default), dark-blue, green

app = customtkinter.CTk()  # create CTk window like you do with the Tk window
app.geometry("800x500")
app.title("Excel Auto Formatter")

main_path = os.path.dirname(sys.argv[0])


def button_function(name):

    # file explorer window
    root = tkinter.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    root.update()
    file_path = filedialog.askopenfilename()
    root.destroy()
    print(file_path)
    name = file_path
    # return file_path


with open(f"{main_path}/settings.json", "r") as _settings:
    settings = json.load(_settings)
    inventory_path = settings["inventory_excel"]
    permanent_path = settings["permanent_excel"]
    subject_path = settings["subject_folder"]
    output_path = settings["output_folder"]

dir = {
    "手配品,先行品のExcelファイル": inventory_path,
    "常置品のExcelファイル": permanent_path,
    "加工前の手配リストの入っているフォルダー": subject_path,
    "加工後の手配リストを入れるフォルダー": output_path,
}

period = 0.18
icon = customtkinter.CTkImage(
    light_image=Image.open(f"{main_path}/assets/icons8-folder-48-light.png"),
    dark_image=Image.open(f"{main_path}/assets/icons8-folder-48-dark.png"),
    size=(22, 22),
)
for i in range(len(dir)):
    label = customtkinter.CTkLabel(
        master=app, text=list(dir.keys())[i], font=("Arial", 14, "bold")
    )
    label.place(relx=0.05, rely=0.05 + period * i, anchor=tkinter.W)

    frame = customtkinter.CTkFrame(
        master=app,
        width=750,
        height=40,
    )
    frame.grid(row=0, column=0, padx=20)
    frame.place(
        relx=0.5,
        rely=period * i + 0.12,
        anchor=tkinter.CENTER,
    )
    path = customtkinter.CTkLabel(master=app, font=("Arial", 13))
    path = customtkinter.CTkLabel(
        master=frame, text=list(dir.values())[i], font=("Arial", 11)
    )
    path.place(relx=0.02, rely=0.5, anchor=tkinter.W)
    button = customtkinter.CTkButton(
        master=frame,
        image=icon,
        command=button_function,
        text="",
        width=24,
        height=24,
    )
    button.place(relx=0.99, rely=0.5, anchor=tkinter.E)


def button_func():
    toplevel = customtkinter.CTkToplevel()
    toplevel.geometry("300x200")
    print("button_func")
    if inventory_path != "" and subject_path != "" and output_path != "":
        toplevel.title("Processing...")
        with open("settings.json", "w") as _settings:
            json.dump(
                {
                    "inventory_excel": inventory_path,
                    "permanent_excel": permanent_path,
                    "subject_folder": subject_path,
                    "output_folder": output_path,
                },
                _settings,
            )
        ProcessLookupError = customtkinter.CTkLabel(
            master=toplevel, text="Processing..."
        )
        ProcessLookupError.place(relx=0.5, rely=0.5, anchor=tkinter.CENTER)
        # main()
        fin = customtkinter.CTkLabel(master=toplevel, text="Finished!")
        fin.place(relx=0.5, rely=0.5, anchor=tkinter.CENTER)
    else:
        toplevel.title("Error")
        avoid = customtkinter.CTkLabel(master=toplevel, text="すべての設定を完了してください")
        avoid.place(relx=0.5, rely=0.5, anchor=tkinter.CENTER)


button = customtkinter.CTkButton(master=app, text="Process", command=button_func)
button.place(relx=0.5, rely=0.8, anchor=tkinter.CENTER)


app.mainloop()
