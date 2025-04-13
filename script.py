import time
import pandas as pd
import pyautogui
import pyperclip
import os
import pywhatkit as kit

# Read Excel File
for i in range(1,5):                # Number of files in your Folder
    addr = f"Your_File_address_{i}.xlsx"
    df = pd.read_excel(addr, skiprows=2)  # U can skip number of rows that are used as heading if needed
    print(df.columns.tolist())

    df.columns = df.columns.str.strip()  # Remove extra spaces
    df = df.filter(items=[               # Choose only the rewuired columns
        "Cell No", "Name", "Desig", "per month salary", "w.days", 
        "Basic pay", "HRA", "Spl. Allo", "Earnings", "pf", "esi",             
        "adv", "mess", "store", "other", "EB", "Total Dedu", "Net"
    ])


    # Open WhatsApp Desktop
    os.startfile("whatsapp://")
    time.sleep(2)

    # Iterate over DataFrame rows
    for cellNo, name, desig, perMonth, days, basicPay, hra, splAllo, gross, epf, esi, adv, mess, store, other, eb, totalDed, netRound in zip(
            df["Cell No"], df["Name"], df["Desig"], df["per month salary"], df["w.days"], df["Basic pay"],
            df["HRA"], df["Spl. Allo"], df["Earnings"], df["pf"], df["esi"], df["adv"], df["mess"],
            df["store"], df["other"], df["EB"], df["Total Dedu"], df["Net"]):

        
        # Open new chat
        pyautogui.hotkey("ctrl", "n")
        time.sleep(1)
        pyautogui.write(str(cellNo))
        time.sleep(0.5)
        pyautogui.press("tab")
        time.sleep(0.5)
        pyautogui.press("tab")
        time.sleep(0.5)
        pyautogui.press("enter")
        time.sleep(1)

        # Create message string
        s = f"""```{name.center(38)}

    Number of Days       : {days}
    Gross                : {gross:.2f}
    Net                  : {netRound:.2f}
    Total Deduction      : {totalDed:.2f}
    EPF                  : {epf:.2f}
    ESI                  : {esi:.2f}
    Advance              : {adv:.2f}
    Mess                 : {mess:.2f}
    Store                : {store:.2f}
    Others               : {other:.2f}```"""




        print(s)  # Debugging

        # Copy to clipboard and paste
        pyperclip.copy(s)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(1)
        pyautogui.press("enter")
        time.sleep(0.5)
        pyautogui.press("esc")
        # Delay before next message
        time.sleep(1)
