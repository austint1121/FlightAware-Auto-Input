import time  # Allows use of the sleep function
import tkinter
from tkinter import filedialog
from openpyxl import load_workbook  # Allows Python to load Excel workbooks.
from selenium import webdriver  # Selenium powers the autobrowser
from selenium.webdriver.common.keys import Keys  # Allows sending of keyboard presses to HTML elements


def GetFileLocation():
    """Function for obtaining the location of the user's excel file"""
    global flocation
    flocation = filedialog.askopenfilename()
    print(flocation)
    file_path.destroy()


def FlightAware():
    """Function that uses the Chrome automated web browser to enter flight numbers on flightaware."""
    user = FA_user.get()
    passw = FA_pass.get()
    global flocation
    # Loads the Excel doc, sets up variables.
    workbook = load_workbook(filename=flocation)
    finallist = {""}
    sheet = workbook.active

    for row in sheet.values:
        for value in row:
            finallist.add(value)

    # Start the autobrowser and then go to the login page of Flightaware
    driver = webdriver.Chrome()
    driver.get("https://flightaware.com/account/session")

    # Targets the username and password boxes on the webpage and inputs the credentials for the account
    username = driver.find_element_by_name('flightaware_username')
    username.send_keys(user)
    password = driver.find_element_by_name('flightaware_password')
    password.send_keys(passw)
    password.send_keys(Keys.ENTER)

    # Goes to the flight tracking management page, then targets the box to add flights
    driver.get("https://flightaware.com/me/manage")
    aircraft = driver.find_element_by_id('add_ident')

    # Enters flight numbers into box, presses enter to submit them.
    count = 0
    for items in finallist:
        aircraft.send_keys(items)
        aircraft.send_keys(Keys.ENTER)
        time.sleep(1)
        count += 1
    print(f'Completed {count} flights')
    main_window.destroy()


if __name__ == '__main__':
    main_window = tkinter.Tk()
    main_window.title('FlightAware Auto Input')

    user_label = tkinter.Label(main_window, text='FlightAware Username')
    FA_user = tkinter.Entry(main_window, width=50)
    pass_label = tkinter.Label(main_window, text='FlightAware Password')
    FA_pass = tkinter.Entry(main_window, width=50)
    file_path = tkinter.Button(main_window, text='select file location', command=GetFileLocation)
    close = tkinter.Button(main_window, text='Continue', command=FlightAware)

    user_label.grid(row=0, column=0, padx=5, pady=5)
    FA_user.grid(row=0, column=1, padx=5, pady=5)
    pass_label.grid(row=1, column=0, padx=5, pady=5)
    FA_pass.grid(row=1, column=1, padx=5, pady=5)
    file_path.grid(row=2, column=0, padx=5, pady=5)
    close.grid(row=2, column=1, padx=5, pady=5)
    main_window.mainloop()
