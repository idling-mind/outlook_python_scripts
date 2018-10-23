# coding: utf-8
# Script to mark calendar as busy during holidays
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")

# Defining Holidays
holidays = {
    "May Day" : "2019-05-01",
    "Idul Fitr" : "2019-06-05",
    "Independence Day" : "2019-08-15",
    "Ganesha Charurthi" : "2019-09-02",
    "Mahatma Gandhi Jayanti": "2019-10-02",
    "Diwali" : "2019-10-29",
    "Kannada Rajyotsava": "2019-11-01",
    "Christmas Day" : "2019-12-25"
}

# Defining Optional Holidays
optional_holidays = {
    "Maha Shivarathri" : "2019-03-04",
    "Mahaveera Jayanti" : "2019-04-17",
    "Good Friday": "2019-04-19",
    "Bakrid" : "2019-08-12",
    "Ayudha Pooja" : "2019-10-07",
    "Vijaya Dashami" : "2019-10-08",
    "Kanakadaasa Jayanti" : "2019-11-15"
}

# Creating Events for Holidays
# Busy status is set to Out of office during these days
for key, value in holidays.items():
    app = outlook.CreateItem(1)
    app.Start = value
    app.Subject = key
    app.AllDayEvent = True
    app.BusyStatus = 3
    app.Save()

# Creating events for options holidays.
# The Busy status is set to Free during these days
for key, value in optional_holidays.items():
    app = outlook.CreateItem(1)
    app.Start = value
    app.Subject = key
    app.AllDayEvent = True
    app.BusyStatus = 0
    app.Save()

