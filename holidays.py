# coding: utf-8
# Script to mark calendar as busy during holidays
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")

# Defining Holidays
holidays = {
    "Makara Sankranti" : ["2020-01-15", 0],
    "Maha Shivarathri" : ["2020-02-21", 0],
    "Ugadi" : ["2020-03-25", 0],
    "May Day" : ["2020-05-01", 0],
    "Barkid" : ["2020-07-31", 0],
    "Gandhi Jayanthi" : ["2020-10-02", 0],
    "Kanakadasa Jayanthi" : ["2020-12-03", 0],
    "Christmas Day" : ["2020-12-25", 0],
    "Christmas Day" : ["2019-12-25", 3],
}

# Creating Events for Holidays
# Busy status is set to Out of office during these days
for key, value in holidays.items():
    app = outlook.CreateItem(1)
    app.Start = value[0]
    app.Subject = key
    app.AllDayEvent = True
    app.BusyStatus = value[1]
    app.Save()
