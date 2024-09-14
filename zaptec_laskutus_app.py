# Python app to connect zaptec portal and get charging history data
# Also able to give pool electric price file for invoice
# Author: Kari Sivonen
# Date: 2024-06-11
# Version: 1.0

import requests, json, xmltodict
import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter.messagebox import showerror
import os
from datetime import datetime, timedelta, date, time
import openpyxl


# Create the main window
root = tk.Tk()
root.title("Zapter Laskutus Appi")
root.geometry("800x400")
print("Tkinter version:", tk.TkVersion)
# Menu bar
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)

#ttk.checkbox variable need to be defined globally
unchecked = {}


#App layout
# Right frame for Output data
right_frame = ttk.Frame(root, borderwidth=2)
right_frame.pack(fill=tk.Y, side=tk.RIGHT, padx=5, pady=5)

timeframe_label = ttk.Label(right_frame, text="Laskutus ajanjakso:")
timeframe_label.pack()

console_label = ttk.Label(right_frame, text="Konsoli:")
console_label.pack()

output_text = tk.Text(right_frame)
output_text.pack(fill=tk.X)

output_text.insert("1.0", "Tervetuloa Vanhojen siilojen laskutus appiin\n")

# Left frame for other data data
left_frame = ttk.Frame(root, borderwidth=2)
left_frame.pack(fill=tk.Y, side=tk.LEFT, padx=5, pady=5)


#Change apikeys to zaptec and entsoe classes which are used in the app
class zaptec:
    def __init__(self):
        self.apikey = "eyJhbGciOiJSUzI1NiIsImtpZCI6Ijc0OUVEQTU0REQ2MzRCOUE5NDM0QjFDODVDNTFBNDEzQzlGMUMwNTkiLCJ4NXQiOiJkSjdhVk4xalM1cVVOTEhJWEZHa0U4bnh3RmsiLCJ0eXAiOiJhdCtqd3QifQ.eyJzdWIiOiIzN2Q5NjNhMS1lMDlkLTRlM2MtODNjYy03NWZjMTU5YzE3Y2EiLCJuYW1lIjoiS2FyaSBTaXZvbmVuICIsImVtYWlsIjoia2FyaW1hdHRpLnNpdm9uZW5AZ21haWwuY29tIiwiYXV0aF90aW1lIjoiMjAyNC0wOS0xMFQxNDoyMDozOC4xODMyNjkxWiIsImF1ZCI6IjQxNGUxOTI3YTM4ODRmNjhhYmM3OWY3MjgzODM3ZmQxIiwib2lfdGtuX2lkIjoiZmNhNDBjZTItY2FmOS00OWI0LWRjOTAtMDhkY2NlODRlNjRhIiwianRpIjoiZjlkZjkyNjMtOTM0Yi00OTY1LTg2MTgtYmQ4MjcxODBkYTIxIiwiZXhwIjoxNzI2MDY0NDM4LCJpc3MiOiJodHRwczovL2FwaS56YXB0ZWMuY29tLyIsImlhdCI6MTcyNTk3ODAzOH0.AbkYvskJXMRD73QtTxOudgNqhklSWV0GLMeA0fvEYjULhEOZiXgAlo_q8uo2F328Svz2i7FXyqIrVtiLMvsUKQzlZpZnoGc4eyUgkQ8Q9T9Hak1kDzJGGQkdya8zc5p2r6dWBUV1MLZIzYTHGbre-ROzNLJcfgjE8YB4I9Kvg2R8hLOii3CFoV5Cypr_DD1tVCpwXuCouLnoie9GkYHDMNUoZlilViVheI2AUqx7_jfl2IJSXWp9nCYzkIfK95BTrn4GrkYcSVqJuKml-OUchfzHYmV71iv2rOfFU7r8nMj6u65ZZXihttWQPf4Pst50GaINZR-i0yhM0dgJdBNVVWsrz7MrGDDw9qDM-YOeZWYl79qFDgkejAPL1cOtiT6y5Yh01dCeRZ-B9pl5lbvo2BRC1UUDBUgIPp0qHNL8oQVYB502qw3TyUyisapigTNYWsYfs1b_CkULK4GOJ9AUboDLzNPyaafYq3wKDPU0r7A_gGaYD0q7t-7hPzMzZHoxq9bY99JoYFEZBwSBoAugAGPaRK-h34DBloq6fcMjxmMNqx_DDU39TqRvwLz2AxQetYFuILpyU3mCDs85o4lCT7YVuqPgzLtriqTiR1C81kEqDq8F3mxpdqVytcd-BWJFwARRTdGQKFYFLaF3Xv1BXl126zgwpIEQcF887eGiM80"
        self.headers = {
                'accept': 'text/plain',
                'Authorization' : 'Bearer ' + self.apikey
        }
        self.output = None

        self.chargers = {}
        self.chargers_url = "https://api.zaptec.com/api/chargers?ReturnIdNameOnly=true"

        #Tunnit pitää olla muuten ei ota sen päivän tietoja. From päivä voi olla ilman tunteja
        self.firsthoururl = "00%3A00%3A00.000"
        self.lasthoururl = "23%3A59%3A00.000"
        self.chargerHistories = {}
        self.chargerHistory_url = "https://api.zaptec.com/api/chargehistory?ChargerId=%s&From=%sT%s&To=%sT%s&GroupBy=0&DetailLevel=1"
        #self.chargerHistory_url = "https://api.zaptec.com/api/chargehistory?ChargerId=5bdfee4e-9b03-4662-9258-dd1b14ebf54b&From=2024-05-01&To=2024-05-31&GroupBy=0&DetailLevel=1"

    def storeApikey(self, apikey):
        if apikey != "":
            self.apikey = apikey
            self.headers = {
                'accept': 'text/plain',
                'Authorization' : 'Bearer ' + self.apikey
            }

    def getApikey(self):
        return self.apikey

    def addLogger(self, outputText):
        self.output = outputText

    def GetChargersInfo(self):
        if os.path.exists("chargersInfo.json"):
            with open("chargersInfo.json", "r") as file:
                self.chargers = json.load(file)
            return True
        else:
            response = requests.get(self.chargers_url, headers=self.headers)
            if response.status_code == 200:
                #For debugging
                #If file exist read it

                self.chargers = response.json()
                self.output.insert(tk.END, "Laturitiedot haettu onnistuneesti\n")
                #For debugging
                with open("chargersInfo.json", "w") as write_file:
                    json.dump(self.chargers, write_file, indent=4)
                return True
            else:
                self.output.insert(tk.END, "Error: Request failed with status code %s and reason: %s\n"%(response.status_code, response.reason))
                return response

    def GetChargeHistory(self, chargerName, chargerId, startTime, endTime):
        print ("GetChargeHistory: %s, %s, %s, %s"%(chargerName, chargerId, startTime, endTime))
        fileName = "chargerHistory_%s_%s-%s.json"%(chargerName, startTime, endTime)
        if os.path.exists(fileName):
            with open(fileName, "r") as file:
                self.chargerHistories[chargerName] = json.load(file)
            return True

        else:

            print ("requesturl: %s"%(self.chargerHistory_url%(chargerId, startTime, self.firsthoururl, endTime, self.lasthoururl)))
            #print ("Request header: %s"%self.headers)
            response = requests.get(self.chargerHistory_url%(chargerId, startTime, self.firsthoururl, endTime, self.lasthoururl), headers=self.headers)
            if response.status_code == 200:
                #For debugging
                #If file exist read it

                self.chargerHistories[chargerName] = response.json()
                self.output.insert(tk.END, "Laturitiedot haettu onnistuneesti\n")
                #For debugging
                with open(fileName, "w") as write_file:
                    json.dump(self.chargerHistories[chargerName], write_file, indent=4)

                return True
            else:
                showerror(title='Entsoe portal error', message="Request failed with status code %s and reason: %s\n"%(response.status_code, response.reason))
                return response

class entsoe:
    def __init__(self):
        self.apikey = "a1a19280-3b60-4a66-89ee-8ee9b58e577c"
        self.dayahead_url = "https://web-api.tp.entsoe.eu/api?securityToken=%s&documentType=A44&out_Domain=10YFI-1--------U&in_Domain=10YFI-1--------U&periodStart=%s&periodEnd=%s"
        self.dayahead_dict = {}
        self.output = None
        self.prices_dict = {}


    def getDayAheadData(self,startTime, endTime):

        #If file exist read it
        startimestring = "%s0000"%startTime.strftime("%Y%m%d")
        endTimeString = "%s0000"%endTime.strftime("%Y%m%d")
        if os.path.exists("entsoe_%s_%s.json" % (startimestring, endTimeString)):
            with open("entsoe_%s_%s.json" % (startimestring, endTimeString), "r") as file:
                self.dayahead_dict = json.load(file)
        else:
            url = self.dayahead_url % (self.apikey, startimestring, endTimeString)
            headers = {}
            print ("url: %s"%url)
            response = requests.request("GET", url, headers=headers)

            if response.status_code == 200:
                self.dayahead_dict = xmltodict.parse(response.text)

                print ("Write the json data to output")
                # json file
                with open("entsoe_%s_%s.json" % (startimestring, endTimeString), "w") as write_file:
                    json.dump(self.dayahead_dict, write_file, indent=4)
            else:
                showerror(title='Entsoe portal error', message="Request failed with status code %s and reason: %s\n"%(response.status_code, response.reason))
                return response


        print("Start price time dict handling")
        for TimeSeries in self.dayahead_dict["Publication_MarketDocument"]["TimeSeries"] :
            periodStartTime = datetime.fromisoformat(TimeSeries["Period"]["timeInterval"]["start"])
            periodEndTime = datetime.fromisoformat(TimeSeries["Period"]["timeInterval"]["end"])
            print("Get time from period from %s to %s"%(periodStartTime, periodEndTime))

            for point in TimeSeries["Period"]["Point"] :
                positiontime = periodStartTime + timedelta(hours=int(int(point["position"])-1))
                #print ("%s"%positiontime.strftime("%Y%m%d%H"))
                #print ("Price: %s Euros / MWH"%point["price.amount"])
                self.prices_dict[positiontime.strftime("%Y%m%d%H")] = float(point["price.amount"])/1000

        print("Price dict handled")
        #print(self.prices_dict)
        return True

    def getPriceOfHour(self, date):
        #print("getPriceOfHour for %s %s"%(date.strftime("%Y%m%d%H"), hour))
        global tax
        return (self.prices_dict[date.strftime("%Y%m%d%H")]*tax)

    def addLogger(self, outputText):
        self.output = outputText

    def storeApikey(self, apikey):
        print("Entsoe apikey: \"%s\""%apikey)
        if apikey != "":
            self.apikey = apikey

    def getApikey(self):
        return self.apikey

zaptecApi = zaptec()
entsoeApi = entsoe()

#From and to dates
fromDate = date.fromisoformat('2024-07-01')
toDate = date.fromisoformat('2024-07-31')

#Billing info
tax = 1.24
transferPrice = 0.0
margin = 0.5


def apikey_window():
    # implementation of the login window.
    login_window = tk.Toplevel(root)
    login_window.title("Anna Apikeyt")
    login_window.geometry("200x150")

    #Add request for apikeys
    label1 = ttk.Label(login_window, text="Zaptec apikey: ")
    label1.pack()
    apikey1Entry = ttk.Entry(login_window, show='*')
    apikey1Entry.focus()
    apikey1Entry.pack()

    label2 = ttk.Label(login_window, text="Entsoe apikey: ")
    label2.pack()
    apikey2Entry = ttk.Entry(login_window, show='*')
    apikey2Entry.pack()

    def close_and_output():
        global zaptecApi
        zaptecApi.storeApikey(apikey1Entry.get())
        output_text.insert(tk.END, "Zaptek apikey: %s\n"%zaptecApi.getApikey())
        global entsoeApi
        entsoeApi.storeApikey(apikey2Entry.get())
        output_text.insert(tk.END, "Entsoe apikey: %s\n"%entsoeApi.getApikey())
        login_window.destroy()

    # login button
    submit_button = ttk.Button(login_window, text="Hyväksy", command=close_and_output)
    submit_button.pack()

def timeframe_window():
    # implementation of the login window.
    login_window = tk.Toplevel(root)
    login_window.title("Laskutus ajanjakso")
    login_window.geometry("200x150")

    #Add request for dates
    fromdateLabel = ttk.Label(login_window, text="Alkaen: ")
    fromdateLabel.pack()
    calFrom = DateEntry(login_window, width=12, background='darkblue',
                    foreground='white', borderwidth=2)
    calFrom.pack(padx=10, pady=10)

    toDateLabel = ttk.Label(login_window, text="Päättyy: ")
    toDateLabel.pack()
    calTo = DateEntry(login_window, width=12, background='darkblue',
                    foreground='white', borderwidth=2)
    calTo.pack(padx=10, pady=10)

    def close_and_output():
        global fromDate
        fromDate = calFrom.get_date()
        output_text.insert(tk.END, "Laskutus jakso alkaa : %s\n"%fromDate)
        global toDate
        toDate = calTo.get_date()
        output_text.insert(tk.END, "Laskutus jakso päättyy: %s\n"%toDate)
        login_window.destroy()
        timeframe_label["text"] = "Laskutus ajanjakso: %s - %s"%(fromDate, toDate)

    # login button
    submit_button = ttk.Button(login_window, text="Hyväksy", command=close_and_output)
    submit_button.pack()

def billingInfo():
    # implementation of the login window.
    billingInfo_window = tk.Toplevel(root)
    billingInfo_window.title("Hinta tiedot")
    billingInfo_window.geometry("300x250")

    taxLabel = ttk.Label(billingInfo_window, text="Vero: ")
    taxLabel.pack()
    taxEntry = ttk.Entry(billingInfo_window)
    taxEntry.focus()
    taxEntry.pack()

    marginLabel = ttk.Label(billingInfo_window, text="Marginaali: ")
    marginLabel.pack()
    marginEntry = ttk.Entry(billingInfo_window)
    marginEntry.pack()


    transferPriceLabel = ttk.Label(billingInfo_window, text="Siirto hinta: ")
    transferPriceLabel.pack()
    transferPriceEntry = ttk.Entry(billingInfo_window)
    transferPriceEntry.pack()

    def close_and_output():
        global tax
        tax = taxEntry.get()
        output_text.insert(tk.END, "Vero: %s\n"%tax)
        global margin
        margin = marginEntry.get()
        output_text.insert(tk.END, "Marginaali: %s\n"%margin)
        global transferPrice
        transferPrice = transferPriceEntry.get()
        output_text.insert(tk.END, "Siirto hinta: %s\n"%transferPrice)
        billingInfo_window.destroy()


    # login button
    submit_button = ttk.Button(billingInfo_window, text="Hyväksy", command=close_and_output)
    submit_button.pack()

settings_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Asetukset", menu=settings_menu)
settings_menu.add_command(label="Avaimet", command=apikey_window )
settings_menu.add_command(label="Ajanjakso", command=timeframe_window )
settings_menu.add_command(label="Maksutiedot", command=billingInfo )

def print_apikey():
    global zaptecApi
    global entsoeApi
    output_text.insert(tk.END, "Zaptec Apikey: %s\n"%zaptecApi.getApikey())
    output_text.insert(tk.END, "Entsoe Apikey: %s\n"%entsoeApi.getApikey())

settings_menu.add_command(label="Näytä Avaimet", command=print_apikey )

def isDateInWinterPriceTime(daytime):
    winterDateStart = datetime(year=daytime.year, month=11, day=1, tzinfo=daytime.tzinfo)
    winterDateEnd = datetime(year=daytime.year, month=3, day=31, tzinfo=daytime.tzinfo)
    winterHourStart = datetime(year=daytime.year, month=daytime.month, day=daytime.day, hour=7, tzinfo=daytime.tzinfo)
    winterHourEnd = datetime(year=daytime.year, month=daytime.month, day=daytime.day,hour=22, tzinfo=daytime.tzinfo)



    #If date is between 31.3 - 1.11
    if winterDateEnd < daytime < winterDateStart:
        print ("Date is not in winter time")
        return False
    else:
        if winterHourStart < daytime < winterHourEnd :
            return True
        else:
            return False

def calculate_invoice():
    #Implement here invoice generation. Connect to entsoe and zaptec portal for data and calculate all eslected chargers invoice
    global fromDate
    global toDate
    print (type(fromDate))
    print (fromDate)
    print (toDate)

    if  fromDate == None or toDate == None:
        showerror(title='Laskutus virhe', message="Anna laskutus ajanjakso\n")
        return False

    entsoeApi.addLogger(output_text)
    entsoeApi.getDayAheadData(fromDate, toDate)

    wb = openpyxl.Workbook()
    summarySheet = wb.active
    summarySheet.title = "Yhteenveto"
    summarySheet['A1'] = "Ajanjakso:"
    summarySheet['B1'] = "%s - %s"%(fromDate, toDate)
    summarySheet["B2"] = "Talvipäivä"
    summarySheet["B3"] = "Muu aika"

    summarySheet["A3"] = "Kokonaiskulutus:"
    summarySheet["A4"] = "Keskihinta:"
    summarySheet["A5"] = "Kokonaishinta:"

    summarySheet['A7'] = "Laturi"
    summarySheet['B7'] = "Kokonaiskulutus"
    summarySheet['C7'] = "Keskihinta"
    summarySheet['D7'] = "Hinta yhteensä"

    totalEnergy = 0.0
    totalPrice = 0.0

    chargerSheets= {}
    #Get all selected chargers
    chargerIndex = 0
    for widget in left_frame.winfo_children():
        if isinstance(widget, ttk.Checkbutton):
            #print("%s: %s"%(widget["text"], widget.state()))
            if widget.instate(['selected']):
                chargerName = widget["text"]
                print(chargerName)
                #Get charger id
                for charger in zaptecApi.chargers["Data"]:
                    if charger["Name"] == chargerName:
                        chargerIndex += 1
                        chargerId = charger["Id"]
                        #print(chargerId)
                        chargerSheets[chargerName] = wb.create_sheet(chargerName)
                        #Get charger history
                        return_value = zaptecApi.GetChargeHistory(chargerName, chargerId, fromDate, toDate)
                        if (return_value):
                            chargerHistory = zaptecApi.chargerHistories[chargerName]
                            print("Now we have Charger usage history so we can check all the data and calculate invoice")
                            # Create invoice file
                            # Loop all data from chargerHistory and calculate invoice
                            # For each hour charger has energy, get the price from entsoe and calculate the invoice

                            chargerSheets[chargerName].cell(row=1, column=1, value="%s lataustiedot aikavälillä %s - %s"%(chargerName, fromDate, toDate))

                            chargerSheets[chargerName].cell(row=2, column=2, value="Talvipäivä")
                            chargerSheets[chargerName].cell(row=2, column=3, value="Muu aika")
                            chargerSheets[chargerName].cell(row=3, column=1, value="Kokonaiskulutus:")
                            chargerSheets[chargerName].cell(row=4, column=1, value="Keskihinta:")
                            chargerSheets[chargerName].cell(row=5, column=1, value="hinta yhteensä:")

                            chargerSheets[chargerName].cell(row=7, column=1, value="Latausjaksojen yhteenveto:")
                            header = ["jakso", "Aikaväli", "Kulutus", "keskihinta", "hinta", "Kulutus talviaikana", "Kulutus muuna aikana"]
                            chargerSheets[chargerName].append(header)

                            sessioIndex = 0
                            ChargerTotalEnergy = 0.0
                            chargerTotalPrice = 0.0
                            ChargerTotalEnergyWinter = 0.0

                            for session in chargerHistory["Data"]:
                                sessioIndex += 1

                                sessionTotalEnergy = float(session["Energy"])
                                if sessionTotalEnergy == 0.0:
                                    continue

                                #This have to change as sesson can be in winter and non winter time
                                #We need total energy and winter enrgy. Total energy includes also winter energy as it is calulated for spot price.abs
                                #Winter eneggy is calculated only for transfer price
                                ChargerTotalEnergy = ChargerTotalEnergy + sessionTotalEnergy
                                totalEnergyFromHours = 0.0
                                totalEnergyPriceFromHours = 0.0
                                averagePriceForSessionEnergy = 0.0
                                totalEnergyFromHoursWinter = 0.0

                                for EnergyDetails in session["EnergyDetails"]:
                                    timestamp = EnergyDetails["Timestamp"]
                                    format_data = "%Y-%m-%dT%H:%M:%S.%f%z"

                                    try:
                                        daytime = datetime.strptime(timestamp, format_data)
                                    except:
                                        daytime = datetime.strptime(timestamp, "%Y-%m-%dT%H:%M:%S%z")

                                    energy = float(EnergyDetails["Energy"])

                                    minutes = daytime.minute
                                    if minutes == 0 :
                                        daytime = daytime - timedelta(hours=1)
                                    price = float(entsoeApi.getPriceOfHour(daytime))

                                    #Check if time is in winter price time
                                    if isDateInWinterPriceTime(daytime) :
                                        print (daytime)
                                        print ("is on winter time")
                                        totalEnergyFromHoursWinter = (totalEnergyFromHoursWinter + energy)
                                    else:
                                        print ("is not on winter time")

                                    totalEnergyFromHours = (totalEnergyFromHours + energy)
                                    totalEnergyPriceFromHours = totalEnergyPriceFromHours + (energy*price)

                                averagePriceForSessionEnergy = totalEnergyPriceFromHours/totalEnergyFromHours
                                chargerSheets[chargerName].append([sessioIndex, "%s - %s"%(session["StartDateTime"], session["EndDateTime"]), sessionTotalEnergy, averagePriceForSessionEnergy, (sessionTotalEnergy*averagePriceForSessionEnergy), totalEnergyFromHoursWinter, (totalEnergyFromHours - totalEnergyFromHoursWinter)])
                                chargerTotalPrice = chargerTotalPrice + totalEnergyPriceFromHours
                                ChargerTotalEnergyWinter = ChargerTotalEnergyWinter + totalEnergyFromHoursWinter

                            summaryRow=chargerIndex+6
                            #print("Summary row: %s"%summaryRow)

                            totalEnergy = totalEnergy + ChargerTotalEnergy
                            totalPrice = totalPrice + chargerTotalPrice
                            if sessioIndex > 0:

                                chargerSheets[chargerName].cell(row=3, column=2, value=ChargerTotalEnergy)
                                chargerSheets[chargerName].cell(row=3, column=3, value=ChargerTotalEnergyWinter)
                                chargerSheets[chargerName].cell(row=3, column=4, value=(ChargerTotalEnergy - ChargerTotalEnergyWinter))
                                if ChargerTotalEnergy != 0 :
                                    chargerSheets[chargerName].cell(row=4, column=2, value=(chargerTotalPrice/ChargerTotalEnergy))
                                chargerSheets[chargerName].cell(row=5, column=2, value=chargerTotalPrice)


                                summarySheet.cell(row=summaryRow, column=1, value=chargerName)
                                summarySheet.cell(row=summaryRow, column=2, value=ChargerTotalEnergy)
                                if ChargerTotalEnergy != 0 :
                                    summarySheet.cell(row=summaryRow, column=3, value=(chargerTotalPrice/ChargerTotalEnergy))
                                summarySheet.cell(row=summaryRow, column=4, value=chargerTotalPrice)
                            else:
                                summarySheet.cell(row=summaryRow, column=1, value=chargerName)
                                summarySheet.cell(row=summaryRow, column=2, value=0)
                                summarySheet.cell(row=summaryRow, column=3, value=0)
                                summarySheet.cell(row=summaryRow, column=4, value=0)



                        else:
                            showerror(title='Zaptec portal error', message="Request failed with status code %s and reason: %s\n"%(return_value.status_code, return_value.reason))
                            return False


    summarySheet["A2"] = "Kokonaiskulutus:"
    summarySheet["B2"] = totalEnergy
    summarySheet["A3"] = "Keskihinta:"
    summarySheet["B3"] = totalPrice/totalEnergy
    summarySheet["A4"] = "Kokonaishinta"
    summarySheet["B4"] = totalPrice

    wb.save("Lasku %s - %s.xlsx"%(fromDate,toDate))
    output_text.insert(tk.END, "Lasku %s - %s.xlsx luotu\n"%(fromDate,toDate))

def show_chargers():
    # implementation of the show_chargers function. It have to show all chargers as selectable list
    output_text.insert(tk.END, "Haetaan laturien tietoja... \n")
    zaptecApi.addLogger(output_text)
    return_value = zaptecApi.GetChargersInfo()
    if (return_value):
        charger_list = zaptecApi.chargers["Data"]
    else:
        #showerror(title='Zaptec portal error', message="Request failed with status code %s and reason: %s\n"%(return_value.status_code, return_value.reason))
        return False

    ttk.Label(left_frame,text="Laturit:").pack()

    #Empty leftframe
    for widget in left_frame.winfo_children():
        if isinstance(widget, ttk.Button):
            if widget["text"] == "Hae laturien tiedot":
                pass
            else:
                 widget.destroy()
        else:
            widget.destroy()

    for charger in charger_list:
        global unchecked
        unchecked[charger["Name"]] = tk.BooleanVar(value=True)
        check_button = ttk.Checkbutton(left_frame,text=str(charger["Name"]),variable=unchecked[charger["Name"]] )
        check_button.pack()

    #Invoice button
    invoice_button = ttk.Button(left_frame, text="Luo laskut", command=calculate_invoice)
    invoice_button.pack()

#Charger button
charger_button = ttk.Button(left_frame, text="Hae laturien tiedot", command=show_chargers)
charger_button.pack()

# Start the application's main loop
root.mainloop()