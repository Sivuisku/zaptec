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
        self.apikey = ""
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
        #global vat
        return (self.prices_dict[date.strftime("%Y%m%d%H")])

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
vat = 1.24
transferPrice = 0.92
transferPriceWinter = 1.31
margin = 0.49
basePriceTransfer = 0.0
tehoMaksu = 0.0
loisTehoMaksu = 0.0
basePrice = 0.0
energyTax = 2.24
huoltovarmuusmaksu = 0.01

def zapteclogin_window():
    # implementation of the login window.
    login_window = tk.Toplevel(root)
    login_window.title("Zaptec login")
    login_window.geometry("200x150")

    #Add request for apikeys
    label1 = ttk.Label(login_window, text="Zaptec käyttäjä ja salasana: ")
    label1.pack()
    usernameEntry = ttk.Entry(login_window)
    usernameEntry.focus()
    usernameEntry.pack()
    passwordEntry = ttk.Entry(login_window, show='*')
    passwordEntry.pack()

    def close_and_output():

        token_url = 'https://api.zaptec.com/oauth/token'

        # The payload for the token request
        payload = {
            'grant_type': 'password',
            'username': usernameEntry.get(),
            'password': passwordEntry.get()
        }

        # Make the token request
        response = requests.post(token_url, data=payload)

        # Check if the request was successful
        if response.status_code == 200:
            token_info = response.json()
            access_token = token_info['access_token']
            output_text.insert(tk.END,"Access Token: {access_token}")
            global zaptecApi
            zaptecApi.storeApikey(access_token)
        else:
            output_text.insert(tk.END,"Failed to obtain token: {response.status_code}")
            output_text.insert(tk.END,response.text)

        login_window.destroy()

    # login button
    submit_button = ttk.Button(login_window, text="Hyväksy", command=close_and_output)
    submit_button.pack()

def entsoekey_window():
    # implementation of the login window.
    login_window = tk.Toplevel(root)
    login_window.title("entsoe key")
    login_window.geometry("200x150")

    label = ttk.Label(login_window, text="Entsoe apikey: ")
    label.pack()
    apikeyEntry = ttk.Entry(login_window, show='*')
    apikeyEntry.pack()

    def close_and_output():
        global entsoeApi
        entsoeApi.storeApikey(apikeyEntry.get())
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

def billInfo():
    # implementation of the login window.
    billInfo_window = tk.Toplevel(root)
    billInfo_window.title("Laskun kiinteät hinnat (Ei kulutkseen perustuvat)")
    billInfo_window.geometry("400x400")

    global basePriceTransfer
    global tehoMaksu
    global loisTehoMaksu
    global basePrice

    basePriceTransferLabel = ttk.Label(billInfo_window, text="Pientehojännitemaksun perusmaksu (€) sisältää ALV:in")
    basePriceTransferLabel.pack()
    basePriceTransferEntry = ttk.Entry(billInfo_window)
    basePriceTransferEntry.delete(0, tk.END)
    basePriceTransferEntry.insert(0,basePriceTransfer)
    basePriceTransferEntry.focus()
    basePriceTransferEntry.pack()

    basePriceLabel = ttk.Label(billInfo_window, text="Kulutuksen perusmaksu (€) sisältää ALV:in ")
    basePriceLabel.pack()
    basePriceEntry = ttk.Entry(billInfo_window)
    basePriceEntry.delete(0, tk.END)
    basePriceEntry.insert(0,basePrice)
    basePriceEntry.focus()
    basePriceEntry.pack()

    tehoMaksuLabel = ttk.Label(billInfo_window, text="Tehomaksu (€): ")
    tehoMaksuLabel.pack()
    tehoMaksuEntry = ttk.Entry(billInfo_window)
    tehoMaksuEntry.delete(0, tk.END)
    tehoMaksuEntry.insert(0,tehoMaksu)
    tehoMaksuEntry.pack()

    loisTehoMaksuLabel = ttk.Label(billInfo_window, text="loistehomaksu (€): ")
    loisTehoMaksuLabel.pack()
    loisTehoMaksuEntry = ttk.Entry(billInfo_window)
    loisTehoMaksuEntry.delete(0, tk.END)
    loisTehoMaksuEntry.insert(0,loisTehoMaksu)
    loisTehoMaksuEntry.pack()

    def close_and_output():
        global basePriceTransfer
        basePriceTransfer = float(basePriceTransferEntry.get())
        output_text.insert(tk.END, "Perusmaksu liittymä: %s\n"%basePriceTransfer)
        global basePrice
        basePrice = float(basePriceEntry.get())
        output_text.insert(tk.END, "Perusmaksu: %s\n"%basePrice)
        global tehoMaksu
        global loisTehoMaksu
        tehoMaksu = float(tehoMaksuEntry.get())
        loisTehoMaksu = float(loisTehoMaksuEntry.get())
        output_text.insert(tk.END, "Tehomaksu Talvi: %s snt/kWh\n"%tehoMaksu)
        output_text.insert(tk.END, "loistehomaksu %s snt/kWh\n"%loisTehoMaksu)
        billInfo_window.destroy()


    # submit button
    submit_button = ttk.Button(billInfo_window, text="Hyväksy", command=close_and_output)
    submit_button.pack()

def pricesInfo():
    # implementation of the login window.
    pricesInfo_window = tk.Toplevel(root)
    pricesInfo_window.title("Hinta tiedot")
    pricesInfo_window.geometry("300x250")

    global vat
    global margin
    global transferPrice
    global transferPriceWinter
    global energyTax
    global huoltovarmuusmaksu

    print("vat: %s"%vat)

    vatLabel = ttk.Label(pricesInfo_window, text="Alv: ")
    vatLabel.pack()
    vatEntry = ttk.Entry(pricesInfo_window)
    vatEntry.delete(0, tk.END)
    vatEntry.insert(0,vat)
    vatEntry.focus()
    vatEntry.pack()

    marginLabel = ttk.Label(pricesInfo_window, text="Marginaali: ")
    marginLabel.pack()
    marginEntry = ttk.Entry(pricesInfo_window)
    marginEntry.delete(0, tk.END)
    marginEntry.insert(0,margin)
    marginEntry.pack()


    transferPriceLabel = ttk.Label(pricesInfo_window, text="Siirtohinta Talvi (snt/kWh): ")
    transferPriceLabel.pack()
    transferPriceEntry = ttk.Entry(pricesInfo_window)
    transferPriceEntry.delete(0, tk.END)
    transferPriceEntry.insert(0,transferPrice)
    transferPriceEntry.pack()
    transferPriceLabelWinter = ttk.Label(pricesInfo_window, text="Siirtohinta muu aika (snt/kWh): ")
    transferPriceLabelWinter.pack()
    transferPriceEntryWinter = ttk.Entry(pricesInfo_window)
    transferPriceEntryWinter.delete(0, tk.END)
    transferPriceEntryWinter.insert(0,transferPriceWinter)
    transferPriceEntryWinter.pack()

    energyTaxLabel = ttk.Label(pricesInfo_window, text="Energia vero (snt/kWh): ")
    energyTaxLabel.pack()
    energyTaxEntry = ttk.Entry(pricesInfo_window)
    energyTaxEntry.delete(0, tk.END)
    energyTaxEntry.insert(0,energyTax)
    energyTaxEntry.pack()

    huoltovarmuusmaksuLabel = ttk.Label(pricesInfo_window, text="Huoltovarmuusmaksu (snt/kWh): ")
    huoltovarmuusmaksuLabel.pack()
    huoltovarmuusmaksuEntry = ttk.Entry(pricesInfo_window)
    huoltovarmuusmaksuEntry.delete(0, tk.END)
    huoltovarmuusmaksuEntry.insert(0,huoltovarmuusmaksu)
    huoltovarmuusmaksuEntry.pack()


    def close_and_output():
        global vat
        vat = float(vatEntry.get())
        output_text.insert(tk.END, "Vero: %s\n"%vat)
        global margin
        margin = float(marginEntry.get())
        output_text.insert(tk.END, "Marginaali: %s\n"%margin)
        global transferPrice
        global transferPriceWinter
        transferPrice = float(transferPriceEntry.get())
        transferPriceWinter = float(transferPriceEntryWinter.get())
        output_text.insert(tk.END, "Siirtohinta Talvi: %s snt/kWh\n"%transferPriceWinter)
        output_text.insert(tk.END, "Siirtohinta muu aika %s snt/kWh\n"%transferPrice)
        global energyTax
        energyTax = float(energyTaxEntry.get())
        output_text.insert(tk.END, "Energia vero: %s snt/kWh\n"%energyTax)
        global huoltovarmuusmaksu
        huoltovarmuusmaksu = float(huoltovarmuusmaksuEntry.get())
        output_text.insert(tk.END, "Huoltovarmuusmaksu: %s snt/kWh\n"%huoltovarmuusmaksu)
        pricesInfo_window.destroy()


    # submit button
    submit_button = ttk.Button(pricesInfo_window, text="Hyväksy", command=close_and_output)
    submit_button.pack()

settings_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Asetukset", menu=settings_menu)
settings_menu.add_command(label="Zaptec login", command=zapteclogin_window )
settings_menu.add_command(label="entsoe avain", command=entsoekey_window )
settings_menu.add_command(label="Ajanjakso", command=timeframe_window )
settings_menu.add_command(label="Hintatiedot", command=pricesInfo )
settings_menu.add_command(label="Laskun tiedot", command=billInfo )

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
        #print ("Date is not in winter time")
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
    global transferPrice
    global transferPriceWinter
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
    summarySheet["C2"] = "Talvipäivä (%s snt/kWh)"%transferPriceWinter
    summarySheet["D2"] = "muu aika (%s snt/kWh)"%transferPrice

    summarySheet["A3"] = "Kokonaiskulutus (kWh):"
    summarySheet["A4"] = "Pörssisähkön Keskihinta (€/kWh):"
    summarySheet["A5"] = "Pörssisähkön Kokonaishinta (€):"
    summarySheet["A6"] = "Siirtomaksu (€):"

    summarySheet['A8'] = "Laturi"
    summarySheet['B8'] = "Kokonaiskulutus (kWh)"
    summarySheet['C8'] = "Keskihinta (€/kWh)"
    summarySheet['D8'] = "Siirtohinta Talviaika (€)"
    summarySheet['E8'] = "Siirtohinta muu aika (€)"
    summarySheet['F8'] = "Marginaali (€)"
    summarySheet['F8'] = "Energiavero (€)"
    summarySheet['G8'] = "ALV (€)"
    summarySheet['H8'] = "Kulutuksen hinta (€)"
    summarySheet['H8'] = "Kiinteät kulut (€)"
    summarySheet['H8'] = "Yhteensä (€)"

    totalEnergy = 0.0
    totalEnergyWinter = 0.0
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

                            chargerSheets[chargerName].cell(row=3, column=3, value="Talvipäivä")
                            chargerSheets[chargerName].cell(row=3, column=4, value="Muu aika")
                            chargerSheets[chargerName].cell(row=4, column=1, value="Kokonaiskulutus:")
                            chargerSheets[chargerName].cell(row=5, column=1, value="Keskihinta:")
                            chargerSheets[chargerName].cell(row=6, column=1, value="hinta yhteensä:")

                            chargerSheets[chargerName].cell(row=8, column=1, value="Latausjaksojen yhteenveto:")
                            header = ["jakso", "Aikaväli", "Kulutus (kWh)", "keskihinta (€/kWh)", "hinta (€)", "Kulutus talviaikana (kWh)", "Kulutus muuna aikana (kWh)","","","Aika","Kulutus","hinta"]
                            chargerSheets[chargerName].append(header)

                            sessioIndex = 0
                            ChargerTotalEnergy = 0.0
                            chargerTotalPrice = 0.0
                            ChargerTotalEnergyWinter = 0.0
                            hourPriceTimeColumn = 10
                            hourPriceUsageColumn = 11
                            hourPriceColumn = 12
                            hourPriceInfoRow = 10

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

                                    #Charger session can be from one day to another day started before time period we are interesting
                                    #and ended after time period we are interesting
                                    #We have to check if time is in the time period we are interesting
                                    if daytime.date() < fromDate or daytime.date() > toDate:
                                        print("Time is not in the time period we are interesting: %s"%daytime)
                                        continue

                                    price = float(entsoeApi.getPriceOfHour(daytime))

                                    chargerSheets[chargerName].cell(row=hourPriceInfoRow, column=hourPriceTimeColumn, value=timestamp)
                                    chargerSheets[chargerName].cell(row=hourPriceInfoRow, column=hourPriceUsageColumn, value=energy)
                                    chargerSheets[chargerName].cell(row=hourPriceInfoRow, column=hourPriceColumn, value=price)
                                    hourPriceInfoRow += 1

                                    #Check if time is in winter price time
                                    if isDateInWinterPriceTime(daytime) :
                                        print (daytime)
                                        print ("is on winter time")
                                        totalEnergyFromHoursWinter = (totalEnergyFromHoursWinter + energy)
                                    #else:
                                    #    print ("is not on winter time")

                                    totalEnergyFromHours = (totalEnergyFromHours + energy)
                                    totalEnergyPriceFromHours = totalEnergyPriceFromHours + (energy*price)

                                if totalEnergyFromHours == 0:
                                    averagePriceForSessionEnergy = 0
                                else:
                                    averagePriceForSessionEnergy = totalEnergyPriceFromHours/totalEnergyFromHours

                                chargerSheets[chargerName].append([sessioIndex, "%s - %s"%(session["StartDateTime"], session["EndDateTime"]), sessionTotalEnergy, averagePriceForSessionEnergy, (sessionTotalEnergy*averagePriceForSessionEnergy), totalEnergyFromHoursWinter, (totalEnergyFromHours - totalEnergyFromHoursWinter)])
                                chargerTotalPrice = chargerTotalPrice + totalEnergyPriceFromHours
                                ChargerTotalEnergyWinter = ChargerTotalEnergyWinter + totalEnergyFromHoursWinter

                            summaryRow=chargerIndex+8
                            #print("Summary row: %s"%summaryRow)

                            totalEnergy = totalEnergy + ChargerTotalEnergy
                            totalEnergyWinter = totalEnergyWinter + ChargerTotalEnergyWinter
                            totalPrice = totalPrice + chargerTotalPrice
                            if sessioIndex > 0:

                                chargerSheets[chargerName].cell(row=4, column=2, value=ChargerTotalEnergy)
                                chargerSheets[chargerName].cell(row=4, column=3, value=ChargerTotalEnergyWinter)
                                chargerSheets[chargerName].cell(row=4, column=4, value=(ChargerTotalEnergy - ChargerTotalEnergyWinter))
                                if ChargerTotalEnergy != 0 :
                                    chargerSheets[chargerName].cell(row=5, column=2, value=(chargerTotalPrice/ChargerTotalEnergy))
                                chargerSheets[chargerName].cell(row=6, column=2, value=chargerTotalPrice)


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


    summarySheet["B3"] = totalEnergy
    summarySheet["C3"] = totalEnergyWinter
    summarySheet["D3"] = (totalEnergy -totalEnergyWinter)
    summarySheet["B4"] = totalPrice/totalEnergy
    summarySheet["B5"] = totalPrice
    summarySheet["C6"] = (totalEnergyWinter*transferPriceWinter)
    summarySheet["D6"] = ((totalEnergy - totalEnergyWinter)*transferPrice)/100

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