import time
from datetime import datetime
import pyautogui
pyautogui.FAILSAFE = True
import requests
import urllib
import re
from bs4 import BeautifulSoup as bs
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from contextlib import redirect_stdout
import pandas as pd
import sys
import sqlite3
from collections import defaultdict
import os


#___________________________________ LOGGING _______________________________
''' Standardfunktionen die keiner speziellen Klasse zuzuordnen sind
    print_: Leitet Konsolenausgabe zusätzlich in textdatei um
    input_: Leitet Konsoleneingabe zusätzlich in textdatei um
    dateiname = Standardpfad in dem die Log-Dateien abgelegt werden'''

def createTimeStamp():
    secondsSinceEpoch = time.time()
    timeObj = time.localtime(secondsSinceEpoch)
    timestamp = '%d%d%d_%d_%d_%d' % (timeObj.tm_year, timeObj.tm_mon, timeObj.tm_mday, timeObj.tm_hour, timeObj.tm_min, timeObj.tm_sec)
    return timestamp

def getYear():
    secondsSinceEpoch = time.time()
    timeObj = time.localtime(secondsSinceEpoch)
    return '%d' % (timeObj.tm_year)

def getMonth():
    secondsSinceEpoch = time.time()
    timeObj = time.localtime(secondsSinceEpoch)
    monat = '%d' % (timeObj.tm_mon)
    switcher = {
        1: "januar",
        2: "februar",
        3: "maerz",
        4: "april",
        5: "mai",
        6: "juni",
        7: "juli",
        8: "august",
        9: "september",
        10: "oktober",
        11: "november",
        12: "dezember"
    }
    return switcher.get(int(monat), " - ")


def getDate():
    secondsSinceEpoch = time.time()
    timeObj = time.localtime(secondsSinceEpoch)
    return '%d-%d-%d' % (timeObj.tm_year, timeObj.tm_mon, timeObj.tm_mday)

dateiname = './logs/lizenzenUpdateLog_' + createTimeStamp() + '.txt'

def print_(text, logdatei = dateiname):
    print(text)
    with open(logdatei, 'a') as f:
        with redirect_stdout(f):
            print(text)

def input_():
    eingabe = input();
    print_(eingabe);
    return eingabe;

#_____________________________ INFOSERVER SCRAPING ____________________________

class Scraper():
    ''' Scraper Klasse mit Standardmässigen Scraping-Funktionen '''

    def __init__(self, url):
        self.screen_width = pyautogui.size()[0]
        self.screen_height = pyautogui.size()[1]
        self.now = datetime.now()
        self.startzeit = self.now.strftime("%H:%M:%S")
        print_("Web-Scraper-Instanz um "+str(self.startzeit)+" angelegt...")
        self.url = url
        self.start()

    def start(self):
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        self.driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=options)
        self.driver.get(self.url)
        self.wait = WebDriverWait(self.driver, 10)

    def login(self, user, password, submit):
        pyautogui.write(user)
        pyautogui.press("tab")
        pyautogui.write(password)
        time.sleep(1)
        self.clickByClass(submit)                   # submit = submitButton Merkmal
        print_("Eingeloggt als " + user + "...")
        pyautogui.sleep(1)

    def enterEmail(self, mail):
        atPosition = self.findATinMail(mail)
        pyautogui.write(mail[0:atPosition], 0.1)
        pyautogui.hotkey('ctrl','alt','q')
        pyautogui.write(mail[atPosition:len(mail)], 0.1)

    def findATinMail(self, mail):
        return mail.rfind("@")

    def clearField(self):
        pyautogui.hotkey('ctrl','a')
        time.sleep(0.5)
        pyautogui.press('delete')

    def iframe(self, frame = " "):
        ''' Wechseln eines IFRAMES durch Angabe der iFrame id '''

        self.driver.switch_to.frame(frame)
        print_("In das IFRAME '"+frame+"' gewechselt...")

    def multiPress(self, anzahl = 2, taste = 'tab'):
        while 0 < anzahl:
            pyautogui.press(taste, interval = 0.1)
            anzahl -= 1

    def waitForPageLoad(self, id, delay = 10):
        ''' Diese Funktion sucht nach einem element auf einer Seite und gibt True zurück, sobald es geladen wurde '''
        try:
            myElem = WebDriverWait(self.driver, delay).until(EC.presence_of_element_located((By.ID, id)))
            print_("Seite vollständig geladen...")
            return True
        except TimeoutException:
            print_('Laden der Seite hat zu lange gedauert...')
            return False

    def waitForPageLoad_xpath(self, xpath, delay = 10):
        ''' Diese Funktion sucht nach einem element auf einer Seite und gibt True zurück, sobald es geladen wurde '''
        try:
            myElem = WebDriverWait(self.driver, delay).until(EC.presence_of_element_located((By.XPATH, xpath)))
            print_("Seite vollständig geladen...")
            return True
        except TimeoutException:
            print_('Laden der Seite hat zu lange gedauert...')
            return False

    def centerMouse(self):
        pyautogui.moveTo(self.screen_width/2, self.screen_height/2, duration=0.2)

    def scroll(self, intensity):
        ''' positive Zahlen scroll hoch, negative Zahlen runter '''
        pyautogui.scroll(intensity)

    def clickByClass(self, klasse):
        wait = WebDriverWait(self.driver, 10)
        element = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, klasse)))
        element.click()

    def clickByLinkText(self, linktext):
        wait = WebDriverWait(self.driver, 10)
        element = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, linktext)))
        element.click()

    def clickByID(self, id):
        wait = WebDriverWait(self.driver, 10)
        element = wait.until(EC.element_to_be_clickable((By.ID, id)))
        element.click()

    def clickByName(self, name):
        wait = WebDriverWait(self.driver, 10)
        element = wait.until(EC.element_to_be_clickable((By.NAME, name)))
        element.click()

    def clickByXPath(self, xpath):
        wait = WebDriverWait(self.driver, 10)
        element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        element.click()

    def clickByXPath_fast(self, XPath):
        self.driver.implicitly_wait(3)
        self.driver.find_element_by_xpath(XPath).click()

    def findByXPath(self, XPath):
        return self.driver.find_elements_by_xpath(XPath)

    def write(self, text):
        pyautogui.write(text, 0.1)

    def wait(self, time = 0.5):
        time.sleep(time)


#------------------------------------------------------------

class InfoServerScraper(Scraper):
    ''' Klasse die speziell die Infoserver-Daten zieht, erbt Methoden/Attribute von Scraper Klasse '''

    def __init__(self, url = 'https://infoserver.ecs-deutschland.de/?'):
        super().__init__(url)
        self.now = datetime.now()
        self.startzeit = self.now.strftime("%H:%M:%S")
        self.fileManager = FileManager()
        print_("Beginne InfoServer-Scraping um: " + str(self.startzeit))

    def login(self, user, pw):
        super().login(user,pw,"buttonSpezial")

    def switchToOrder(self):
        ''' geht in den Reiter "Order" '''
        self.clickByLinkText('Order')

    def switchToSearch(self):
        #self.clickByXPath("//img[@title='Suche ein-/ausblenden']")
        search = pyautogui.locateCenterOnScreen('./img/infoserver_search.png', confidence=0.7)
        pyautogui.click(search)
        pyautogui.press("tab")
        print_("Suche geöffnet...")

    def switchToDateien(self):
        self.iframe('OrderDLG_IFRAME')
        self.clickByXPath("//td[contains(text(), 'Dateien')]")
        self.waitForPageLoad_xpath("//i[contains(text(), 'abrechnungsrelevante Dokumente')]")
        self.centerMouse()
        self.scroll(-300)
        print_("In den Reiter 'Dateien' gewechselt...")

    def downloadDateien(self):
        ''' Zählt alle Dateien durch und speichert "Dateiname: URL-Link" in einem Dictionary '''

        elementliste = self.dateien_Count()
        download_infos = {}
        counter = 1
        for webelement in elementliste:
            download_infos[webelement.text] = str(webelement.get_attribute('href'))
            print_('URL '+str(counter)+' = '+webelement.get_attribute('href')+"...")
            counter += 1

        for datei in download_infos:
            self.fileManager.saveDownloadedFile(self.auftragsnummer, download_infos[datei], datei)

    def dateien_Count(self):
        urls = self.findByXPath("//div[@id='currentOrderFilesTab']/table/tbody/tr/td/a[@href!='#']")
        print_(str(len(urls))+ " Dateien zur Auftragsnummer '"+str(self.auftragsnummer)+"' gefunden...")
        return urls

    def auftragsnummerEingeben(self,auftragsnummer):
        ''' gibt Auftragsnummer ein und klickt, falls vorhanden, auf das passende Ergebnis '''

        self.auftragsnummer = auftragsnummer # speichert die Auftragsnummer so lange, bis sie bei
        # Bearbeitung des nächsten Auftrags mit dieser Objektinstanz überschrieben wird

        self.switchToSearch()
        self.write(auftragsnummer)
        print_('Auftragsnummer: '+str(auftragsnummer)+ " eingegeben...")
        if self.checkExistence():
            print_("Auftrag mit Auftragsnummer: " +auftragsnummer+" gefunden...")
            time.sleep(0.1)
            self.clickByXPath("//div[@id='ICD_OUTPUT']/table/tbody/tr[2]/td/table/tbody/tr[1]")
            print_('Auftragsfenster mit Auftragsnummer: '+auftragsnummer+" geöffnet...")
            pyautogui.click()
            time.sleep(0.5)
        else:
            pass

    def checkExistence(self):
        ''' Prüft erfolgreiche Suche beim Infoserver Lupen-Icon '''
        try:
            element = WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, "//div[@id='ICD_OUTPUT']/table/tbody/tr/td")))
            return True
        except selenium.common.exceptions.TimeoutException:
            return False


#_____________________________ EXCEL ZEITRAUM BESTIMMEN ____________________________

class DataHandler():
    ''' Zweck: Excel-Dateien, Bilder, Dokumente, geordnet in einer Datenbank ablegen.
        Klasse DataHandler stellt hauptsächlich Schnittstellen Funktionen zur Konvertierung
        von Excel-Dateien in Datenbanken und zur Bearbeitung mit Python bereit '''

    def __init__(self, dateipfad, tabellenname, datenbankname):
        print_("Lade "+tabellenname+"...")
        try:
            self.dataframe = pd.read_excel(open(dateipfad, 'rb'),sheet_name=tabellenname)
            print_(tabellenname+" vollständig geladen...")
        except FileNotFoundError:
            print_("Exceldatei: '"+dateipfad+ "' nicht gefunden!")
            print_(dateipfad+" nicht gefunden...")
            sys.exit()
        self.dateipfad = dateipfad
        self.tabellenname = tabellenname
        self.con = sqlite3.connect(datenbankname + '.db')
        self.cursor = self.con.cursor()

    def print(self):
        print_(self.dataframe)

    def to_sql(self, table):
        self.dataframe.to_sql(table, self.con, if_exists='replace', index = False)

    def createTable(self, table):
        ''' createTable legt eine SQL-Tabelle an, mit den aufbereiteten Spaltennamen der Excel-Datei. '''

        spaltennamen = [spalte for spalte in self.dataframe.columns]    # holt sich spaltennamen anhand des Dataframes (ursprüngliche Excel)
        bereinigte_spaltennamen = self.cleanNames(spaltennamen)         # bereinigt Namen durch Zeichen entfernen und Umlaute
        self.dataframe.columns = bereinigte_spaltennamen                # passt die Dataframe-Spaltennamen an die Namen aus der Liste an
        self.dataframe.loc[:,~self.dataframe.columns.duplicated()]      # sollten immer noch Duplikate vorkommen, werden sie gelöscht
        #self.createCreateStatement(table, list(self.dataframe.columns))
        self.to_sql(table)
        print_("SQL Tabelle '"+str(table)+"' aus Excel-Datei: " +str(self.dateipfad)+ ' / "'+ str(self.tabellenname) +'" angelegt...')

    # def scrub(self, table):
    #     # attributation: https://stackoverflow.com/a/3247553/7505395
    #     return ''.join( chr for chr in table if chr.isalnum() )
    #
    # def createCreateStatement(self, table, spalten):
    #     ''' Baut ein Create-Statement anhand von Listenwerte zusammen,
    #         sodass eine passende SQL-Tabelle angelegt werden kann '''
    #
    #     return f"create table {self.scrub(table)} ({spalten[0]}" + (",{} "*(len(spalten)-1)).format(*map(self.scrub,spalten[1:])) + ")"

    def cleanNames(self, namensliste):
        ''' Bereinigt Spaltennamen um Umlaute, Zeichen und Leerzeichen, Groß-Kleinschreibung und doppelte Einträge '''

        namensliste = [(lambda spalte: spalte.lower())(spalte) for spalte in namensliste]       # alles wird Klein geschrieben

        bereinigte_namensListe = []
        for spaltenName in namensliste:
            if "ä" in spaltenName:
                spaltenName = spaltenName.replace("ä","ae")
            if "ü" in spaltenName:
                spaltenName = spaltenName.replace("ü","ue")
            if "ö" in spaltenName:
                spaltenName = spaltenName.replace("ö","oe")
            if "ß" in spaltenName:
                spaltenName = spaltenName.replace("ß","ss")
            for zeichen in [".",",",";",":","/","-","!","?","#","§","$","%"]:
                # For Schleife ersetzt jedes Zeichen durch " "
                if  zeichen in spaltenName:
                    spaltenName = spaltenName.replace(zeichen,"_")
            if " " in spaltenName:
                spaltenName = spaltenName.replace(" ","")
            bereinigte_namensListe.append(spaltenName)

        duplikatBereinigte_spaltennamen = self.uniqify(bereinigte_namensListe) # bennent duplikate um (zB: reg, reg -> reg1, reg2)

        return duplikatBereinigte_spaltennamen

    def uniqify(self, mylist):
        ''' Ändert Duplikate einer Liste [a,a] zu [a1,a2] '''

        finalList = []
        dictCount = defaultdict(int)
        anotherDict = defaultdict(int)
        for t in mylist:
           anotherDict[t] += 1
        for m in mylist:
           dictCount[m] += 1
           if anotherDict[m] > 1:
               finalList.append(str(m)+str(dictCount[m]))
           else:
               finalList.append(m)
        return finalList

#------------------------------------------------------------

class DataHandler_Kauffrauenliste(DataHandler):
    ''' Zur Bewahrung der Trennung von Belangen bezieht sich dieser Data-Handler
        speziell auf die zu bearbeitende KauffrauenListe '''

    def __init__(self, dateipfad='./dokumente/kauffrauen_liste.xlsx', tabellenname='Auftragsbestandsliste', datenbankname='datenbasis'):
        ''' Da die Exceldatei: KauffrauenListe und die zugehörige Tabelle stets gleich bleibt werden die Vorgabewerte einfach direkt übergeben '''
        super().__init__(dateipfad, tabellenname, datenbankname)

    def getCurrentData(self, datum:str = '2020-09-01'):
        ''' Hier werden alle Zeilen aus der Kauffrauenliste geladen die für den Provisionierungmonat relevant sind,
            Basis dafür ist der Zahlungseingang der Kunden '''

        heute_datum = getDate()
        # Ines darauf Hinweisen Datumszeile für Datumswerte zu nutzen ...
        query = "SELECT * FROM kauffrauen_liste WHERE zahlungs_eingang > (?) "\
                "AND zahlungs_eingang != 'x' "\
                "AND zahlungs_eingang NOT LIKE ('%.%') "\
                "AND 'ja' NOT IN (zahlungs_eingang) "\
                "AND 'Lastschrift' NOT IN (zahlungs_eingang) "\
                "AND 'storno' NOT IN (zahlungs_eingang) "\
                "AND 'xx' NOT IN (zahlungs_eingang) "\
                "AND 'keine' NOT IN (zahlungs_eingang) "\
                "AND 'keine Bezahlung' NOT IN (zahlungs_eingang)"
        self.cursor.execute(query,(datum,))

        aktuelle_daten = self.cursor.fetchall()

        print_('Alle Daten deren Zahlungseingang zwischen dem ' +str(datum) +' und dem ' +str(heute_datum)+' liegen erfolgreich aus "kauffrauen_liste.xlsx" importiert...')
        print_('Anzahl erhaltener Daten: ' + str(len(aktuelle_daten)) +"...")

        aktuelles_dataframe = pd.DataFrame(aktuelle_daten)      # Rückgabewert ist ein Dataframe zur einfacheren Handhabung
        aktuelles_dataframe.columns = self.dataframe.columns    # beschriftet Spalten wie üblich
        return aktuelles_dataframe

    def getAuftragsnummern(self, datum = '2020-09-01'):
        currentData = self.getCurrentData(datum)
        return currentData['auftragsnummer'].to_list()

#_____________________________ INFOSERVER DATEIEN SPEICHERN ____________________________

class FileManager():
    ''' Diese Klasse bildet die Ordnerstruktur des Info-Servers lokal ab um die Dokumente,
        den Auftragsnummern korrekt zuordnen und letztendlich so die Kosten ermitteln zu können '''

    def __init__(self):
        self.doc_ordner = './dokumente_infoserver'
        self.jahresordner = self.doc_ordner +'/'+ str(getYear())
        self.monatsordner = self.jahresordner +'/'+ str(getMonth())
        self.init_structure()

    def init_structure(self):
        ''' legt die grundsätzliche Ordnerstruktur an, falls noch nicht vorhanden '''
        self.ordnerAnlegen(self.doc_ordner)
        self.ordnerAnlegen(self.jahresordner)
        self.ordnerAnlegen(self.monatsordner)

    def add_auftragsOrdner(self, auftragsnummer):
        ''' legt den ordner mit der auftragsnummer im passenden Ordner ab '''
        self.init_structure()
        self.ordnerAnlegen(self.monatsordner + "/" + str(auftragsnummer))

    def saveDownloadedFile(self, auftragsnummer, url, filename = None):
        ''' Speichert Datei vom Infoserver in einem passenden, mit Auftragsnummer beschrifteten
            Ordner und mit - im Infoverserver hinterlegten Dateinamen '''

        self.add_auftragsOrdner(auftragsnummer)
        if filename == None:
            # Falls kein Filename übergeben wurde, wird ein eigener anhand der URL generiert
            firstpos=url.rfind("/")                                         # ermittelt Slash
            lastpos=len(url)                                                # ermittelt Url Länge
            response = urllib.request.urlopen(url)                                      # Anfrage an Datei-URL
            dateiname = self.monatsordner + '/' + auftragsnummer + '/' + url[firstpos+1:lastpos]   # erstellt Dateinamen
        else:
            response = urllib.request.urlopen(url)
            dateiname = self.monatsordner + '/' + auftragsnummer + '/' + filename
        if not (self.isFile(dateiname)):
            #open(dateiname, 'wb').write(datei.content)                      # speichert Datei
            with open(dateiname,'wb') as f:
                f.write(response.read())
            print_("Datei '"+dateiname+"' im Ordner: '" +self.monatsordner+"/"+auftragsnummer+"' gespeichert...")
        else:
            print_("Datei '"+filename+"' bereits vorhanden...")
        time.sleep(2)


    def ordnerAnlegen(self, pfad):
        if not (self.isDirectory(pfad)): # falls Order noch nicht existiert
            os.makedirs(pfad)
            print_("Ordner im Pfad '"+pfad+"' angelegt...")

    def isDirectory(self, pfad):
        return os.path.isdir(pfad)

    def isFile(self, pfad):
        return os.path.isfile(pfad)




#======================= CONSOLE GUI FUNKTIONEN ========================

def bestaetigen(ueberschrift, text):
    print_('================================= ' + ueberschrift + ' =================================')
    print_(text + '(y/n)')
    eingabe = input_()
    if eingabe == "y" or eingabe == "Y" or eingabe == "":
        return True
    else:
        sys.exit()

#======================= Programmablauf ========================

# if bestaetigen('Pruefe Excel-Dokumente','KauffrauenListe aktuell?'):
#     kauffrauen_liste = DataHandler_Kauffrauenliste()        # erstellt Kauffrauenlisten Objekt
#     kauffrauen_liste.createTable('kauffrauen_liste')        # erstellt SQL-Tabelle aus Excel
#     auftragsnummern = kauffrauen_liste.getAuftragsnummern('2020-08-01') # relevante Auftragsnummern speichern
#
# print(auftragsnummern)

if bestaetigen('Datenscraping Infoserver','Infoserver-Scraping beginnen?'):
    scraper = InfoServerScraper()
    scraper.login("mikko.holfeld","Schweden21.")
    scraper.auftragsnummerEingeben("20181015_21")
    scraper.switchToDateien()
    scraper.downloadDateien()
