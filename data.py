#!/usr/bin/python
# -*- coding: utf-8 -*-
import pickle
import tkinter.messagebox as mb
import data as d
import tercard
import datetime
import time
import os
from glob import glob
import urllib.request
import zipfile
import _thread
import locale
from xlwt import Formula
langFile=[]
from contacts import Mod

class Ter(): # territory class

    def __init__(self, number="", type="", address="", note="", image="", map="", work=[]):
        self.number=number
        self.type=type
        self.address=address
        self.note=note
        self.map=map
        self.image=image
        self.works=[]
        self.extra=[]        
        if len(work)>0:
            self.works=[]
            self.works.append([work[0][0], work[0][1], work[0][2]])
        
    def retrieve(self, root):
        """ How to show individual ter line in list """        
        output=""
        try:           ###################################################### 
            if root.lines.get()==1: line="│"
            else: line=" "
            if root.fields[0].get()==1:
                if self.getStatus(root)==0:     output += "v"#"√"
                elif self.getStatus(root)==1:   output += "o"
                else:                           output += "!"#"‼"
            if root.fields[1].get()==1:         output += "%s№%-6s" % (line, self.number[:6])
            if root.fields[2].get()==1:         output += "%s%-8s" % (line, self.type[:8])        
            if root.fields[3].get()==1:
                if root.doubleAddress.get()==0: output += "%s%-40s" % (line, self.address[:40])
                else:                           output += "%s%-80s" % (line, self.address[:80]) 
            if root.fields[4].get()==1:         output += "%s%-17s" % (line, self.getCurrentPublisher()[:17])
            if root.fields[5].get()==1:         output += "%s%-9s" % (line, dateFilter(root, self.getDateLastSubmit()[:9]))
            if root.fields[6].get()==1:         output += "%s%-3s" % (line, str(self.getWorks()))        
            if root.fields[7].get()==1:         output += "%s%s " % (line, self.note)
            if root.contactsEnabled.get()==1 and Mod==True and len(self.extra)>0 and len(self.extra[0])!=0: output += "( %d к.)" % (len(self.extra[0]))
        except: print("output error")
        return output

    def show(self, root, new=False):
        card=tercard.Tercard(self, root, new=new)
        if card.saved==False:
            if card.savedAction==True:
                result, root.db, root.settings=d.load(root)
                root.log(root.msg[332] % self.number)
                root.updateS()
            elif card.new==True:
                del root.db[len(root.db)-1] # if creation of new ter is cancelled without saving, delete it (by last index)            
        
    def getStatus(self, root):
        if len(self.works)==0 or self.getDate2()!="": return 0
        elif self.getDate2()=="" and self.getDelta1()<root.timeoutDays: return 1
        else: return 2
        
    def getWorks(self):
        """ Return number of works """
        if len(self.works)==0: result=0
        elif self.works[len(self.works)-1][2]!="": result=len(self.works)
        else: result=len(self.works)-1
        return result
        
    def getPublisher(self):
        """ Return publisher of last opened work, whether finished or not """
        if len(self.works)==0: return ""
        else: return self.works[len(self.works)-1][0]
        
    def getPublisherFinished(self):
        """ Return publisher of last finished work """
        if self.getWorks()==0: return ""
        elif self.works[len(self.works)-1][2]!="": return self.works[len(self.works)-1][0]
        else: return self.works[len(self.works)-2][0]
        
    def getCurrentPublisher(self):
        """ Return current publisher, if he works on ter """
        if self.getDate1()!="" and self.getDate2()=="": return self.getPublisher()
        else: return ""
       
    def getDate1(self):
        """ Return date 1 of last work, if exists """
        if len(self.works)==0: return ""
        else: return self.works[len(self.works)-1][1]
        
    def getDate2(self):
        """ Return date 2 of last work, if exists """
        if len(self.works)==0: return ""
        else: return self.works[len(self.works)-1][2]
        
    def getDateLastSubmit(self):
        """ Return last submission date """
        if len(self.works)==0: return ""
        elif self.works[len(self.works)-1][2]!="": return self.works[len(self.works)-1][2]
        elif len(self.works)>1 and self.works[len(self.works)-2][2]!="": return self.works[len(self.works)-2][2]
        else: return ""
        
    def getDate2Prev(self):
        """ Return date 2 of previous to last work, if exists """
        return self.works[len(self.works)-2][2]
        
    def give(self, root, silent=False, fromTerCard=False):
        done=False
        if silent==False and root.chosenPublisher.get().strip()=="": root.setPublisher()
        if root.actPrompts.get()==1:
            answer=mb.askyesno(root.msg[333], root.msg[334] % (self.number, root.chosenPublisher.get(), root.chosenDate.get().strip()))
        else: answer=True
        if answer==True:
            self.works.append([root.chosenPublisher.get(), root.chosenDate.get().strip(), ""])
            root.log(root.msg[335] % (self.number, root.chosenPublisher.get(), root.chosenDate.get().strip()))
            if fromTerCard==False: root.save()
            done=True
        return done
        
    def submit(self, root, silent=False, fromTerCard=False):
        done=False
        if silent==False:
            if root.actPrompts.get()==1: answer=mb.askyesno(root.msg[336], root.msg[337] % (self.number, self.getPublisher(), root.chosenDate.get()))
            else: answer=True
            if answer==True:
                self.works[len(self.works)-1][2]=root.chosenDate.get().strip()
                if fromTerCard==False: root.save()
                root.log(root.msg[338] % (self.number, self.getPublisher(), root.chosenDate.get()))
                done=True
        else: self.works[len(self.works)-1][2]=root.chosenDate.get().strip()
        return done

    def getDelta1(self):
        """ Calculates number of days since last date1 of selected ter """
        try: 
            d0 = datetime.date( int(d.convert(self.getDate1())[0:4]), int(d.convert(self.getDate1())[5:7]), int(d.convert(self.getDate1())[8:10]) )
            ds = time.strftime("%Y-%m-%d", time.localtime())
            d1 = datetime.date( int(ds[0:4]), int(ds[5:7]), int(ds[8:10]) )
            return (d1-d0).days
        except: return 999999
        
    def getDelta2(self):
        """ Calculates number of days since last date2 of selected ter """
        try: 
            d0 = datetime.date( int(d.convert(self.getDate2())[0:4]), int(d.convert(self.getDate2())[5:7]), int(d.convert(self.getDate2())[8:10]) )
            ds = time.strftime("%Y-%m-%d", time.localtime())
            d1 = datetime.date( int(ds[0:4]), int(ds[5:7]), int(ds[8:10]) )
            return (d1-d0).days
        except:
            try:
                d0 = datetime.date( int(d.convert(self.getDate2Prev())[0:4]), int(d.convert(self.getDate2Prev())[5:7]), int(d.convert(self.getDate2Prev())[8:10]) )
                ds = time.strftime("%Y-%m-%d", time.localtime())
                d1 = datetime.date( int(ds[0:4]), int(ds[5:7]), int(ds[8:10]) )
                return (d1-d0).days
            except: return 999999
            
    def getAverageWork(self):
        """ Return average number of days between works of this ter (as list)"""        
        average=[]
        if self.getWorks()!=0:            
            for i in range(len(self.works)):
                if self.works[i][2]!="":
                    d0 = datetime.date( int(d.convert(self.works[i][1])[0:4]), int(d.convert(self.works[i][1])[5:7]), int(d.convert(self.works[i][1])[8:10])) 
                    d1 = datetime.date( int(d.convert(self.works[i][2])[0:4]), int(d.convert(self.works[i][2])[5:7]), int(d.convert(self.works[i][2])[8:10]))
                    average.append((d1-d0).days)
        return average
            
    def exportXLS_List(self, root):
        if self.getCurrentPublisher()!="":
            publisher=self.getCurrentPublisher()
            date1=self.getDate1()+" "
            date2=""
        else:
            publisher=self.getPublisher()
            date1=self.getDate1()+" "
            date2=self.getDate2()+" "
        if publisher=="": date1=""
        if root.images.get()==1: return ["%s\u00A0" % self.number, self.type+"\u00A0", self.address+"\u00A0", self.note+"\u00A0", Formula("HYPERLINK(\"%s\u00A0\";\"%s\")" % (self.map, self.map)), self.image+"\u00A0", publisher+"\u00A0", date1+"\u00A0", date2+"\u00A0"]
        else: return ["%s\u00A0" % self.number, self.type+"\u00A0", self.address+"\u00A0", self.note+"\u00A0", Formula("HYPERLINK(\"%s\u00A0\";\"%s\")" % (self.map, self.map)), publisher+"\u00A0", date1+"\u00A0", date2+"\u00A0"]                                                                      #  if ter pics disabled, don't export
        
    def exportXLS_S13(self):
        return self.number

def load(root, filename="", loadSettings=True, loadDB=True):  
    result=False
    db=[]
    settings=[]
    
    if loadDB==True:                                                            # load database 
        if filename=="": filename="core.hal"
        try:
            with open(filename, "rb") as f: db=pickle.load(f)                   
            if len(db)!=0 and db[0].address=="" and db[0].number=="" and db[0].type=="": pass
            result=True
        except:
            print("corrupt or absent file, create blank array")                 # if none, create blank list
    
    if loadSettings==True:                                                      # create default settings first    
        settings.append(1)                                                      # 0 default sort type
        settings.append(1)                                                      # 1 auto update
        settings.append(1)                                                      # 2 grid on list
        settings.append(1)                                                      # 3 images in ters
        settings.append(1)                                                      # 4 search in history
        settings.append(1)                                                      # 5 bottom scrollbar
        settings.append(0)                                                      # 6 lines in grid
        settings.append("01010100")                                             # 7 table fields
        settings.append(getWinFonts()[0])                                       # 8 list font
        settings.append("9")                                                    # 9 list font size
        settings.append(1)                                                      # 10 double address field
        settings.append(0)                                                      # 11 worked ter should be given within year (inactive for now)
        settings.append(0)                                                      # 12 note in card as text field
        settings.append(1)                                                      # 13 create new ter on Insert        
        settings.append(1)                                                      # 14 show splash screen
        settings.append(180)                                                    # 15 days of timedout ters
        settings.append(0)                                                      # 16 exact search
        settings.append("")                                                     # 17 search range
        settings.append(5000)                                                   # 18 log length
        settings.append(1)                                                      # 19 prompts on give/submit
        settings.append(1)                                                      # 20 save window geometry
        settings.append("")                                                     # 21 images folder
        settings.append(1)                                                      # 22 contacts module
        
        # Try to load existing settings    
        try:
            with open("settings.ini", "r", encoding="utf-8") as file: content=file.read()            
            for i in range(len(settings)):
                try: settings[i]=getString(content, i)
                except: print("setting %d not found" % i) 
        except: print("settings file not found")
        else: settings=checkOldSettings(settings)
        
    return result, db, settings

def checkOldSettings(settings):
    """Load settings in legacy format, deactivate after some time"""
    with open("settings.ini", "r", encoding="utf-8") as file: content=file.read()#content = [line.rstrip() for line in file]
    try:
        if "List grid" in content:          settings[2]=content[content.index("List grid=")+10]
        if "Images in cards" in content:    settings[3]=content[content.index("Images in cards=")+16]
        if "Search in history" in content:  settings[4]=content[content.index("Search in history=")+18]
        if "Bottom scrollbar" in content:   settings[5]=content[content.index("Bottom scrollbar=")+17]
        if "Lines in grid" in content:      settings[6]=content[content.index("Lines in grid=")+14]
        if "Note as text" in content:       settings[12]=content[content.index("Note as text=")+13]
        if "Create new on Insert" in content:settings[13]=content[content.index("Create new on Insert=")+21]
        if "Show splash" in content:        settings[14]=content[content.index("Show splash=")+12]
        if "Timeout days" in content:       settings[15]=content[ content.index("Timeout days=")+13 : content.index("Timeout days=")+17 ]
        if "Exact search" in content:       settings[16]=content[content.index("Exact search=")+13]
        if "Search range" in content:       settings[17]=content[content.index("Search range=")+1]
        if "Log length" in content:         settings[18]=content[ content.index("Log length=")+11 : content.index("Log length=")+15 ]
        if "Prompts on acts" in content:    settings[19]=content[content.index("Prompts on acts=")+16]
    except: print("some error loading legacy settings")    
    return settings

def convert(d):
    """Convert DD.MM.YY date into YYYY-MM-DD"""    
    try: return "20"+d[6]+d[7]+"-"+d[3]+d[4]+"-"+d[0]+d[1]
    except: return ""
    
def convertBack(root, d):
    """Convert YYYY-MM-DD date into DD.MM.YY date"""        
    try: result = d[8]+d[9]+"."+d[5]+d[6]+"."+d[2]+d[3]
    except: result="01.01.01"
    else: 
        if verifyDate(root, result, silent=True)==False: result="01.01.01"
    return result

def ifInt(char):
    """Check if value is integer"""    
    try: int(char) + 1
    except: return False
    else: return True
    
def dateFilter(root, date):
    """Gets regular date (DD.MM.YY) and modifies it according to current language (root.msg[0])"""
    if root.msg[0]=="ru": return date
    elif root.msg[0]=="en" and date!="": return date[3]+date[4]+"/"+date[0]+date[1]+"/"+date[6]+date[7]    
    else: return date
    
def dateEnToRu(date):
    """Convert date format from English into Russian"""    
    if date=="": return ""
    return date[3]+date[4]+"."+date[0]+date[1]+"."+date[6]+date[7]

def verifyDate(root, date, silent=False):
    """Return False if date is incorrect, and shows warning"""    
    if root.msg[0]=="en": date=dateEnToRu(date)    
    try:
        if  ifInt(date[0])==True and\
            ifInt(date[1])==True and\
            date[2]=="." and\
            ifInt(date[3])==True and\
            ifInt(date[4])==True and\
            date[5]=="." and\
            ifInt(date[6])==True and\
            ifInt(date[7])==True and\
            int(date[0]+date[1])<=31 and\
            int(date[3]+date[4])<=12 and\
            len(date)==8:        
                check=True
        else: check=False
    except: check=False    
    if check==False and silent==False: mb.showwarning(root.msg[339], root.msg[340])
    return check
    
def verifyDateYYYY(date):
    """The same as verifyDate, but with YYYY (only for Excel imports), and returns truncated date to DD.MM.YY"""
    correct=False    
    if len(date)==8:    
        try:
            if  ifInt(date[0])==True and\
                ifInt(date[1])==True and\
                date[2]=="." and\
                ifInt(date[3])==True and\
                ifInt(date[4])==True and\
                date[5]=="." and\
                ifInt(date[6])==True and\
                int(date[0]+date[1])<=31 and\
                int(date[3]+date[4])<=12 and\
                ifInt(date[7])==True: 
                    correct=True
            else: pass
        except: pass    
    return correct
    
def updateApp(root):
    def run(threadName, delay):
        try:
            print("checking updates")
            file=urllib.request.urlopen("https://raw.githubusercontent.com/antorix/Halieus/master/version.txt")
        except:
            print("update check failed")
            return
        else:
            version=str(file.read())
            newversion=[int(version[2]), int(version[4]), int(version[6: len(version)-3 ])]
            with open("Halieus.pyw", "r", encoding="utf-8") as file: content = [line.rstrip() for line in file]
            thisversion=[int(content[0][10:][0]), int(content[0][10:][2]), int(content[0][10:][4: len(content)-3 ])]    
            print(thisversion)
            print(newversion)
            if newversion>thisversion:
                print("new version found!")
                try: urllib.request.urlretrieve("http://github.com/antorix/Halieus/raw/master/update.zip", "update.zip")
                except: print("download failed")
                else: print("download complete")                    
                try:
                    zip=zipfile.ZipFile("update.zip", "r")
                    zip.extractall("")
                    zip.close()
                    os.remove("update.zip")
                except: print("unpacking failed")
                else:
                    print("unpacking complete")
                    mb.showinfo(root.msg[341], root.msg[342])
                    root.log(root.msg[343])
                    deleteObsolete()                 
                    if os.path.exists("contacts_mod.py"):                       # separate download, if present in folder
                        try: urllib.request.urlretrieve("https://raw.githubusercontent.com/antorix/Halieus/master/contacts_mod.py", "contacts_mod.py")
                        except: print("contacts_mod module downloaded successfully")
                        else: print("contacts_mod module download error")
            else: print("no new version found")
    try: _thread.start_new_thread(run,("Thread-Load", 1,))
    except: print("can't run thread")

def getDelta(date):
    """ Calculates number of days since given date """
    try: 
        d0 = datetime.date( int(d.convert(date)[0:4]), int(d.convert(date)[5:7]), int(d.convert(date)[8:10]) )
        ds = time.strftime("%Y-%m-%d", time.localtime())
        d1 = datetime.date( int(ds[0:4]), int(ds[5:7]), int(ds[8:10]) )            
        return (d1-d0).days
    except: return 999999

def deleteObsolete():
    """Delete obsolete files from older versions"""
    print("trying to delete obsolete files")
    try:        
        for f in glob ("images/*.png"): os.remove(f)
    except: pass
    try:
        for f in glob ("images/*.ico"): os.remove(f)
    except: pass
    files=[]
    files.append("images/png_file24.gif")
    files.append("images/db_export16.gif")
    files.append("images/db_save16.gif")
    files.append("images/gear16.gif")
    files.append("images/gear12.gif")
    files.append("images/lens13.gif")
    files.append("images/timeout16.gif")
    files.append("images/book16.gif")
    files.append("images/notebook16.gif")
    files.append("icons.py")
    files.append("lang/ru.txt")
    for file in files:
        if os.path.exists(file): os.remove(file)

def getWinFonts():
    """Try to get a nice font on Windows as default, return first font or Courier if none found"""
    fonts=[]
    if os.name=="nt":
        try:
            osFonts=os.listdir("C:/WINDOWS/fonts")
            if "LiberationMono-Regular.ttf" in osFonts: fonts.append("Liberation Mono")
            if "DejaVuSansMono_0.ttf" in osFonts: fonts.append("DejaVu Sans Mono")
            if "lucon.ttf" in osFonts: fonts.append("Lucida Console")
            if "cousine-regular.ttf" in osFonts: fonts.append("Cousine")
            if "firamono-regular.ttf" in osFonts: fonts.append("Fira Mono")
            if "PTM55F.ttf" in osFonts: fonts.append("PT Mono")
            if "ubuntumono-r.ttf" in osFonts: fonts.append("Ubuntu Mono")            
        except: print("no good font found on Windows")
    fonts.append("Courier New")
    return fonts

def getString(input, key):
    """Returns string from XML input (text string) between <a001></a001> according to key"""
    return input[input.index("<a%03d>" % key)+6 : input.index("</a%03d>" % key)]

def loadLanguage(preroot=False):
    """Get language and load all strings"""
    global langFile
    while 1:
        try:
            with open("lang.ini", "r", encoding="utf-8") as f:
                if preroot==True: selectedLanguage=f.read(500)
                else: selectedLanguage=f.read().strip()
        except:
            if locale.getdefaultlocale()[0][0:2]=="ru":     selectedLanguage="ru"
            elif locale.getdefaultlocale()[0][0:2]=="en":   selectedLanguage="en"
            else: selectedLanguage="en"
            print("can't find lang.ini, set system locale or en")
            with open("lang.ini", "w", encoding="utf-8") as f: f.write(selectedLanguage)
        try:        
            with open("lang/%s.xml" % selectedLanguage, "r", encoding="utf-8") as f: langFile = f.read()
        except:
            print("can't find language file, exit")
            break
        else: break                                                             # load strings into msg array        
    msg=[]
    if preroot==True:                                                           # load strings only for pre-root operations
        msg.append(getString(langFile, 0))
        msg.append(getString(langFile, 1))
        msg.append(getString(langFile, 2))
        msg.append(getString(langFile, 106))
        msg.append(getString(langFile, 102))        
    else:
        key=0
        while 1:            
            try: msg.append(getString(langFile, key))
            except: break
            else:
                #print("%d: %s" % (key, msg[key]))
                key+=1
    return msg
