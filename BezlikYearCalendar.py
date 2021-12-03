#!/usr/bin/python
#-*- coding: utf-8 -*-
from __future__ import division # overrules Python 2 integer division
import sys
import locale
import calendar
import datetime
from datetime import date, timedelta
import csv
import platform

try:
    from scribus import *
except ImportError:
    print("This Python script is written for the Scribus \
      scripting interface.")
    print("It can only be run from within Scribus.")
    sys.exit(1)

os = platform.system()
if os != "Windows" and os != "Linux":
    print("Your Operating System is not supported by this script.")
    messageBox("Script failed",
        "Your Operating System is not supported by this script.",
        ICON_CRITICAL)	
    sys.exit(1)

python_version = platform.python_version()
if python_version[0:1] != "3":
    print("This script runs only with Python 3.")
    messageBox("Script failed",
        "This script runs only with Python 3.",
        ICON_CRITICAL)	
    sys.exit(1)

try:
    from tkinter import * # python 3
    from tkinter import messagebox, filedialog, font
except ImportError:
    print("This script requires Python Tkinter properly installed.")
    messageBox('Script failed',
               'This script requires Python Tkinter properly installed.',
               ICON_CRITICAL)
    sys.exit(1)



localization = [
    ['English', 'CP1252', 'en_US.UTF8'],
    ['German', 'CP1252', 'de_DE.UTF8'],
    ['Polish', 'CP1250', 'pl_PL.UTF8'],
    ['Russian', 'CP1251', 'ru_RU.UTF8'],
]

months_names = [
    "Styczeń",
    "Luty",
    "Marzec",
    "Kwiecień",
    "Maj",
    "Czerwiec",
    "Lipiec",
    "Sierpień",
    "Wrzesień",
    "Październik",
    "Listopad",
    "Grudzień"
]

class ScYearCalendar:

    def __init__(
            self,
            year,
            lang='English',
            holidaysList=list()
    ):
        self.year = year
        self.months = [month for month in range(1, 13)]
        self.nrVmonths = 12
        self.holidaysList = holidaysList
        if len(self.holidaysList) != 0:
            self.drawHolidays = True
        else:
            self.drawHolidays = False
        self.lang = lang
        self.calUniCode = "UTF-8"
        self.dayOrder = []
        for i in range(0, 7):
            try:            
                self.dayOrder.append((calendar.day_abbr[i][:1]).upper())
            except UnicodeError:
                # for Greek, Russian, etc.
                self.dayOrder.append((calendar.day_abbr[i][:2]).upper())
        self.mycal = calendar.Calendar()
        # layers
        self.layerCal = 'Calendar'
        # character styles
        self.cStylMonth = "char_style_Month"
        self.cStylDayNames = "char_style_DayNames"
        self.cStylWeekNo = "char_style_WeekNo"
        self.cStylHolidays = "char_style_Holidays"
        self.cStylDate = "char_style_Date"
        self.cStylDateWeekend = "char_style_DateWeekend"
        self.cStylLegend = "char_style_Legend"
        # paragraph styles
        self.pStyleMonth = "par_style_Month"
        self.pStyleDayNames = "par_style_DayNames"
        self.pStyleWeekNo = "par_style_WeekNo"
        self.pStyleHolidays = "par_style_Holidays"
        self.pStyleDate = "par_style_Date"
        self.pStyleWeekend = "par_style_Weekend"
        self.pStyleHoliday = "par_style_Holiday"
        self.pStyleLegend = "par_style_Legend"
        # line styles
        self.gridLineStyle = "grid_Line_Style"
        self.gridMonthHeadingStyle = "grid_Month_Heading_Style"
        # other settings
        calendar.setfirstweekday(0)
        progressTotal(12)

    def createCalendar(self):
        # UNIT_MILLIMETERS
        newDocument(PAPER_A1, (15, 15, 100, 15), LANDSCAPE, 1, UNIT_MILLIMETERS, NOFACINGPAGES, FIRSTPAGERIGHT, 1)
        # setUnit(UNIT_PT)
        setUnit(UNIT_MILLIMETERS)
        zoomDocument(16)
        scrollDocument(0,0)
        self.setupDocVariables()
        setActiveLayer(self.layerCal)
        run = 0
        for month in self.months:
             run += 1
             progressSet(run)
             cal = self.mycal.monthdatescalendar(self.year, month)
             self.createMonthCalendar(month, cal)
        setUnit(UNIT_MILLIMETERS)
        return None

    def setupDocVariables(self):
        """ Compute base metrics here. Page layout is bordered by margins
            and empty image frame(s). """
        # page dimensions
        page = getPageSize()
        self.pageX = page[0]
        self.pageY = page[1]
        marg = getPageMargins()
        self.marginT = marg[0]
        self.marginL = marg[1]
        self.marginR = marg[2]
        self.marginB = marg[3]
        self.width = self.pageX - self.marginL - self.marginR
        self.height = self.pageY - self.marginT - self.marginB
        self.rows = 32
        self.rowSize = (self.height) / self.rows
        print("self.rowSize={}".format(self.rowSize))
        self.mthcols = 2
        self.cols = 12
        self.colSize = (self.width) / self.cols
        print("self.colSize={}".format(self.colSize))
        # with ascender and descender characters
        # default calendar colors
        defineColorCMYK("Black", 0, 0, 0, 255)
        defineColorCMYK("White", 0, 0, 0, 0)
        defineColorCMYK("fillMonthHeading", 0, 0, 0, 0)  # default is White
        defineColorCMYK("txtMonthHeading", 0, 0, 0, 255)  # default is Black
        defineColorCMYK("fillDayNames", 0, 0, 0, 200)  # default is Dark Grey
        defineColorCMYK("txtDayNames", 0, 0, 0, 0)  # default is White
        defineColorCMYK("fillWeekNo", 0, 0, 0, 200)  # default is Dark Grey
        defineColorCMYK("txtWeekNo", 0, 0, 0, 0)  # default is White
        defineColorCMYK("fillDate", 0, 0, 0, 0)  # default is White
        defineColorCMYK("txtDate", 0, 0, 0, 255)  # default is Black
        defineColorCMYK("fillWeekend", 0, 0, 0, 25)  # default is Light Grey
        defineColorCMYK("txtWeekend", 0, 0, 0, 200)  # default is Dark Grey
        defineColorCMYK("fillWeekend2", 0, 0, 0, 0)  # default is White
        defineColorCMYK("fillHoliday", 0, 0, 0, 25)  # default is Light Grey
        defineColorCMYK("txtHoliday", 0, 234, 246, 0)  # default is Red
        defineColorCMYK("fillSpecialDate", 0, 0, 0, 0)  # default is White
        defineColorCMYK("txtSpecialDate", 0, 0, 0, 128)  # default is Middle Grey
        defineColorCMYK("fillVacation", 0, 0, 0, 25)  # default is Light Grey
        defineColorCMYK("txtVacation", 0, 0, 0, 255)  # default is Black
        defineColorCMYK("gridColor", 0, 0, 0, 255)  # default is Dark Grey
        defineColorCMYK("gridMonthHeading", 0, 0, 0, 255)  # default is Dark Grey
        # styles
        scribus.createCharStyle(name=self.cStylMonth, font="Lato Bold", fontsize=70, fillcolor="txtMonthHeading")
        scribus.createCharStyle(name=self.cStylDayNames, font="Lato Regular", fontsize=30, fillcolor="txtDayNames")
        scribus.createCharStyle(name=self.cStylWeekNo, font="Lato Regular", fontsize=30, fillcolor="txtWeekNo")
        scribus.createCharStyle(name=self.cStylHolidays, font="Lato Regular", fontsize=70, fillcolor="txtHoliday")
        scribus.createCharStyle(name=self.cStylDate, font="Lato Regular", fontsize=70, fillcolor="txtDate")
        scribus.createCharStyle(name=self.cStylDateWeekend, font="Lato Bold", fontsize=70, fillcolor="txtDate")
        scribus.createParagraphStyle(name=self.pStyleMonth, linespacingmode=2, alignment=ALIGN_CENTERED, charstyle=self.cStylMonth)
        scribus.createParagraphStyle(name=self.pStyleDayNames, linespacingmode=2, alignment=ALIGN_CENTERED, charstyle=self.cStylDayNames)
        scribus.createParagraphStyle(name=self.pStyleWeekNo,  linespacingmode=2, alignment=ALIGN_CENTERED, charstyle=self.cStylWeekNo)
        scribus.createParagraphStyle(name=self.pStyleWeekNo,  linespacingmode=2, alignment=ALIGN_CENTERED, charstyle=self.cStylWeekNo)
        scribus.createParagraphStyle(name=self.pStyleHolidays, linespacingmode=2, alignment=ALIGN_LEFT, charstyle=self.cStylHolidays)
        scribus.createParagraphStyle(name=self.pStyleDate, linespacingmode=2, alignment=ALIGN_LEFT, charstyle=self.cStylDate)
        scribus.createParagraphStyle(name=self.pStyleWeekend, linespacingmode=2, alignment=ALIGN_LEFT, charstyle=self.cStylDateWeekend)
        scribus.createParagraphStyle(name=self.pStyleHoliday, linespacingmode=2, alignment=ALIGN_LEFT, charstyle=self.cStylDateWeekend)
        scribus.createCustomLineStyle(self.gridLineStyle, [
            {
                'Color': "gridColor",
                'Width': 0
            }
        ])
        scribus.createCustomLineStyle(self.gridMonthHeadingStyle, [
            {
                'Color': "gridMonthHeading",
                'Width': 0
            }
        ])
        # layers
        createLayer(self.layerCal)

    def createMonthCalendar(self, month, cal):
        """ Draw one month calendar """
        self.createMonthHeader(month)
        row = 0
        for week in cal:
            for day in week:
                if day.month == month:
                    row += 1
                    cel = createText(
                        self.marginL + 10 + (month-1) * self.colSize,
                        self.marginT + row * self.rowSize,
                        self.colSize - 20,
                        self.rowSize
                    )
                    setFillColor("fillDate", cel)
                    # setCustomLineStyle(self.gridLineStyle, cel)
                    setText(str(day.day), cel)

                    # m variable - margin for keep distance between day number, and label
                    m = 0
                    if day.day > 9:
                        m = 12
                    cel_label = createText(
                        self.marginL + (28 + m) + (month - 1) * self.colSize,
                        self.marginT + row * self.rowSize,
                        (60 - m),
                        34
                    )
                    setText(str(day.strftime('%a')), cel_label)
                    deselectAll()
                    selectObject(cel_label)
                    setTextVerticalAlignment(ALIGNV_BOTTOM, cel_label)
                    setParagraphStyle(self.pStyleDayNames, cel_label)

                    deselectAll()
                    selectObject(cel)
                    if day.weekday() > 4:
                        setParagraphStyle(self.pStyleWeekend, cel)
                        setTextColor("txtWeekend", cel)
                        setTextColor("txtWeekend", cel_label)
                        setFillColor("fillWeekend", cel)
                    else:
                        setParagraphStyle(self.pStyleDate, cel)
                    setTextVerticalAlignment(ALIGNV_CENTERED, cel)
                    for x in range(len(self.holidaysList)):
                        if (self.holidaysList[x][0] == (day.year) and
                                self.holidaysList[x][1] == str(day.month) and
                                self.holidaysList[x][2] == str(day.day)):
                            if self.holidaysList[x][4] == "":
                                if getFillColor(cel) != "fillWeekend":
                                    setTextColor("txtVacation", cel)
                                    setTextColor("txtVacation", cel_label)
                                    setFillColor("fillVacation", cel)
                            elif self.holidaysList[x][4] == '0':
                                setTextColor("txtSpecialDate", cel)
                                if getFillColor(cel) == "fillWeekend":
                                    setFillColor("fillWeekend", cel)
                                elif getFillColor(cel) == "fillVacation":
                                    setFillColor("fillVacation", cel)
                                else:
                                    setFillColor("fillSpecialDate", cel)
                            else:
                                setTextColor("txtHoliday", cel)
                                setTextColor("txtHoliday", cel_label)
                                setFillColor("fillHoliday", cel)
        return

    def createMonthHeader(self, month):
        """ Draw month calendars header """
        monthName = months_names[month-1]
        cel = createText(
            self.marginL + 10 + (month-1) * self.colSize,
            self.marginT,
            self.colSize - 20,
            self.rowSize
        )
        setText(monthName.upper(), cel)
        setFillColor("fillMonthHeading", cel)
        # setCustomLineStyle(self.gridMonthHeadingStyle, cel)
        deselectAll()
        selectObject(cel)
        setParagraphStyle(self.pStyleMonth, cel)
        setTextVerticalAlignment(ALIGNV_TOP, cel)


class calcHolidays:
    """ Import local holidays from '*holidays.txt'-file and convert the variable
    holidays into dates for the given year."""

    def __init__(self, year):
        self.year = year

    def calcEaster(self):
        """ Calculate Easter date for the calendar Year using Butcher's Algorithm. 
        Works for any date in the Gregorian calendar (1583 and onward)."""
        a = self.year % 19
        b = self.year // 100
        c = self.year % 100
        d = (19 * a + b - b // 4 - ((b - (b + 8) // 25 + 1) // 3) + 15) % 30
        e = (32 + 2 * (b % 4) + 2 * (c // 4) - d - (c % 4)) % 7
        f = d + e - 7 * ((a + 11 * d + 22 * e) // 451) + 114	
        easter = datetime.date(int(self.year), int(f//31), int(f % 31 + 1))
        return easter

    def calcEasterO(self):
        """ Calculate Easter date for the calendar Year using Meeus Julian Algorithm. 
        Works for any date in the Gregorian calendar between 1900 and 2099."""
        d = (self.year % 19 * 19 + 15) % 30
        e = (self.year % 4 * 2 + self.year % 7 * 4 - d + 34) % 7 + d + 127
        m = e / 31
        a = e % 31 + 1 + (m>4)
        if a>30:a,m=1,5
        easter = datetime.date(int(self.year), int(m), int(a))
        return easter

    def calcVarHoliday(self, base, delta):
        """ Calculate variable Christian holidays dates for the calendar Year. 
        'base' is Easter and 'delta' the days from Easter."""
        holiday = base + timedelta(days=int(delta))
        return holiday

    def calcNthWeekdayOfMonth(self, n, weekday, month, year):
        """ Returns (month, day) tuple that represents nth weekday of month in year.
        If n==0, returns last weekday of month. Weekdays: Monday=0."""
        if not 0 <= n <= 5:
            raise IndexError("Nth day of month must be 0-5. Received: {}".format(n))
        if not 0 <= weekday <= 6:
            raise IndexError("Weekday must be 0-6")
        firstday, daysinmonth = calendar.monthrange(year, month)
        # Get first WEEKDAY of month
        first_weekday_of_kind = 1 + (weekday - firstday) % 7
        if n == 0:
        # find last weekday of kind, which is 5 if these conditions are met, else 4
            if first_weekday_of_kind in [1, 2, 3] and first_weekday_of_kind + 28 <= daysinmonth:
                n = 5
            else:
                n = 4
        day = first_weekday_of_kind + ((n - 1) * 7)
        if day > daysinmonth:
            raise IndexError("No {}th day of month {}".format(n, month))
        return (year, month, day)

    def importHolidays(self):
        """ Import local holidays from '*holidays.txt'-file."""
        # holidaysFile = filedialog.askopenfilename(title="Open the 'holidays.txt'-file or cancel")
        print(__file__)
        from os import path
        file_path = path.dirname(__file__)
        holidaysFile = path.join(file_path, "PL_holidays.txt")
        print(holidaysFile)
        holidaysList=list()
        try:
            csvfile = open(holidaysFile, mode="rt",  encoding="utf8")
        except Exception as e:
            print(e)
            print("Holidays wil NOT be shown.")
            messageBox("Warning:", "Holidays wil NOT be shown.", ICON_CRITICAL)
            return holidaysList # returns an empty holidays list
        csvReader = csv.reader(csvfile, delimiter=",")
        for row in csvReader:
            try:
                if row[0] == "fixed":
                    holidaysList.append((self.year, row[1], row[2], row[4], row[5]))
                    holidaysList.append((self.year + 1, row[1], row[2], row[4], row[5]))
                elif row[0] == "nWDOM":  # nth WeekDay Of Month
                    dt=self.calcNthWeekdayOfMonth(int(row[3]), int(row[2]), int(row[1]), int(self.year))
                    holidaysList.append((self.year, str(dt[1]), str(dt[2]), row[4], row[5]))
                    dt=self.calcNthWeekdayOfMonth(int(row[3]), int(row[2]), int(row[1]), int(self.year + 1))
                    holidaysList.append((self.year + 1, str(dt[1]), str(dt[2]), row[4], row[5]))
                elif row[0] == "variable":
                    if row[1] == "easter" :
                        base=self.calcEaster()
                        dt=self.calcVarHoliday(base, int(row[2]))
                        holidaysList.append(((dt.year), str(dt.month), str(dt.day), row[4], row[5]))
                        self.year =self. year + 1
                        base=self.calcEaster()
                        dt=self.calcVarHoliday(base, int(row[2]))
                        holidaysList.append(((dt.year), str(dt.month), str(dt.day), row[4], row[5]))
                        self.year = self.year - 1
                    elif row[1] == "easterO" :
                        base=self.calcEasterO()
                        dt=self.calcVarHoliday(base, int(row[2]))
                        holidaysList.append(((dt.year), str(dt.month), str(dt.day), row[4], row[5]))
                        self.year =self. year + 1
                        base=self.calcEaster()
                        dt=self.calcVarHoliday(base, int(row[2]))
                        holidaysList.append(((dt.year), str(dt.month), str(dt.day), row[4], row[5]))
                        self.year = self.year - 1
                else:
                    pass #do nothing
            except:
                print("Not a valid Holidays file.\nHolidays wil NOT be shown.")
                messageBox("Warning:",
                    "Not a valid Holidays file.\nHolidays wil NOT be shown.",
                    ICON_CRITICAL)
                break
        csvfile.close()
        return holidaysList


class TkCalendar(Frame):
    """ GUI interface for Scribus calendar wizard with tkinter"""

    def __init__(self, master=None):
        """ Setup the dialog """
        Frame.__init__(self, master)
        self.grid()
        self.master.resizable(0, 0)
        self.master.title('Scribus Year Calendar')

        #define widgets
        self.statusVar = StringVar()
        self.statusLabel = Label(self, fg="red", textvariable=self.statusVar)
        self.statusVar.set('Select Options and Values')

        self.langLabel = Label(self, text='Select language:')
        self.langFrame = Frame(self)
        self.langFrame.grid()
        self.langScrollbar = Scrollbar(self.langFrame, orient=VERTICAL)
        self.langScrollbar.grid(row=0, column=1, sticky=N+S)
        self.langListbox = Listbox(self.langFrame, selectmode=SINGLE, height=12,
            yscrollcommand=self.langScrollbar.set)
        self.langListbox.grid(row=0, column=0, sticky=N+S+E+W)
        self.langScrollbar.config(command=self.langListbox.yview)
        for i in range(len(localization)):
            self.langListbox.insert(END, localization[i][0])
        self.langButton = Button(self, text='Change language', command=self.languageChange)
        # choose font
        self.fontLabel = Label(self, text='Change font:')
        self.fontFrame = Frame(self)
        self.fontScrollbar = Scrollbar(self.fontFrame, orient=VERTICAL)
        self.fontListbox = Listbox(self.fontFrame, selectmode=SINGLE, height=12, yscrollcommand=self.fontScrollbar.set)
        self.fontScrollbar.config(command=self.fontListbox.yview)
        fonts = getFontNames()
        fonts.sort()
        for i in fonts:
            self.fontListbox.insert(END, i)
        self.font = 'Lato Regular'
        self.font_bold = 'Lato Bold'
        self.fontButton = Button(self, text='Apply selected font', command=self.fontApply)

        # year
        self.yearLabel = Label(self, text='Year:')
        self.yearVar = StringVar()
        self.yearEntry = Entry(self, textvariable=self.yearVar, width=4)

        # start of week
        self.weekStartsLabel = Label(self, text='Week begins with:')
        self.weekVar = IntVar()
        self.weekMondayRadio = Radiobutton(self, text='Mon', variable=self.weekVar, value=calendar.MONDAY)
        self.weekSundayRadio = Radiobutton(self, text='Sun', variable=self.weekVar, value=calendar.SUNDAY)

        # include week number
        self.weekNrLabel = Label(self, text='Show week numbers:')
        self.weekNrVar = IntVar()
        self.weekNrCheck = Checkbutton(self, variable=self.weekNrVar)
        self.weekNrHdLabel = Label(self, text='Week numbers heading:')
        self.weekNrHdVar = StringVar()
        self.weekNrHdEntry = Entry(self, textvariable=self.weekNrHdVar, width=6)
        
        # offsetX, offsetY and inner margins
        self.marginTopLabel = Label(self, text='Margin top (mm):')
        self.marginTop = DoubleVar()
        self.marginTopEntry = Entry(self, textvariable=self.marginTop, width=7)
        
        self.marginRightLabel = Label(self, text='Margin right (mm):')
        self.marginRight = DoubleVar()
        self.marginRightEntry = Entry(self, textvariable=self.marginRight, width=7)
        self.marginBottomLabel = Label(self, text='Margin bottom (mm):')
        self.marginBottom = DoubleVar()
        self.marginBottomEntry = Entry(self, textvariable=self.marginBottom, width=7)
        self.marginLeftLabel = Label(self, text='Margin left (mm):')
        self.marginLeft = DoubleVar()
        self.marginLeftEntry = Entry(self, textvariable=self.marginLeft, width=7)

        # holidays
        self.holidaysLabel = Label(self, text='Show holidays:')
        self.holidaysVar = IntVar()
        self.holidaysCheck = Checkbutton(self, variable=self.holidaysVar)

        # legend
        self.legendLabel = Label(self, text='Show holiday texts:')
        self.legendVar = IntVar()
        self.legendCheck = Checkbutton(self, variable=self.legendVar)

        # closing/running
        self.okButton = Button(self, text="OK", width=6, command=self.okButton_pressed)
        self.cancelButton = Button(self, text="Cancel", command=self.quit)

        # setup values
        self.yearVar.set(str(datetime.date(1, 1, 1).today().year+1))  # +1 for next year
        self.weekMondayRadio.select()
        self.weekNrHdVar.set("")
        self.marginTop.set("100")
        self.marginRight.set("15")
        self.marginBottom.set("15")
        self.marginLeft.set("15")
        self.holidaysCheck.select()
        self.legendCheck.select()

        # make layout
        self.columnconfigure(0, pad=6)
        currRow = 0
        self.statusLabel.grid(column=0, row=currRow, columnspan=4)
        currRow += 1
        self.langLabel.grid(column=0, row=currRow, sticky=W)
        self.fontLabel.grid(column=1, row=currRow, sticky=W) 
        currRow += 1
        self.langFrame.grid(column=0, row=currRow, rowspan=6, sticky=N)
        self.fontFrame.grid(column=1, row=currRow, sticky=N)
        self.fontScrollbar.grid(column=1, row=currRow, sticky=N+S+E)
        self.fontListbox.grid(column=0, row=currRow, sticky=N+S+W)
        currRow += 2
        self.langButton.grid(column=0, row=currRow)
        self.fontButton.grid(column=1, row=currRow)
        currRow += 1
        self.yearLabel.grid(column=0, row=currRow, sticky=S+E)
        self.yearEntry.grid(column=1, row=currRow, sticky=S+W)
        currRow += 1
        self.marginRightLabel.grid(column=2, row=currRow, sticky=S+E)
        self.marginRightEntry.grid(column=3, row=currRow, sticky=S+W)
        currRow += 1
        self.weekStartsLabel.grid(column=0, row=currRow, sticky=S+E)
        self.weekMondayRadio.grid(column=1, row=currRow, sticky=S+W)
        self.marginLeftLabel.grid(column=2, row=currRow, sticky=N+E)
        self.marginLeftEntry.grid(column=3, row=currRow, sticky=W)
        currRow += 1
        self.weekSundayRadio.grid(column=1, row=currRow, sticky=N+W)
        self.marginTopLabel.grid(column=2, row=currRow, sticky=S+E)
        self.marginTopEntry.grid(column=3, row=currRow, sticky=S+W)
        currRow += 1
        self.weekNrLabel.grid(column=0, row=currRow, sticky=N+E)
        self.weekNrCheck.grid(column=1, row=currRow, sticky=N+W)
        self.marginBottomLabel.grid(column=2, row=currRow, sticky=N+E)
        self.marginBottomEntry.grid(column=3, row=currRow, sticky=W)
        currRow += 1
        self.weekNrHdLabel.grid(column=0, row=currRow, sticky=N+E)
        self.weekNrHdEntry.grid(column=1, row=currRow, sticky=N+W)
        currRow += 1
        self.holidaysLabel.grid(column=0, row=currRow, sticky=N+E)
        self.holidaysCheck.grid(column=1, row=currRow, sticky=N+W)
        self.legendLabel.grid(column=2, row=currRow, sticky=N+E)
        self.legendCheck.grid(column=3, row=currRow, sticky=N+W)
        currRow += 1
        self.rowconfigure(currRow, pad=6)
        self.okButton.grid(column=1, row=currRow, sticky=E)
        self.cancelButton.grid(column=2, row=currRow, sticky=W)

        # fill the months values
        self.realLangChange()

    def languageChange(self):
        """ Called by Change button. Get language list value and
            call real re-filling. """
        ix = self.langListbox.curselection()
        if len(ix) == 0:
            self.statusVar.set('languageChange')
            return
        langX = self.langListbox.get(ix[0])
        self.lang = langX
        if os == "Windows":
            x = langX
        else: # Linux
            iy = [[x[0] for x in localization].index(self.lang)]
            x = (localization[iy[0]][2])
        try:
            locale.setlocale(locale.LC_CTYPE, x)
            locale.setlocale(locale.LC_TIME, x)
        except locale.Error:
            print("Language " + x + " is not installed on your operating system.")
            self.statusVar.set("Language '" + x + "' is not installed on your operating system")
            returnroot
        self.realLangChange(langX)

    def realLangChange(self, langX='Polish'):
        """ Real widget setup. It takes values from localization list.
        [0] = months, [1] Days """
        self.lang = langX
        if os == "Windows":
            ix = [[x[0] for x in localization].index(self.lang)]
            self.calUniCode = (localization[ix[0]][1]) # get unicode page for the selected language
        else: # Linux
            self.calUniCode = "UTF-8"

    def fontApply(self):
        """ Font selection. Called by "Apply selected font" button click. """
        ix = self.fontListbox.curselection()
        if len(ix) == 0:
            self.statusVar.set('Please select a font.')
            return
        self.font = self.fontListbox.get(ix[0])

    def okButton_pressed(self):
        """ User variables testing and preparing """
        # start year
        try:
            year = self.yearVar.get().strip()
            if len(year) != 4:
                raise ValueError
            year = int(year, 10)
        except ValueError:
            self.statusVar.set('Year must be in the "YYYY" format e.g. 2020.')
            return
        fonts = getFontNames()
        if self.font not in fonts:
            self.statusVar.set('Please select a font.')
            return
        if self.holidaysVar.get() == 0:
            holidaysList = list()
        else:
            holidaysList = []
            hol = calcHolidays(year)
            holidaysList = hol.importHolidays()
            self.statusVar.set('holidaysList')
            holidaysList.sort(key=lambda i: int(i[2]))  # sort on day
            holidaysList.sort(key=lambda i: int(i[1]))  # sort on month
            holidaysList.sort(key=lambda i: i[0])  # sort on year
        self.statusVar.set('defining ScYearCalendar')
        cal = ScYearCalendar(
            year=year,
            lang=self.lang,
            holidaysList=holidaysList
        )
        self.statusVar.set('withdraw')
        self.master.withdraw()
        self.statusVar.set('ScYearCalendar.createCalendar()')
        err = cal.createCalendar()
        if err != None:
            self.master.deiconify()
            self.statusVar.set(err)
        else:
            self.quit()

    def quit(self):
        self.master.destroy()


def main():
    """ Application/Dialog loop with Scribus sauce around """
    try:
        statusMessage('Running script...')
        progressReset()
        original_locale1 = locale.getlocale(locale.LC_CTYPE)
        original_locale2 = locale.getlocale(locale.LC_TIME)
        root = Tk()
        app = TkCalendar(root)
        root.mainloop()
        locale.setlocale(locale.LC_CTYPE, original_locale1)
        locale.setlocale(locale.LC_TIME, original_locale2)
    finally:
        if haveDoc() > 0:
            redrawAll()
        statusMessage('Done.')
        progressReset()


if __name__ == '__main__':
    main()
