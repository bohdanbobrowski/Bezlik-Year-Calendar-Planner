import calendar
import csv
import datetime
import locale
import platform
import sys
from datetime import timedelta

try:
    from scribus import *
except ImportError:
    print("This Python script is written for the Scribus scripting interface.")
    print("It can only be run from within Scribus.")
    sys.exit(1)

python_version = platform.python_version()
if python_version[0:1] != "3":
    print("This script runs only with Python 3 (Scribus 1.5.x).")
    messageBox("Script failed",
        "This script runs only with Python 3 (Scribus 1.5.x).",
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
    ['Polish', 'CP1250', 'pl_PL.UTF8', 'PL'],
    ['English', 'CP1252', 'en_US.UTF8', 'US'],
    ['French', 'CP1252', 'fr_FR.UTF8', 'FR'],
    ['German', 'CP1252', 'de_DE.UTF8', 'DE'],
    ['Russian', 'CP1251', 'ru_RU.UTF8', 'RU'],
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

class BezlikYearCalendar:

    def __init__(
            self,
            year,
            margin_top=15,
            margin_right=15,
            margin_bottom=15,
            margin_left=15,
            font_regular="Lato Regular",
            font_header="Lato Bold",
            font_holiday="Lato Light Italic",
            lang='English',
            holidays_list=list()
    ):
        self.year = year
        self.marginTop = margin_top
        self.marginRight = margin_right
        self.marginBottom = margin_bottom
        self.marginLeft = margin_left

        self.font_regular = font_regular
        self.font_header = font_header
        self.font_holiday = font_holiday

        self.months = [month for month in range(1, 13)]
        self.nrVmonths = 12
        self.holidaysList = holidays_list
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
        self.cStyleMonth = "char_style_Month"
        self.cStyleDayNames = "char_style_DayNames"
        self.cStyleHolidayLabel = "char_style_WeekNo"
        self.cStyleHolidays = "char_style_Holidays"
        self.cStyleDate = "char_style_Date"

        # paragraph styles
        self.pStyleMonth = "par_style_Month"
        self.pStyleDayNames = "par_style_DayNames"
        self.pStyleHolidayLabel = "par_style_WeekNo"
        self.pStyleHolidays = "par_style_Holidays"
        self.pStyleDate = "par_style_Date"

        # other settings
        calendar.setfirstweekday(0)
        progressTotal(12)

    def createCalendar(self):
        newDocument(PAPER_A1, (self.marginLeft, self.marginRight, self.marginTop, self.marginBottom),
                    LANDSCAPE, 1, UNIT_MILLIMETERS, NOFACINGPAGES, FIRSTPAGERIGHT, 1)
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
        self.mthcols = 2
        self.cols = 12
        self.colSize = (self.width) / self.cols

        defineColorCMYK("Black", 0, 0, 0, 255)
        defineColorCMYK("White", 0, 0, 0, 0)
        defineColorCMYK("LightGrey", 0, 0, 0, 25)  
        defineColorCMYK("DarkGrey", 0, 0, 0, 210)
        defineColorCMYK("MediumGrey", 0, 0, 0, 128)
        defineColorCMYK("Red", 0, 234, 246, 0)

        # Month name
        scribus.createCharStyle(name=self.cStyleMonth, font=self.font_header, fontsize=70, fillcolor="Black")
        scribus.createParagraphStyle(name=self.pStyleMonth, linespacingmode=0, linespacing=30, alignment=ALIGN_CENTERED, charstyle=self.cStyleMonth)

        # Holiday date
        scribus.createCharStyle(name=self.cStyleHolidays, font=self.font_regular, fontsize=70, fillcolor="Red")
        scribus.createParagraphStyle(name=self.pStyleHolidays, linespacingmode=0, linespacing=30, alignment=ALIGN_LEFT, charstyle=self.cStyleHolidays)

        # Date (day number)
        scribus.createCharStyle(name=self.cStyleDate, font=self.font_regular, fontsize=70, fillcolor="Black")
        scribus.createParagraphStyle(name=self.pStyleDate, linespacingmode=2, alignment=ALIGN_LEFT, charstyle=self.cStyleDate)

        # Day name
        scribus.createCharStyle(name=self.cStyleDayNames, font=self.font_regular, fontsize=30, fillcolor="Black")
        scribus.createParagraphStyle(name=self.pStyleDayNames, linespacingmode=0, linespacing=30, alignment=ALIGN_LEFT, charstyle=self.cStyleDayNames)

        # Day name
        scribus.createCharStyle(name=self.cStyleHolidayLabel, font=self.font_holiday, fontsize=20, fillcolor="MediumGrey")
        scribus.createParagraphStyle(name=self.pStyleHolidayLabel,  linespacingmode=0, linespacing=30, alignment=ALIGN_LEFT, charstyle=self.cStyleHolidayLabel)

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
                    setFillColor("White", cel)
                    setTextColor("DarkGrey", cel)
                    setText(str(day.day), cel)
                    setParagraphStyle(self.pStyleDate, cel)
                    # m variable - margin for keep distance between day number, and label
                    m = 0
                    if day.day > 9:
                        m = 15
                    cel_label = createText(
                        self.marginL + (25 + m) + (month - 1) * self.colSize,
                        self.marginT + row * self.rowSize,
                        (60 - m),
                        self.rowSize
                    )
                    setText(str(day.strftime('%a')), cel_label)
                    deselectAll()
                    selectObject(cel_label)
                    setTextVerticalAlignment(ALIGNV_CENTERED, cel_label)
                    setParagraphStyle(self.pStyleDayNames, cel_label)
                    setTextColor("Black", cel_label)
                    deselectAll()
                    selectObject(cel)
                    if day.weekday() > 4:
                        setTextColor("Black", cel)
                        setFillColor("LightGrey", cel)
                        setTextColor("Black", cel_label)
                    setTextVerticalAlignment(ALIGNV_CENTERED, cel)
                    for x in range(len(self.holidaysList)):
                        if (
                            self.holidaysList[x][0] == (day.year) and
                            self.holidaysList[x][1] == str(day.month) and
                            self.holidaysList[x][2] == str(day.day)
                        ):
                            cel_holiday = createText(
                                self.marginL + 85 + (month - 1) * self.colSize,
                                self.marginT + row * self.rowSize,
                                self.colSize - 105,
                                self.rowSize
                            )
                            setTextColor("DarkGrey", cel_holiday)
                            setText(self.holidaysList[x][3], cel_holiday)
                            deselectAll()
                            selectObject(cel_label)
                            setTextVerticalAlignment(ALIGNV_CENTERED, cel_holiday)
                            setParagraphStyle(self.pStyleHolidayLabel, cel_holiday)
                            setTextColor("Black", cel_holiday)
                            if self.holidaysList[x][4] == "1":
                                setTextColor("Red", cel)
                                setTextColor("Red", cel_label)
                                setTextColor("Red", cel_holiday)
        return

    def createMonthHeader(self, month):
        """ Draw month calendars header """
        if self.lang == 'Polish':
            month_name = months_names[month-1]
        else:
            month_name = calendar.month_name[month]
        cel = createText(
            self.marginL + 10 + (month-1) * self.colSize,
            self.marginT,
            self.colSize - 20,
            self.rowSize
        )
        setText(month_name.upper(), cel)
        setFillColor("White", cel)
        deselectAll()
        selectObject(cel)
        setParagraphStyle(self.pStyleMonth, cel)
        setTextVerticalAlignment(ALIGNV_TOP, cel)


class calcHolidays:
    """ Import local holidays from '*holidays.txt'-file and convert the variable
    holidays into dates for the given year."""

    def __init__(self, year, locale_code):
        self.year = year
        self.locale_code = locale_code

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
        from os import path
        file_path = path.dirname(__file__)
        holidaysFile = path.join(file_path, "{}_holidays.csv".format(self.locale_code))
        holidaysList = list()
        try:
            csvfile = open(holidaysFile, mode="rt",  encoding="utf8")
        except Exception as e:
            error = "Holidays wil NOT be shown."
            print(error)
            messageBox("Warning:", error, ICON_CRITICAL)
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
                    if row[1] == "easter":
                        base = self.calcEaster()
                        dt = self.calcVarHoliday(base, int(row[2]))
                        holidaysList.append(((dt.year), str(dt.month), str(dt.day), row[4], row[5]))
                        self.year =self. year + 1
                        base = self.calcEaster()
                        dt = self.calcVarHoliday(base, int(row[2]))
                        holidaysList.append(((dt.year), str(dt.month), str(dt.day), row[4], row[5]))
                        self.year = self.year - 1
                    elif row[1] == "easterO":
                        base = self.calcEasterO()
                        dt = self.calcVarHoliday(base, int(row[2]))
                        holidaysList.append(((dt.year), str(dt.month), str(dt.day), row[4], row[5]))
                        self.year = self. year + 1
                        base = self.calcEaster()
                        dt = self.calcVarHoliday(base, int(row[2]))
                        holidaysList.append(((dt.year), str(dt.month), str(dt.day), row[4], row[5]))
                        self.year = self.year - 1
                else:
                    pass
            except:
                error = "Not a valid Holidays file.\nHolidays wil NOT be shown."
                print(error)
                messageBox("Warning:", error, ICON_CRITICAL)
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
        self.master.title('Bezlik Year Calendar')

        #default language variables
        self.lang = localization[0][0]
        self.locale_str = localization[0][2]
        self.locale_code = localization[0][3]

        #define widgets
        self.statusVar = StringVar()
        self.statusLabel = Label(self, fg="red", textvariable=self.statusVar)
        self.statusVar.set('Select Options and Values')

        self.langLabel = Label(self, text='Select language:')
        self.langFrame = Frame(self)
        self.langFrame.grid()
        self.langScrollbar = Scrollbar(self.langFrame, orient=VERTICAL)
        self.langScrollbar.grid(row=0, column=1, sticky=N+S)
        self.langListbox = Listbox(self.langFrame, selectmode=SINGLE, height=4, yscrollcommand=self.langScrollbar.set)
        self.langListbox.grid(row=0, column=0, sticky=N+S+E+W)
        self.langScrollbar.config(command=self.langListbox.yview)
        for i in range(len(localization)):
            self.langListbox.insert(END, localization[i][0])
        self.langButton = Button(self, text='Change language', command=self.languageChange)

        # year
        self.yearLabel = Label(self, text='Year:')
        self.yearVar = StringVar()
        self.yearEntry = Entry(self, textvariable=self.yearVar, width=4)

        # fonts
        self.fontHeaderLabel = Label(self, text='Date font:')
        self.fontHeaderVar = StringVar()
        self.fontHeaderEntry = Entry(self, textvariable=self.fontHeaderVar, width=15)
        
        self.fontLabel = Label(self, text='Day name font:')
        self.fontVar = StringVar()
        self.fontEntry = Entry(self, textvariable=self.fontVar, width=15)

        self.fontHolidayLabel = Label(self, text='Holiday font:')
        self.fontHolidayVar = StringVar()
        self.fontHolidayEntry = Entry(self, textvariable=self.fontHolidayVar, width=15)

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
        self.fontVar.set("Lato Regular")
        self.fontHeaderVar.set("Lato Bold")
        self.fontHolidayVar.set("Lato Light Italic")
        self.marginTop.set("300")
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
        self.langLabel.grid(column=0, row=currRow, columnspan=4)
        currRow += 1
        self.langFrame.grid(column=0, row=currRow, columnspan=4, rowspan=3, sticky=N)
        currRow += 4
        self.langButton.grid(column=0, row=currRow, columnspan=4)
        currRow += 1
        self.yearLabel.grid(column=0, row=currRow, sticky=S+E)
        self.yearEntry.grid(column=1, row=currRow, sticky=S+W)
        self.marginRightLabel.grid(column=2, row=currRow, sticky=S+E)
        self.marginRightEntry.grid(column=3, row=currRow, sticky=S+W)
        currRow += 1
        self.fontLabel.grid(column=0, row=currRow, sticky=N+E)
        self.fontEntry.grid(column=1, row=currRow, sticky=N+W)
        self.marginLeftLabel.grid(column=2, row=currRow, sticky=N+E)
        self.marginLeftEntry.grid(column=3, row=currRow, sticky=W)
        currRow += 1
        self.fontHeaderLabel.grid(column=0, row=currRow, sticky=N+E)
        self.fontHeaderEntry.grid(column=1, row=currRow, sticky=N+W)
        self.marginTopLabel.grid(column=2, row=currRow, sticky=S+E)
        self.marginTopEntry.grid(column=3, row=currRow, sticky=S+W)
        currRow += 1
        self.fontHolidayLabel.grid(column=0, row=currRow, sticky=N+E)
        self.fontHolidayEntry.grid(column=1, row=currRow, sticky=N+W)
        self.marginBottomLabel.grid(column=2, row=currRow, sticky=N+E)
        self.marginBottomEntry.grid(column=3, row=currRow, sticky=W)
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
        self.lang = self.langListbox.get(ix[0])
        iy = [[locale_str[0] for locale_str in localization].index(self.lang)]
        self.locale_str = (localization[iy[0]][2])
        self.locale_code = localization[iy[0]][3]
        try:
            locale.setlocale(locale.LC_CTYPE, self.locale_str)
            locale.setlocale(locale.LC_TIME, self.locale_str)
        except locale.Error:
            error = "Language {} is not installed on your operating system.".format(self.locale_str)
            print(error)
            self.statusVar.set(error)
            returnroot
        self.realLangChange(self.lang)

    def realLangChange(self, lang_x='Polish'):
        """ Real widget setup. It takes values from localization list.
        [0] = months, [1] Days """
        self.lang = lang_x
        if os == "Windows":
            ix = [[x[0] for x in localization].index(self.lang)]
            self.calUniCode = (localization[ix[0]][1])  # get unicode page for the selected language
        else:  # Linux
            self.calUniCode = "UTF-8"

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
        font_regular = self.fontVar.get().strip()
        font_header = self.fontHeaderVar.get().strip()
        font_holiday = self.fontHolidayVar.get().strip()
        if font_regular not in fonts or font_header not in fonts or font_holiday not in fonts:
            self.statusVar.set('Please select a font.')
            return
        if self.holidaysVar.get() == 0:
            holidaysList = list()
        else:
            hol = calcHolidays(year, self.locale_code)
            holidaysList = hol.importHolidays()
            self.statusVar.set('holidaysList')
            holidaysList.sort(key=lambda i: int(i[2]))  # sort on day
            holidaysList.sort(key=lambda i: int(i[1]))  # sort on month
            holidaysList.sort(key=lambda i: i[0])  # sort on year
        cal = BezlikYearCalendar(
            year=year,
            lang=self.lang,
            font_regular=font_regular,
            font_header=font_header,
            font_holiday=font_holiday,
            margin_top=int(self.marginTopEntry.get()),
            margin_right=int(self.marginRightEntry.get()),
            margin_bottom=int(self.marginBottomEntry.get()),
            margin_left=int(self.marginLeftEntry.get()),
            holidays_list=holidaysList
        )
        self.statusVar.set('withdraw')
        self.master.withdraw()
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
