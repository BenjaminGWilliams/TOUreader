import csv
import sys
import calendar
import datetime
from dateutil.relativedelta import relativedelta
import os

from MSOffice import Excel
from MSOffice.Excel.Worksheets.Worksheet import Sheet

from operator import itemgetter
# from plotly.api.v2.grids import row

'''
Created on 8/09/2017

@author: bgw


'''
ICPDIC = {}

HHR_ICP_DIR = r"X:\NRD"
PERCENT = 99


STARTDATE = datetime.datetime(2017, 10, 1)
ENDDATE = datetime.datetime(2017, 10, 1)


DESTFILEPATH = "Test.xlsx"
ERRORMSGPATH = "Log.txt"

NUMLARGESTVALUES = 10


class Data():

    def __init__(self, kwh, kvarh, **kwargs):
        self.kwh = kwh
        self.kvarh = kvarh
        self.pflowimport = kwargs.get("pflowimport", False)
        self.pflowexport = kwargs.get("pflowexport", False)

    def __repr__(self):
        return "Data(kwh: %.3f, kvarh: %.2f)" % (self.kwh, self.kvarh)
        # return "%.3f" % self.S

    def __add__(self, other):

        return Data(self.kwh + other.kwh, self.kvarh + other.kvarh,
                    pflowimport=(self.pflowimport or other.pflowimport),
                    pflowexport=(self.pflowexport or other.pflowexport))


def dstcalc(STARTDATE, ENDDATE):

    startyear = STARTDATE.year
    endyear = ENDDATE.year
    dstlist = []

    for year in range(startyear, endyear + 1):

        last_sunday = datetime.datetime(year, 9, max(week[-1] for
                                                   week in calendar.monthcalendar(year, 9)))  # last Sunday of September

        first_sunday = datetime.datetime(year, 4, calendar.monthcalendar(year, 4)[0][-1])  # first Sunday of April

        dstlist.append((last_sunday, first_sunday))
    return dstlist


def writetofile(errmsg):

    if os.path.exists(ERRORMSGPATH):
        append_write = 'a'
    else:
        append_write = 'w'
    with open(ERRORMSGPATH, append_write) as text_file:
        text_file.write(errmsg)
        text_file.write("\n")

def addtoicpdic(icpnum, meterno, date, data, dstchanged, row, filepath):
    meters = ICPDIC.get(icpnum)
    if meters is None:
    # ICPDIC[icpnum] = {date:Data(S,PF)}
        ICPDIC[icpnum] = {meterno:{date:data}}
    else:


        meter = meters.get(meterno)

        if meter is None:
            ICPDIC[icpnum][meterno] = {date:data}
        else:
            metervalue = meter.get(date)
            if metervalue is None:
                ICPDIC[icpnum][meterno][date] = data

            elif (metervalue.pflowimport == data.pflowimport) and not dstchanged and meterno != 'Total':
                errmsg = "ERROR! Row %s of file: %s already existed in dictionary" % (row, filepath)
                print errmsg
                writetofile(errmsg)
                return

            else:
                ICPDIC[icpnum][meterno][date] += data



def Read_HHR_File(filepath, dstend):
    """Read a particular HHR file (1 month), and create the nesseccary data structures, 
    or append to an existing structure."""
    filestring = "%s..." % (filepath)
    sys.stdout.write(filestring)
    sys.stdout.flush()

    # Open a retailer meter file (contains 1/2 hour readings for 1 month)
    with open(filepath, 'rb') as TOU_File:


        TOU_Reader = csv.reader(TOU_File)
        # Skip the header column
        try:
            TOU_Reader.next()
        except:
            # Seems to happen for particular cases where where the file is named "xx.xls.csv"
            errmsg = "ERROR! Could not add file: %s" % filepath
            print errmsg
            writetofile(errmsg)
            # TOU_Reader = csv.reader(TOU_File, dialect=csv.excel_tab)
            # TOU_Reader.next()
            TOU_Reader = []

        # Every file starts on the 1st day of the month 30 minutes past midnight.
        # This file has 3 columns of readings: \\Pnlicp1\nrd\2015\201503\EIEP3_ICPHH\TRUS_E_OTPO_ICPHH_201503_20150401_1537.csv
        # (but the first two columns of values are still kwh and kvarh). The third column is kVA.
        for row in TOU_Reader:
            # Unpack a row of data
            if len(row) == 1:
                try:
                    row = row[0].split("\t")
                except:
                    row = []
            if len(row) == 13:
                (_, icpnum, meterno, _, date, hr, kwh, kvarh, _, pflow, _, _, _) = row
            elif len(row) == 12:
                (_, icpnum, meterno, _, date, hr, kwh, kvarh, _, pflow, _, _) = row
            elif len(row) == 11:
                (_, icpnum, meterno, _, date, hr, kwh, kvarh, _, pflow, _) = row
            elif len(row) == 10:
                (_, icpnum, meterno, _, date, hr, kwh, kvarh, _, pflow) = row
            else:
                errmsg = ("ERROR! %s has a line length of %s which is unknown. "
                          "The row follows the format: %s" % (filepath, len(row), row))
                print errmsg
                writetofile(errmsg)


                break
            if pflow.lower() not in ("x", "i"):

                errmsg = ("WARNING! %s of file %s does not have X or I as power flow, assumed X"
                          % (row, filepath))
                print errmsg
                writetofile(errmsg)





        # Create a unique datetime stamp
            date = date.strip(" ")  # Some dates have trailing whitespace
            try:
                date = datetime.datetime.strptime(date, "%d/%m/%Y")
            except ValueError:
                # Randomly some files are in this format
                try:
                    date = datetime.datetime.strptime(date, "%m/%d/%Y")
                except:
                    try:
                        date = datetime.datetime.strptime(date, "%d/%m/%y")

                    except:
                    # E.g. PLEL_E_DUNE_ICPHH_201608_20160907_426206750879977.txt
                        errmsg = ("ERROR! %s of file %s does not have a valid "
                                  "date format: %s" % (row, filepath, date))
                        print errmsg
                        writetofile(errmsg)
                        break


            dstchanged = False
            if date in dstend and int(hr) > 48:
                hr = int(hr) - 2
                dstchanged = True
                # hr -= 2
            elif date not in dstend and int(hr) > 48 or int(hr) > 50:
                errmsg = "ERROR! Found %s half hours on line %s of file %s:" % (str(hr), row, filepath)
                print errmsg
                writetofile(errmsg)
                break
    #===============================================================================================
    #
    #         The first reading of the day is at 0030 and the last is at 0000 (which is actually
    #         the start of the next day)
    #         So the 48 meter readings in a given day will actually span two days by 1 sample point.
    #===============================================================================================

            if (int(hr) < 48):
                date = date.replace(hour=int(hr) / 2, minute=(int(hr) % 2) * 30)
            else:
                # The very last reading will be at midnight and python uses midnight as the
                # start of a new day
                date = date + datetime.timedelta(days=1)

            # Calculate S, and PF from the half hour meter information
            if len(kwh) == 0:
                kwh = 0
            if len(kvarh) == 0:
                kvarh = 0


            kwh = float(kwh)
            kvarh = float(kvarh)

            if pflow.lower() == "i":
                data = Data(kwh * -1, kvarh * -1, pflowimport=True)
            else:
                data = Data(kwh, kvarh, pflowexport=True)

            addtoicpdic(icpnum, meterno, date, data, dstchanged, row, filepath)
            # addtoicpdic(icpnum, 'Total', date, data, dstchanged, row, filepath)

    sys.stdout.write(" Finished.\n")


def CreateTOUFileDirs(basedir, starttime, endtime):

    starttime = datetime.date(starttime.year, starttime.month, 1)
    endtime = datetime.date(endtime.year, endtime.month + 1, 1)


    yearmo = datetime.date(starttime.year, starttime.month, 1)

    filepaths = []

    while yearmo < endtime:
        yearstr = "%s" % (yearmo.year)
        yearmostr = "%s%02d" % (yearmo.year, yearmo.month)

        fullfp = os.path.join(basedir, yearstr, yearmostr, "EIEP3_ICPHH")
        if os.path.isdir(fullfp):

            filepaths.append(fullfp)


        yearmo += relativedelta(months=1)

    return filepaths



if __name__ == '__main__':

    dstlist = dstcalc(STARTDATE, ENDDATE)
    dstend = [dstdate[1] for dstdate in dstlist]
    # headinglist = ["Date", "kVA", "PF"]
    headinglist = ["Date", "kWh", "kVarh"]
    numheadings = len(headinglist)
    sheetname = "Sheet1"



    DIR_LIST = CreateTOUFileDirs(HHR_ICP_DIR, STARTDATE, ENDDATE)



    for mydir in DIR_LIST:
        for filen in os.listdir(mydir):
            if filen.endswith(".txt"):
                fullfilepath = os.path.join(mydir, filen)
                Read_HHR_File(fullfilepath, dstend)




    xl = Excel.Launch.Excel(visible=True, runninginstance=False,
                            BookVisible=True, filename=DESTFILEPATH)
    wkbdest = Sheet(xl)
    DestSheet = xl.xlBook.Worksheets[sheetname]

    xcounter = 1
    for icpkey, icpvalue in ICPDIC.iteritems():
        wkbdest.setCell(sheetname, 1, xcounter, "ICP:")
        wkbdest.setCell(sheetname, 2, xcounter, icpkey)



        for meterkey, metervalue in icpvalue.iteritems():
            wkbdest.setCell(sheetname, 4, xcounter, "METER:")
            wkbdest.setCell(sheetname, 5, xcounter, meterkey)


            datalist = []

            for key, value in sorted(metervalue.iteritems()):

                S = ((2 * float(value.kwh)) ** 2 + (2 * float(value.kvarh)) ** 2) ** 0.5

        # Catch divide by 0 case
                if S != 0:
                    PF = (2 * float(value.kwh)) / S
                else:
                    PF = 1.0

                datalist.append([key, value.kwh, value.kvarh])
                # datalist.append([key, S, PF])

            sortedlist = sorted(datalist, key=itemgetter(1))
            largestvaluelist = sortedlist[-NUMLARGESTVALUES:]

            percentvalue = sortedlist[int(len(metervalue.keys()) * PERCENT / 100.00)]  # % rounded down



            wkbdest.setCell(sheetname, 7, xcounter, "%s%% largest value:" % PERCENT)

            wkbdest.setCell(sheetname, 10, xcounter, "%s largest values:" % NUMLARGESTVALUES)
            for i in range(NUMLARGESTVALUES):
                for ii in range(len(headinglist)):
                    wkbdest.setCell(sheetname, 12 + i, xcounter + ii, largestvaluelist[i][ii])


            for i in range(len(headinglist)):

                wkbdest.setCell(sheetname, 8, xcounter + i, percentvalue[i])
                wkbdest.setCell(sheetname, 11, xcounter + i, headinglist[i])

                wkbdest.setCell(sheetname, 13 + NUMLARGESTVALUES, xcounter + i, headinglist[i])

            wkbdest.setRange(sheetname, 14 + NUMLARGESTVALUES, xcounter, datalist)

            xcounter += numheadings
        xcounter += 1



    print ("Complete")
