'''

Created on 8/09/2017

@author: bgw

'''

import csv
import sys
import datetime
import os
from dateutil.relativedelta import relativedelta


from MSOffice import Excel
from MSOffice.Excel.Worksheets.Worksheet import Sheet


ICPDIC = {}

HHR_ICP_DIR = r"\\PNLSQLICP1\NRD"
PERCENT = 99


STARTDATE = datetime.datetime(2017, 10, 1)
ENDDATE = datetime.datetime(2017, 10, 1)


DESTFILENAME = "TOU.xlsx"
ERRORMSGPATH = "LogMaree.txt"
COMPLETEDPATH = "Completed.txt"

WORKINGPATH = os.path.dirname(os.path.abspath(__file__))

TOTALLIST = []



def writetofile(errmsg, filepath):
    """Writes provided errmsg to a inputted filepath. Appends or Writes if exists or not"""

    if os.path.exists(filepath):
        append_write = 'a'
    else:
        append_write = 'w'
    with open(filepath, append_write) as text_file:
        text_file.write(errmsg)
        text_file.write("\n")

def Read_HHR_File(filepath):
    """Read a particular HHR file (1 month), and create the nesseccary data structures,
    or append to an existing structure."""
    filestring = "Processing %s..." % (filepath)
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
            writetofile(errmsg, ERRORMSGPATH)
            # TOU_Reader = csv.reader(TOU_File, dialect=csv.excel_tab)
            # TOU_Reader.next()
            TOU_Reader = []

        # Every file starts on the 1st day of the month 30 minutes past midnight.
        # This file has 3 columns of readings: \\Pnlicp1\nrd\2015\201503\EIEP3_ICPHH\
        # TRUS_E_OTPO_ICPHH_201503_20150401_1537.csv
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
                writetofile(errmsg, ERRORMSGPATH)


                break
            if pflow.lower() not in ("x", "i"):

                errmsg = ("WARNING! %s of file %s does not have X or I as power flow, skipping"
                          % (row, filepath))
                print errmsg
                writetofile(errmsg, ERRORMSGPATH)





        # Create a unique datetime stamp
            date = date.strip(" ")  # Some dates have trailing whitespace

    #===============================================================================================
    #
    #         The first reading of the day is at 0030 and the last is at 0000 (which is actually
    #         the start of the next day)
    #         So the 48 meter readings in a given day will actually span two days by 1 sample point.
    #===============================================================================================


            # Calculate S, and PF from the half hour meter information

            if pflow.lower() == 'x':
                TOTALLIST.append(row)

            # addtoicpdic(icpnum, meterno, date, data, dstchanged, row, filepath)
            # addtoicpdic(icpnum, 'Total', date, data, dstchanged, row, filepath)

    sys.stdout.write(" Finished.\n")


def CreateTOUFileDirs(basedir, starttime, endtime):

    ''' Creates a list of all directories for given base directory from starttime to endtime'''

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

    starttimestr = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    writetofile('===STARTED SCRIPT AT {}==='.format(starttimestr), ERRORMSGPATH)
    writetofile('===STARTED SCRIPT AT {}==='.format(starttimestr), COMPLETEDPATH)

    completedlist = []
    if os.path.isfile(COMPLETEDPATH):
        f = open(COMPLETEDPATH, 'r')
        completedlist = f.readlines()


    destfilepath = os.path.join(WORKINGPATH, DESTFILENAME)

    xl = Excel.Launch.Excel(visible=True, runninginstance=False,
                            BookVisible=True, filename=destfilepath)
    wkbdest = Sheet(xl)
    sheetname = "Sheet1"



    DIR_LIST = CreateTOUFileDirs(HHR_ICP_DIR, STARTDATE, ENDDATE)

    maxrow = 0




    for mydir in DIR_LIST:
        for filen in os.listdir(mydir):
            fullfilepath = os.path.join(mydir, filen)

            if fullfilepath + '\n' in completedlist:
                errormsg = ("WARNING! %s was skipped as already added to spreadsheet"
                            % (fullfilepath))
                print errormsg
                writetofile(errormsg, ERRORMSGPATH)
            else:
                if filen.endswith(".txt"):
                    TOTALLIST = []


                    Read_HHR_File(fullfilepath)

                    if not TOTALLIST:  # no items to add to spreadsheet
                        errormsg = ("WARNING! %s had no 'X' values" % (fullfilepath))
                        print errormsg
                        writetofile(errormsg, ERRORMSGPATH)

                    else:

                        wkbdest.setRange(sheetname, maxrow + 1, 1, TOTALLIST)


                        maxrow = wkbdest.getMaxRow(sheetname, 1, 1)
                    writetofile(fullfilepath, COMPLETEDPATH)


    endtimestr = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    writetofile('===FINISHED SCRIPT AT {}==='.format(endtimestr), ERRORMSGPATH)
    writetofile('===FINISHED SCRIPT AT {}==='.format(endtimestr), COMPLETEDPATH)

    xl.save()
    print "Complete"
