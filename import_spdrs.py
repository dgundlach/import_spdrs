#!/usr/bin/python3

# This program populates watchlists for import into IBKR TWS Pro from the State Street Select
# SPDR ETFs and Industry Group ETFS.  It also populates similar lists from the S&P 500, S&P
# Mid Cap and S&P Small Cap Growth and Value ETFs.  It uses the spreadsheets that State Street
# provides generate the lists.  The files are written to a new subdirectory of the TWS
# configuration directory so that there is less effort to import them.  That can be overridden
# by changing the 'watchlists' variable to the value ''.
#
# Author: Dan Gundlach <cyclist2918@gmail.com>
# License: GPL2

import requests
import openpyxl
import os
from tempfile import mkstemp
import platform

tws_home = '/'                              # Windows drive where TWS resides.  On posix systems, this
                                            # is overridden to the user home directory.
watchlists = 'watchlists'                   # Directory where watchlists are created.
first_row = 6                               # First row of spreadsheet with stock data.

# These lists can be modified to add or remove any State Street ETF.

# Select SPDR ETFs.
sspdrs = ['xlb', 'xlc', 'xle', 'xlf', 'xli', 'xlk', 'xlp', 'xlre', 'xlu', 'xlv', 'xly']

# SPDR Industry Group ETFs.
igetfs = ['kbe', 'kre', 'kie', 'xar', 'xtn', 'xbi', 'xph', 'xhe', 'xhs', 'xop', 'xes', 'xme', 'xrt',
         'xhb', 'xsd', 'xsw', 'xntk', 'xitk', 'xtl', 'xweb']

# Simulated Industry Group ETFs are populated from these.
simetfs = ['mdy', 'mdyg', 'mdyv', 'sly', 'slyg', 'slyv', 'spyg', 'spyv']

# These values are not likely to be have to be altered.

worksheet = 'holdings'                      # Worksheet name within the spreadsheet.
no_sector = 'Unassigned'                    # When this value is found there are no more entries.
extension = '.xlsx'                         # Filename extension for spreadsheet.
baseurl = "https://www.ssga.com/library-content/products/fund-data/etfs/us/holdings-daily-us-en-"

# Populate the suffix list with the generic designations.

suffixes = {'Materials' : 'b',
            'Communication Services': 'c',
            'Energy' : 'e',
            'Financials' : 'f',
            'Industrials' : 'i',
            'Information Technology' : 'k',
            'Consumer Staples' : 'p',
            'Real Estate' : 're',
            'Utilities' : 'u',
            'Health Care' : 'v',
            'Consumer Discretionary' : 'y'
}

temp_file = ''

# Change to the TWS home directory, create the watchlist directory if needed, and
# create a temporary file to write the spreadsheets into.

def init():
    global tws_home, watchlists, temp_file
    system = platform.system()

    if system != 'Windows':
        tws_home = os.getenv("HOME")
    os.chdir(os.path.join(tws_home, 'Jts'))

    if watchlists != '':
        os.makedirs(watchlists, exist_ok=True)
        os.chdir(watchlists)

    fd, temp_file = mkstemp(suffix = extension)
    os.close(fd)

# Remove the temporary file used to store spreadsheets.

def finish():
    global temp_file
    os.unlink(temp_file)

# Populate the ETF lists by downloading the Excel spreadsheets from State Street and
# extracting the columns needed from them.

def createCSVs(etf_list, subdir = '', update_suffixes = False, split_sectors = False):
    global suffixes, temp_file

    tickers = {}

    for etf in etf_list:

# Download a spreadsheet.

        url = baseurl + etf + extension
        r = requests.get(url, allow_redirects=True)
        with open(temp_file, 'wb') as f:
            f.write(r.content)

# Extract the data from the spreadsheet and populate a dictionary.

        ps = openpyxl.load_workbook(temp_file)
        sheet = ps[worksheet]
        for row in range(first_row, sheet.max_row + 1):
            sector = sheet['F' + str(row)].value
            if sector == no_sector:
                break

# Add any new sector to the suffixes list.  This should only be done on the Select
# SPDR ETFs.

            if update_suffixes == True:
                suffixes[sector] = etf[2:]
            ticker = sheet['B' + str(row)].value

# If splitting the ETF into sectors, add the same suffix that State Street uses to the
# file name to generate the siumlated symbol for the ETF.

            if split_sectors == True:
                csv = etf + suffixes[sector]
            else:
                csv = etf

# Set up the ticker data for this ETF and any sector ETF.

            if csv not in tickers:
                tickers[csv] = {}
                tickers[csv]['Base ETF'] = etf
                tickers[csv]['Equities'] = {}
            tickers[csv]['Equities'][ticker] = 1

# Create the .csv files for each ETF, whether real or simulated.

    for etf in tickers:
        dir = ''

# If splitting the ETF, or creating the files in a new directory, set it up.

        if subdir != '':
            dir = subdir
            os.makedirs(dir, exist_ok=True)
        else:
            if split_sectors == True:
                dir = tickers[etf]['Base ETF']
                os.makedirs(dir, exist_ok=True)
        csv = os.path.join(dir, etf + '.csv')

# Create the .csv.

        with open(csv, 'w') as f:
            f.write("COLUMN,0\n")
            for ticker in tickers[etf]['Equities']:
                f.write('DES,' + ticker + ',STK,SMART/AMEX,,,,,\n')

# Main program.

init()
createCSVs(sspdrs,update_suffixes=True)
createCSVs(igetfs, subdir='spy')
createCSVs(simetfs,split_sectors=True)
finish()
