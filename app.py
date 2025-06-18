#!/usr/bin/python3

import openpyxl

SEKTORS_HEADER = "SEKTOR"
ETFS_HEADERS = ["ETF","ISIN","POSITION","TICKER","TITEL","SEKTOR","PERCENT"]
PORTFOLIO_HEADERS = ["BANK","titel","id (ISIN)","pieces","sector","value","currency"]

def get_contiguous_range(ws):
    """
    read etf data from any sheet in workbook
    :param ws: worksheet with many columns and many rows
    :return: list of lists
    """
    csv = []
    for row in ws.values:
        csv.append(list(row))
    return csv

def get_sectors(ws):
    """
    read sector data from worksheet SEKTOR
    :param ws: worksheet with one column and many rows
    :return: list of sectors
    """
    sectors_li = get_contiguous_range(ws)
    sectors = []
    for sector in sectors_li:
        if sector[0] is None:
            break
        sectors.append(sector[0])
    if sectors[0] != SEKTORS_HEADER:
        raise Exception('SEKTORS sheet has illegal header in cell A1')
    return sectors[1:] # w.o. header

def get_etfs(ws):
    """
    read etf data from worksheet ETFS
    :param ws: worksheet with many columns and many rows
    :return: list of lists
    """
    etfs = get_contiguous_range(ws)
    for row in etfs:
        if len(row) != 7:
            raise Exception('ETFS has row with illegal length:'+str(row))
    if etfs[0] != ETFS_HEADERS:
        raise Exception("ETFS Sheet has illegal header row:"+str(etfs[0]))
    return etfs[1:] # w.o. headers

def get_portfolio(ws):
    """
    read portfolio data from worksheet PORTFOLIO
    :param ws: worksheet with many columns and many rows
    :return: list of lists
    """
    portfolio = get_contiguous_range(ws)
    for row in portfolio:
        if len(row) != 7:
            raise Exception('PORTFOLIO has row with illegal length:'+str(row))
    if portfolio[0] != PORTFOLIO_HEADERS:
        raise Exception("PORTFOLIO Sheet has illegal header row:"+str(portfolio[0]))
    return portfolio[1:] # w.o. headers

def calculate(sectors, etfs, portfolio):
    """
    calculate investment in each sector
    :param sectors: sectors array
    :param etfs: list of lists with etf data
    :param portfolio: list of lists with portfolio data
    :return: list of lists with total investments per sector
    """
    try:
        # build values, one for each sector
        values = []
        for sector in sectors:
            values.append(list((sector, 0)))
        # add values from the portfolio
        for item in portfolio:
            if item[0] is None:
                break
            sec = item[4] # item sector
            val = item[5] # item value
            if sec != 'ETF':
                # stock in portfolio
                idx = sectors.index(sec)
                values[idx][1] += val # add value array
            else:
                # exchange traded fund in portfolio
                tit = item[1] # etf titel
                for row in etfs:
                     if row[0] == tit:
                         # rows with sectors in the exchange traded fund
                         sect = row[5] # etf row sector
                         pct = row[6]  # etf row percentage
                         frac = int(val*pct/100) # calculate etf row fraction of value
                         idx = sectors.index(sect)
                         values[idx][1] += frac # add fraction to array
                     pass
                pass
            pass
        return values
    except TypeError:
        print('Shit happens')
    except:
        print('More shit')


def display_in_workbook(source_file, wb, values):
    """
    display the investments per sector in sheet SEKTOR
    :param values: workbook, write enabled
    :param values: list of lists with total investments per sector
    :return: None
    """
    print(values)
    ws = wb["SEKTORS"]
    for y in range(2, len(values)+2, 1):
        ws.cell(row=y, column=2).value = values[y-2][1]
    wb.save(source_file)
    pass

def run(source_file):
    try:
        # open workbook
        wb = openpyxl.load_workbook(source_file, read_only=False, data_only=True)
        # SEKTORS ====
        sectors_ws = wb["SEKTORS"]
        sectors = get_sectors(sectors_ws)
        # ETFS ====
        etfs_ws = wb["ETFS"]
        etfs = get_etfs(etfs_ws)
        # PORTFOLIO ====
        portfolio_ws = wb["PORTFOLIO"]
        portfolio = get_portfolio(portfolio_ws)
        # CALCULATION ====
        values = calculate(sectors, etfs, portfolio)
        display_in_workbook(source_file, wb, values)
        exit(0)
    except Exception as err:
        print(err.args)
        exit(1)

if __name__ == "__main__":
    run('./docs/sektoren.xlsx')
    exit(0)