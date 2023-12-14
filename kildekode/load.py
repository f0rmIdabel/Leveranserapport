# load.py
import pandas as pd

def init_workbook(transporter, week, raw, path):
    """
    Initialiserer excel-fil for transportøren.
    """
    writer = pd.ExcelWriter(path+'utfil/' + transporter + ' - uke ' + str(week) + '.xlsx', engine='xlsxwriter')
    workbook=writer.book
    # Spesifiser format
    format_header = workbook.add_format({'bold': True, 'font_size': 18})
    format_header2 = workbook.add_format({'bold': True, 'font_size': 14})
    format_sum = workbook.add_format({'bold': True, 'font_size': 12, 'border':True, 'bg_color':'yellow', 'num_format': '### ### ##0.00'})

    # Lag loggarket
    worksheet0=workbook.add_worksheet("Logg")
    writer.sheets["Logg"] = worksheet0
    raw.to_excel(writer, sheet_name="Logg",index=False)
    return writer, workbook, [format_header, format_header2, format_sum]

def write_pivot_to_excel(writer, workbook, formats, piv, sheetname, uke):
    """
    Skriver pivot-tabellen til excel-filen.
    """
    
    worksheet=workbook.add_worksheet(sheetname)
    writer.sheets[sheetname] = worksheet

    worksheet.write_string(0, 0, "Leveranserapport - uke "+str(uke), formats[0])
    worksheet.write_string(2, 0, "Palleoversikt", formats[1])

    piv = piv[piv.columns[:-2]]
    piv.sort_values(by=['Kundenavn']).to_excel(writer,sheet_name=sheetname,startrow=4 , startcol=0, index=False)

    return None 

def write_summary_to_excel(writer, workbook, formats, piv, gas, df_TCO, uke):
    """
    Skriv oppsummeringsinfo til excel-filen.
    """
    sname = "Oppsummering"
    worksheet=workbook.add_worksheet(sname)
    writer.sheets[sname] = worksheet

    worksheet.write_string(0, 0, "Leveranserapport - uke "+str(uke), formats[0])
    worksheet.write_string(2, 0, "Palleoversikt", formats[1])

    piv.sort_values(by=['Pris per palle', 'Kundenavn']).to_excel(writer,sheet_name=sname,startrow=4 , startcol=0, index=False)
    worksheet.write_string(piv.shape[0] + 7, 0, "Drivstoffpris (kr/liter)", formats[1])
    gas.round(3).to_excel(writer,sheet_name=sname,startcol=0, startrow=piv.shape[0] + 9, index=False)

    worksheet.write_string(piv.shape[0] + 13, 0, "Drivstofftillegg", formats[1])
    df_TCO.round(3).to_excel(writer,sheet_name=sname,startcol=0, startrow=piv.shape[0] + 15, index=False)

    return worksheet, writer 

def write_sum_to_excel(writer, worksheet, formats, df_sum):
    """
    Skriv totalsummen til excel-filen.
    """
    sname = "Oppsummering"

    worksheet.write_string(2, 11 + 2, "Endelig sum", formats[0])
    df_sum.round(2).to_excel(writer,sheet_name=sname,startcol=11 + 2, startrow=4, index=False)
    
    total_sum = df_sum["Pris (kr)"].sum()
    total_sum_drivstoff = df_sum["Drivstofftillegg"].sum()
    worksheet.write_string(5+df_sum.shape[0], 11 + 3, str(round(total_sum,2)))
    worksheet.write_string(5+df_sum.shape[0], 11 + 4, str(round(total_sum_drivstoff,2)))
    worksheet.write_string(5+df_sum.shape[0], 11 + 5, str(round(total_sum + total_sum_drivstoff,2)), formats[2])
    worksheet.write_string(5+df_sum.shape[0], 11 + 2, "SUM", formats[1])
    return None

def append_to_yearly_logfile(df, path):
    """
    Legg til rådata i loggfilen. 
    Hvert år har sin egen fil, og filen 
    opprettes dersom den ikke finnes fra før.   
    """

    years = pd.to_datetime(df["Leveringsdato"]).dt.year
    df['Years'] = years
    years = years.unique()

    # Legg til rådata i loggfilen per år
    for year in years:
        try:
            logfile = pd.read_excel(path+"loggfil/Leveranserapport - " + str(year) + ".xlsx")
            #logfile = logfile.append(raw[raw["Years"]==year])  deprecated 
            logfile = pd.concat([logfile, df[df["Years"]==year]], ignore_index=True)
            logfile = logfile.drop_duplicates()
            logfile.to_excel(path+"loggfil/Leveranserapport - " + str(year) + ".xlsx", index=False)
        except FileNotFoundError:
            df[df["Years"]==year].to_excel(path+"loggfil/Leveranserapport - " + str(year) + ".xlsx", index=False)

    return None
