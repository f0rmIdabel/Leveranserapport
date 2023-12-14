"""
Hovedscript for å kjøre ETL-prosessen.
Sett filsti til kildekode-mappen i PATH-variabelen.
"""

PATH = 'C:/Leveranserapport/'

#############################################################################################
# PROGRAM START
#############################################################################################

import pandas as pd
import numpy as np
pd.options.mode.chained_assignment = None  # default='warn'
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
import sys


sys.path.insert(1, PATH+'kildekode/')

from extract import read_infile, get_input, get_customer_pricelist, get_median_week, get_gas_price
from transform import extract_relevant_columns, get_df_TCO, get_df_sum, get_pivot, get_turtype_sorted
from load import append_to_yearly_logfile, write_pivot_to_excel, write_summary_to_excel, write_sum_to_excel, init_workbook            

if __name__ == "__main__":

    # Les inn rådata
    raw = read_infile(PATH)

    # Legg til rådata i loggfilen
    append_to_yearly_logfile(raw, PATH)

    # Hent ut relevante kolonner fra rådataen
    data = raw.copy()
    data = extract_relevant_columns(data)

    # Hent ut data fra input-filer
    uke = get_median_week(data)
    gas = get_gas_price(uke, PATH)
    pricelist = get_customer_pricelist(PATH) 
    transporters = get_input("Transportører", PATH)
    cars = get_input("Biler", PATH)

    # Loop over transportører i input-filen
    for t in transporters["Transportør"].unique():

        transporter = transporters[transporters["Transportør"]==t]
        raw_data_transporter = raw[raw["Transportør"]==t]
        data_transporter = data[data["Transportør"]==t]

        # Lag excel-fil for transportøren
        writer, workbook, formats = init_workbook(t, uke, raw_data_transporter, PATH)

        # Lag pivot-tabell for hver turtype (Termobil, Fastpris-rutene og vanlig Pallepris-ruter)
        total_pris = []
        bidrag = []
       
       # Sorter turtypene slik at vanlig pallepris kommer først, og termobil sist
        turtype = get_turtype_sorted(data_transporter)

        # Gå gjennom hver turtype og lag pivot-tabell
        for type in turtype:
            data_transporter_ = data_transporter[data_transporter["Turtype"]==type]

            if len(data_transporter_) == 0:
                continue

            # Lag pivot-tabell
            piv = get_pivot(data_transporter_, pricelist)

            if type == "Termobil":
                write_pivot_to_excel(writer, workbook, formats, piv, type, uke)
 
            elif type[:3] == "BIL":
                write_pivot_to_excel(writer, workbook, formats, piv, type, uke)
                bidrag.append(type)
                total_pris.append(cars[cars["Bil"]==int(type[4:])]["Pris"].iloc[0])

            elif type == "Pallepris":
                bidrag.append(type)
                total_pris_paller = piv.iat[-1,-1]
                total_pris.append(total_pris_paller)
                df_TCO = get_df_TCO(gas, transporter, uke)
                worksheet,writer = write_summary_to_excel(writer, workbook, formats, piv, gas, df_TCO, uke)  
            
        # Skriv totalsummen til excel-arket
        df_sum = get_df_sum(df_TCO, transporter, bidrag, total_pris) 
        write_sum_to_excel(writer,  worksheet, formats, df_sum)           
                
        # Lukk og lagre fil
        writer.close()

#############################################################################################
# PROGRAM END
#############################################################################################