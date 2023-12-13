import pandas as pd
import numpy as np
pd.options.mode.chained_assignment = None  # default='warn'
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)

from extract import read_infile, get_input, get_customer_pricelist, get_median_week, get_gas_price
from transform import extract_relevant_columns, get_df_TCO, get_df_sum, get_pivot, get_turtype_sorted
from load import append_to_yearly_logfile, write_pivot_to_excel, write_summary_to_excel, write_sum_to_excel, init_workbook            

if __name__ == "__main__":
    raw = read_infile()
    append_to_yearly_logfile(raw)

    data = raw.copy()
    data = extract_relevant_columns(data)

    uke = get_median_week(data)
    gas = get_gas_price(uke)
    pricelist = get_customer_pricelist() 
    transporters = get_input("Transportører")
    cars = get_input("Biler")

    for t in transporters["Transportør"].unique():

        transporter = transporters[transporters["Transportør"]==t]
        raw_data_transporter = raw[raw["Transportør"]==t]
        data_transporter = data[data["Transportør"]==t]
        writer, workbook, formats = init_workbook(t, uke, raw_data_transporter)

        ####################################################################################
        # Håndter pallepriser, tørrbil og fastpris-biler separat 
        ####################################################################################

        total_pris = []
        bidrag = []
       
        turtype = get_turtype_sorted(data_transporter)

        for type in turtype:
            data_transporter_ = data_transporter[data_transporter["Turtype"]==type]

            if len(data_transporter_) == 0:
                continue

            # Lag pivot-tabell
            piv = get_pivot(data_transporter_, pricelist)

            # Skriv til excel
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
            
        df_sum = get_df_sum(df_TCO, transporter, bidrag, total_pris) 
        write_sum_to_excel(writer,  worksheet, formats, df_sum)           
                
        # Lukk og lagre fil 
        writer.close()