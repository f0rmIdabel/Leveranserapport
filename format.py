import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)
import numpy as np
import os


def read_infile():
    """
    Leser innfilen og returnerer en noe redigert dataframe:
        * Fjerner whitespace fra kundenavn og transportør
        * Ekskluderer internleveranser
        * Korrigerer feil i Nortura-data
    """

    # Hent innfil-mappen
    files = os.listdir('innfil/')
    if ".gitkeep" in files:
        files.remove(".gitkeep")
    # Sjekk at det kun ligger en fil i mappen
    if len(files) > 1:
        print(files)
        print("Det ligger flere filer i mappen, eller filen er åpen.")
        exit(1)

    # Les innfilen og fjern whitespace
    raw = pd.read_excel(r'innfil/'+files[0])
    raw["Transportør"] = raw["Transportør"].str.strip()
    raw["Kundenavn"] = raw["Kundenavn"].str.strip()

    # Ekskluder interne kunder
    customers = get_customer_pricelist()
    customers_excluded = customers[customers["Område"]=="Internt"]["Kundenavn"].unique()
    raw = raw[~raw["Kundenavn"].isin(customers_excluded)].reset_index(drop=True)

    # Korriger feil i Nortura-data
    raw = correct_Nortura(raw)

    return raw

def correct_Nortura(df, change_log=False):
    """
    Korrigerer antatte feil i Nortura-data: 
        Dersom antall paller er 2 og vekten er under 50 kg,
        settes antall paller til 1.
    Kan også generere en endringslogg, men denne brukes ikke per 
    Desember 2023, er slått av, og kan antagelig fjernes på sikt. 
    """

    #######################################################################################
    # Generer endringslogg for Nortura. Fjernes?
    #######################################################################################
    if change_log:
        df_Nortura = df[['Transportør', 'Turnavn', 'Kundenavn','Leveringsdato',
                         'Ant paller tot på levering',
                         'Ant paller Nortura',
                         'Vekt Nortura']]

        df_Nortura = df_Nortura[df_Nortura['Ant paller Nortura'] == 2]\
                               [df_Nortura['Vekt Nortura'] < 50]\
                               .sort_values(by=['Transportør', 'Turnavn','Leveringsdato'])
        
        df_Nortura['Ant paller Nortura korrigert'] = np.ones(len(df_Nortura))

        if len(df_Nortura) > 0:
            df_Nortura.to_excel('Nortura endringslogg.xlsx', index=False)

    #######################################################################################
    # Korriger Nortura
    #######################################################################################
    for i in range(len(df)):
        if df['Ant paller Nortura'][i] == 2:
            if df['Vekt Nortura'][i] < 50:
                df['Ant paller Nortura'][i] = 1

    return df 

def append_to_yearly_logfile(df):
    """
    Legg til rådata i loggfilen. 
    Hvert år har sin egen fil, og filen 
    opprettes dersom den ikke finnes fra før.   
    """

    years = pd.to_datetime(df["Leveringsdato"]).dt.year
    raw['Years'] = years
    years = years.unique()

    # Legg til rådata i loggfilen per år
    for year in years:
        try:
            logfile = pd.read_excel("loggfil/Leveranserapport - " + str(year) + ".xlsx")
            #logfile = logfile.append(raw[raw["Years"]==year])  deprecated 
            logfile = pd.concat([logfile, df[df["Years"]==year]], ignore_index=True)
            logfile = logfile.drop_duplicates()
            logfile.to_excel("loggfil/Leveranserapport - " + str(year) + ".xlsx", index=False)
        except FileNotFoundError:
            df[df["Years"]==year].to_excel("loggfil/Leveranserapport - " + str(year) + ".xlsx", index=False)

    return None

def get_input(sheet):
    """
    Leser excelarket basert på arknavn 
    og returnerer en dataframe
    """
    df = pd.read_excel(r'prisliste.xlsx', decimal=",",  sheet_name=sheet)
    return df

def extract_relevant_columns(df):
    """
    Redigerer dataframe for å få relevante kolonner:
        * Fjerner unødvendige kolonner fra dataframe
        * Summerer totalt antall paller Nødvendig siden
          "Ant paller på levering"-kolonnen er feil
        * Legger til kolonner for Termobil 
        * Legger til kolonner for ukedag
        * Fjerner duplikater
    """

    # Fjern unødvendige kolonner
    df = df[['Transportør', 'Kundenavn',
             'Turnavn', 'Leveringsdato',
             'Ant paller tot på levering',
             'Ant paller Tørr', 'Ant paller Kjøl',
             'Ant paller Fersk', 'Ant paller Frys', 'Ant paller Nortura',
             'Ant paller Q', 'Ant paller TM', 'Ant paller RDI',
             'Vekt Nortura']]

    # Summer totalt antall paller
    df['Ant paller summert'] = df['Ant paller Tørr'] + df['Ant paller Kjøl'] + \
                             df['Ant paller Fersk'] +  df['Ant paller Frys'] +\
                             df['Ant paller Nortura'] + df['Ant paller Q'] +\
                             df['Ant paller TM'] +  df['Ant paller RDI']

    # Legg til kolonne for Termobil
    df = categorise_route(df)

    # Legg til kolonner for ukedag
    df["Dato"] = pd.to_datetime(df["Leveringsdato"])
    df["Ukedag"] = df["Dato"].dt.weekday
    ukedag_lib  = {0:'Mandag', 1:'Tirsdag', 2:'Onsdag', 3:'Torsdag', 4:'Fredag', 5:'Lørdag', 6:'Søndag'}
    df["Ukedag navn"] = [ukedag_lib[i] for i in df["Dato"].dt.weekday]


    # Fjern duplikater
    df = df.drop_duplicates().reset_index(drop=True)

    # Velg relevante kolonner
    df = df[['Transportør', 'Kundenavn', "Turtype",
             'Dato', 'Ukedag', 'Ukedag navn',
             'Ant paller summert']]
    return df

def categorise_route(df):

    termo = ["TERMO" in str(tur).upper() for tur in df["Turnavn"]]
    fastpris = [" RUTE " in str(tur).upper() for tur in df["Turnavn"]]
    df["Turtype"] = ["Termobil" if t 
                     else "BIL " +df.loc[i,"Turnavn"].split()[-1] if f 
                     else "Pallepris" 
                     for t,f, i in zip(termo, fastpris, range(len(df)))]

    return df

def get_median_week(df):
    """
    Returnerer medianuke i dataframe
    """
    return df["Dato"].dt.isocalendar().week[int(len(df)/2)]

def get_gas_price():
    """
    Returnerer dataframe med drivstoffpris
    """

    gas = get_input("Drivstoff")
    gas["Snitt uke "+str(uke)] = gas.iloc[:,-7:].mean(axis=1)
    gas = gas[list([gas.columns[0]]) +list([gas.columns[-1]])+list(gas.columns[-8:-1])]


    return gas

def get_customer_pricelist():
    """
    Returnerer dataframe med kundespesifikk prisliste
    """
    customer = get_input("Kunder")
    customer["Kundenavn"] = customer["Kundenavn"].str.strip()

    return customer

def get_df_TCO(gas, transporter):
    df_TCO = pd.DataFrame()
    df_TCO["Prisøkning (kr/liter)"] = gas["Snitt uke "+str(uke)] - gas[gas.columns[0]]
    df_TCO["Prisøkning (%)"] = df_TCO["Prisøkning (kr/liter)"] / gas[gas.columns[0]] *100
    df_TCO["Andel av TCO (%)"] = float(transporter["Diesel"].iloc[0])*100
    df_TCO["Økning i TCO (%)"] = df_TCO["Prisøkning (%)"]/100 * df_TCO["Andel av TCO (%)"]
    
    return df_TCO

def get_df_sum(df_TCO, transporter, bidrag, total_pris):

    bidrag.append("T&M")
    total_pris.append(transporter["TM"].iloc[0])
    bidrag.append("Samlet drivstofftillegg") 
    drivstofftillegg = np.sum(np.asarray(total_pris)*float(df_TCO["Økning i TCO (%)"].iloc[0])/100)
    total_pris.append(drivstofftillegg)
    df_sum = pd.DataFrame()
    df_sum["Bidrag"] = bidrag
    df_sum["Pris (kr)"] = total_pris
    df_sum["Drivstofftillegg, " + str(round(float(df_TCO["Økning i TCO (%)"].iloc[0]),1)) + " %"] = df_sum["Pris (kr)"]*float(df_TCO["Økning i TCO (%)"].iloc[0])/100

    return df_sum

def get_pivot(data, pricelist):
    # Lag pivot-tabell med kunder og leveringsdag
    piv = pd.pivot_table(data, aggfunc='sum',\
                        values='Ant paller summert', columns='Ukedag navn', index='Kundenavn')\
                        .reindex(['Mandag', 'Tirsdag', 'Onsdag','Torsdag','Fredag','Lørdag','Søndag'],axis=1)
    piv['Totalt antall paller'] = piv.sum(axis=1)

    ####################################################################################
    # D. Danielsen spesieltilfelle
    ####################################################################################

    # Sett minimum antall paller for D. Danielsen til 10
    if "D. Danielsen AS" in piv.index:
        if piv.loc["D. Danielsen AS"]["Totalt antall paller"] < 10:
            piv.loc["D. Danielsen AS"]["Totalt antall paller"] = 10


    # Legg til prisliste
    piv = piv.merge(pricelist, how='left',on='Kundenavn')
    piv["Total pris"] = piv["Totalt antall paller"] * piv["Pris"]

    ####################################################################################
    # Kosmetikk
    ####################################################################################

    # Regn ut totalene
    piv.loc["Totalt"] = piv.sum(numeric_only=True)
    piv=piv.rename(columns={"Pris":"Pris per palle"})
    piv.iat[-1,0] = 'Totalt'
    piv.iat[-1,-2] = None
    piv.iat[-1,-2] = None

    piv.index.name = None
    piv.columns.name = None

    # Rund av til to siffer
    piv = piv.round(2)
    
    return piv 

def init_workbook(transporter, week, raw):
    writer = pd.ExcelWriter('utfil/' + transporter + ' - uke ' + str(week) + '.xlsx', engine='xlsxwriter')
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

def write_termo_to_excel(writer, workbook, formats, piv):
    sname = "Termobil"

    worksheet=workbook.add_worksheet(sname)
    writer.sheets[sname] = worksheet

    worksheet.write_string(0, 0, "Leveranserapport - uke "+str(uke), formats[0])
    worksheet.write_string(2, 0, "Palleoversikt", formats[1])

    piv = piv[piv.columns[:-2]]
    piv.sort_values(by=['Kundenavn']).to_excel(writer,sheet_name=sname,startrow=4 , startcol=0, index=False)

    return None 

def write_cars_to_excel(writer, workbook, formats, piv, type):
    
    sname = type

    worksheet=workbook.add_worksheet(sname)
    writer.sheets[sname] = worksheet

    worksheet.write_string(0, 0, "Leveranserapport - uke "+str(uke), formats[0])
    worksheet.write_string(2, 0, "Palleoversikt", formats[1])

    piv = piv[piv.columns[:-2]]
    piv.sort_values(by=['Kundenavn']).to_excel(writer,sheet_name=sname,startrow=4 , startcol=0, index=False)

    return None 


def write_to_excel(writer, workbook, formats, piv, gas, df_TCO):
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
    sname = "Oppsummering"

    worksheet.write_string(2, 11 + 2, "Endelig sum", formats[0])
    df_sum.round(2).to_excel(writer,sheet_name=sname,startcol=11 + 2, startrow=4, index=False)
    
    total_sum = df_sum["Pris (kr)"].sum()
    worksheet.write_string(8+12, 11 + 3, str(round(total_sum,2)), formats[2])

    return None

if __name__ == "__main__":

    raw = read_infile()
    append_to_yearly_logfile(raw)

    data = raw.copy()
    data = extract_relevant_columns(data)

    uke = get_median_week(data)
    gas = get_gas_price()
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

        for type in data_transporter["Turtype"].unique():
            data_transporter_ = data_transporter[data_transporter["Turtype"]==type]

            if len(data_transporter_) == 0:
                continue

            # Lag pivot-tabell
            piv = get_pivot(data_transporter_, pricelist)

            # Skriv til excel
            if type == "Termobil":
                write_termo_to_excel(writer, workbook, formats, piv)

 
            elif type[:3] == "BIL":
                write_cars_to_excel(writer, workbook, formats, piv, type)
                bidrag.append(type)
                total_pris.append(cars[cars["Bil"]==int(type[4:])]["Pris"].iloc[0])
                

            elif type == "Pallepris":
                bidrag.append("Paller")
                total_pris_paller = piv.iat[-1,-1]
                total_pris.append(total_pris_paller)
                df_TCO = get_df_TCO(gas, transporter)
                
                worksheet,writer = write_to_excel(writer, workbook, formats, piv, gas, df_TCO)  
            
        df_sum = get_df_sum(df_TCO, transporter, bidrag, total_pris) 
        write_sum_to_excel(writer,  worksheet, formats, df_sum)           
                
        # Lukk og lagre fil 
        writer.close()