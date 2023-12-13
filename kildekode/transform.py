#transform.py
import pandas as pd
import numpy as np

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

def get_df_TCO(gas, transporter, uke):
    df_TCO = pd.DataFrame()
    df_TCO["Prisøkning (kr/liter)"] = gas["Snitt uke "+str(uke)] - gas[gas.columns[0]]
    df_TCO["Prisøkning (%)"] = df_TCO["Prisøkning (kr/liter)"] / gas[gas.columns[0]] *100
    df_TCO["Andel av TCO (%)"] = float(transporter["Diesel"].iloc[0])*100
    df_TCO["Økning i TCO (%)"] = df_TCO["Prisøkning (%)"]/100 * df_TCO["Andel av TCO (%)"]
    
    return df_TCO

def get_df_sum(df_TCO, transporter, bidrag, total_pris):

    if transporter["TM"].iloc[0] > 0:
        bidrag.append("T&M")
        total_pris.append(transporter["TM"].iloc[0])
    df_sum = pd.DataFrame()
    df_sum["Bidrag"] = bidrag
    df_sum["Pris (kr)"] = total_pris
    df_sum["Drivstofftillegg"] = df_sum["Pris (kr)"]*float(df_TCO["Økning i TCO (%)"].iloc[0])/100

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
    piv = piv.merge(pricelist[["Kundenavn", "Pris"]], how='left',on='Kundenavn')
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

def get_turtype_sorted(data):

    turtype_sorted = []
    turtype = data["Turtype"].unique()

    if "Pallepris" in turtype:
        turtype_sorted.append("Pallepris")

    if sum([x[:3] == "BIL" for x in turtype]) > 0:
        temp = [x[4:] for x in turtype if x[:3] == "BIL"]
        temp = np.array(temp).astype(int)
        for i in np.argsort(temp):
            turtype_sorted.append("BIL " + str(temp[i]))
    
    if "Termobil" in turtype:
        turtype_sorted.append("Termobil")

    return turtype_sorted