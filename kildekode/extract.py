#extract.py
import pandas as pd 
import os
from transform import correct_Nortura

def read_infile():
    """
    Leser innfilen og returnerer en noe redigert dataframe:
        * Fjerner whitespace fra kundenavn og transportør
        * Ekskluderer internleveranser
        * Korrigerer feil i Nortura-data
    """

    # Hent innfil-mappen
    files = os.listdir('../innfil/')
    if ".gitkeep" in files:
        files.remove(".gitkeep")
    # Sjekk at det kun ligger en fil i mappen
    if len(files) > 1:
        print(files)
        print("Det ligger flere filer i mappen, eller filen er åpen.")
        exit(1)

    # Les innfilen og fjern whitespace
    raw = pd.read_excel(r'../innfil/'+files[0])
    raw["Transportør"] = raw["Transportør"].str.strip()
    raw["Kundenavn"] = raw["Kundenavn"].str.strip()

    # Ekskluder interne kunder
    customers = get_customer_pricelist()
    customers_excluded = customers[customers["Område"]=="Internt"]["Kundenavn"].unique()
    raw = raw[~raw["Kundenavn"].isin(customers_excluded)].reset_index(drop=True)

    # Korriger feil i Nortura-data
    raw = correct_Nortura(raw)

    return raw

def get_input(sheet):
    """
    Leser excelarket basert på arknavn 
    og returnerer en dataframe
    """
    df = pd.read_excel(r'../prisliste.xlsx', decimal=",",  sheet_name=sheet)
    return df

def get_customer_pricelist():
    """
    Returnerer dataframe med kundespesifikk prisliste
    """
    customer = get_input("Kunder")
    customer["Kundenavn"] = customer["Kundenavn"].str.strip()

    return customer

def get_median_week(df):
    """
    Returnerer medianuke i dataframe
    """
    return df["Dato"].dt.isocalendar().week[int(len(df)/2)]

def get_gas_price(uke):
    """
    Returnerer dataframe med drivstoffpris
    """

    gas = get_input("Drivstoff")
    gas["Snitt uke "+str(uke)] = gas.iloc[:,-7:].mean(axis=1)
    gas = gas[list([gas.columns[0]]) +list([gas.columns[-1]])+list(gas.columns[-8:-1])]

    return gas
