# ETL-rutine for Leveranserapport

Transformerer den ukentlige leveranserapporten, slik at dataen splittes og oppsummeres per leverandør. 

## Instruksjoner for å kjøre på egen pc

1. Kopier repoet til egen pc. 
2. Lag et nytt python-environment og kjør 
```python
 pip install -r requirements.txt
 ```
2. Oppdater variablen PATH i etl.py
3. Legg inn ukens leveranserapport i mappen innfil. 
4. Kjør etl.py

## Videreutvikling

Programmet er strukturert i scriptene

* extract.py: innhenting av data
* transform.py: transformer dataen til ønsket format
* load.py: skriv data til excel-filer

Transformasjonsdelen er basert på veldig spesifikke regler. Dersom regelsettet endres, kan funksjoner legges til her, og påføres dataen i etl.py (hovedscriptet). 