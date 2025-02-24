#!/usr/bin/env python
# coding: utf-8

import time
start_time = time.time()


import sys
import subprocess

# Install packages if missing, useful with pyinstaller and making .exes
required_packages = ["azure-identity", "sqlalchemy", "pyodbc", "pandas", "pywinauto", "xlsxwriter"]

def install_missing_packages(packages):
    for package in packages:
        try:
            __import__(package.replace("-", "_"))
        except ImportError:
            print(f"Pacchetto '{package}' non trovato. Installazione in corso...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            print(f"'{package}' installato.")

install_missing_packages(required_packages)


import welcome_derto # Simple welcome function, first try at uploading a module to PyPi
import azure.identity
import pyodbc
assert pyodbc
assert azure.identity
from sqlalchemy import create_engine, text
import pandas as pd
import shutil
import urllib
import os
from threading import Thread
from pywinauto import Desktop
from pywinauto.keyboard import send_keys
import datetime as dt
from sqlalchemy.exc import OperationalError
import xlsxwriter
import json

welcome_derto.welcome_user_anfia()

def move_window_to_primary_monitor(window):
    """
    Focus window and avoid writing passwords all over the place.
    """
    window.move_window(x=0, y=0)

def get_login_info_from_config(): # Get login info from config or create one.
    config_file = "config.txt"
    if not os.path.exists(config_file):
        user = input("Inserisci la tua mail di Microsoft: ")
        password = input("Inserisci la tua password: ")
        server = input("Inserisci il nome del server (es. sql-tuazienda-prod...): ")
        database = input("Inserisci il nome del database a cui accedere (es. BuyAnalysis...): ")
        with open(config_file, 'w') as file:
            file.write(user + "\n")
            file.write(password + "\n")
            file.write(server + "\n")
            file.write(database)
        print(f"User e password salvati nel file {config_file}.")
    else:
        with open(config_file, 'r') as file:
            user = file.readline().strip()
            password = file.readline().strip()
            server = file.readline().strip()
            database = file.readline().strip()
    return user, password, server, database

def simulate_user_login(user, password):
    """
    This function takes the result of the previous function and automatically performs the login procedure for Microsoft-based applications.
    
    Parameters
    ----------
    Parameters are passed by the function get_login_info_from_config().
    """
    try:
        
        time.sleep(3.5)
        app = Desktop(backend='win32').window(title_re=".*autenticaz.*", visible_only=False)
        dlg = app
        dlg.set_focus()
        move_window_to_primary_monitor(dlg)
        time.sleep(1.5)

        send_keys(user)
        send_keys('{ENTER}{TAB}{ENTER}', with_spaces=True)

        time.sleep(2)
        if dlg.wait('ready', timeout=10):
            send_keys(password)
            send_keys('{ENTER}') 
            print("Login completato.")
        else:
            print("Login completato. La password non è stata necessaria.")
            pass
        
    except Exception as e:
        print(f"Errore durante la simulazione del login: {e}. Il login è stato effettuato? Allora non preoccuparti, questo messaggio è normale.")

def convert_csv_to_xlsx(csv_file, xlsx_file, output_folder): 
    """
    This function transforms the .csv, resulting from the SQLalchemy query, into an .xlsx 
    that will then be sent to the customer. I found this approach to be faster than
    directly the results as .xlsx. It must also be considered that, as the result
    is quite large and easily exceeds the Excel row limit, it is a way more elegant method.
    """
    os.makedirs(output_folder, exist_ok=True)
    

    xlsx_file_path = os.path.join(output_folder, xlsx_file)
    
    workbook = xlsxwriter.Workbook(xlsx_file_path)
    worksheet = workbook.add_worksheet("Immatricolazioni")
    header_format = workbook.add_format({
        'bold': True,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })

    # Apri il file CSV e scrivi riga per riga
    with open(csv_file, "r", encoding="utf-8") as f:
        for row_idx, line in enumerate(f):
            for col_idx, cell_value in enumerate(line.strip().split(",")):
                if row_idx == 0:
                    worksheet.write(row_idx, col_idx, cell_value, header_format)
                else:
                    worksheet.write(row_idx, col_idx, cell_value)

    workbook.close()
    copied_file = os.path.join(r"L:\09.Knime\BKP_DR_BI_dati", os.path.basename(xlsx_file_path))
    shutil.copy2(xlsx_file_path, copied_file)
    os.remove(csv_file)  # Delete temporary CSV.

def load_dict_from_json(json_path):
    with open(json_path, "r", encoding="utf-8") as json_file:
        return json.load(json_file)


def verify_df_pairs(df, make_col="MARCA", model_col="MODELLO"):
    """
    Useful function to automatically correct known database errors that are still not resolved, and also to discover
    new errors. There is a .json file that is used to check each pair found in the dataframe against the .json file
    that contains all the correct pairs. When new pairs are added to the database, they are manually validated
    and added to the .json file.

    Parameters
    ----------
    df : dataframe. 
        The dataframe obtained by the SQL query.
    json_path : string. 
        The path of the .json file.
    make_col : string, optional.
        The first column for the make of the vehicle. The default is "MARCA".
    model_col : string, optional.
        The second column for the model of the vehicle. The default is "MODELLO".

    Returns
    -------
    df : dataframe
        Returns the corrected dataframe that will then be converted into csv.

    """
    user_input = ""
    while user_input.casefold() != 'skip':
        if os.path.exists(r"L:\03.Articoli_Analisi_(exUtenti)\Day-by-day\template_make_model.json"):
            json_path = r"L:\03.Articoli_Analisi_(exUtenti)\Day-by-day\template_make_model.json"
        elif not os.path.exists(r"L:\03.Articoli_Analisi_(exUtenti)\Day-by-day\template_make_model.json"):
            json_path = (os.getcwd())+r"\template_make_model.json"
        else:
            print("Attenzione: file dizionario non trovato. Inserisci il percorso del file dizionario o scrivi 'Skip' per saltare l'operazione.")
            user_input = input("Inserisci qui il percorso: ")
            json_path = user_input
        
        template_dict = load_dict_from_json(json_path)
        corrections = {
            ('PEUGEOT', 'JUMPER'): ('PEUGEOT', 'BOXER'),
            ('CITROEN', 'BOXER'): ('CITROEN', 'JUMPER'),
            ('CITROEN', 'DOBLÒ'): ('CITROEN', 'JUMPER'),
            ('RENAULT', 'SPRINTER'): ('RENAULT', 'MASTER'),
            ('FIAT', 'BERLINGO'): ('FIAT', 'DOBLÒ'),
            ('DS', 'ND'): ('DS', 'DS7')
        }
    
        corrected_pairs = []
        incorrect_pairs = []
        declared_pairs = set()
        
        for idx, row in df.iterrows():
            make, model = row[make_col], row[model_col]
            if pd.notna(make) and pd.notna(model):
                # Autocorrect if pair is in corrections list
                if (make, model) in corrections:
                    corrected_make, corrected_model = corrections[(make, model)]
                    df.at[idx, make_col] = corrected_make  # Aggiorna il DataFrame
                    df.at[idx, model_col] = corrected_model
                    corrected_pairs.append((idx, make, model, corrected_make, corrected_model))
                # Only save and later show pairs that are not in the dictionary or in corrections.
                elif make not in template_dict or model not in template_dict[make]:
                    incorrect_pairs.append((idx, make, model))
    
    
        if corrected_pairs:
            print("Correzioni applicate:")
        for (row_idx, old_make, old_model, new_make, new_model) in corrected_pairs:
            print(f"Riga {row_idx}: '{old_make} {old_model}' → '{new_make} {new_model}'")
                
        if incorrect_pairs:
            for row_idx, make, model in incorrect_pairs:
                if (make, model) not in declared_pairs:
                    print(f"Marca '{make}' - Modello '{model}' non è corretto.")
                    declared_pairs.add((make,model))
            print(declared_pairs)
        else:
            print("Tutte le coppie Marca-Modello sono corrette.")
        return df
    else:
        return None

current_month = dt.date.today().month
current_year = dt.date.today().year
previous_year = dt.date.today().year - 1


query_hy1 = """
DECLARE @start_date AS DATE = DATEFROMPARTS(YEAR(GETDATE()),1,1)
DECLARE @end_date AS DATE = DATEFROMPARTS(YEAR(GETDATE()),7,1)

SELECT 
    CASE WHEN Mercato.desMercato = 'Autovetture' THEN 'PC' 
         WHEN Mercato.desMercato = 'Veicoli Commerciali Leggeri' THEN 'LCV'
         ELSE 'ND' END AS 'CAT_VEICOLO', 
    ISNULL(codCategoriaInternazionale, 'ND') AS 'CAT_INTERNAZIONALE', 
    CAST(codice_vin AS VARCHAR(3)) AS 'WMI', 
    CAST(numero_telaio AS VARCHAR(17)) AS 'VIN', 
    CAST(imm.numero_targa AS VARCHAR) AS 'TARGA', 
    CAST(imm.omologazione_cuc AS VARCHAR) AS 'OMOL/CUC', 
    ISNULL(GruppoMarca.desGruppoMarcaAutomobili, 'ND') AS 'GRUPPO', 
    ISNULL(Marca.desMarca, 'ND') AS 'MARCA', 
    ISNULL(Modello.desModello, 'ND') AS 'MODELLO',
    Allestimento.desAllestimento AS 'ALLESTIMENTO', 
    ISNULL(
	CASE 
		WHEN Mercato.codMercato = '07' THEN Segmento1.desSegmentoGruppoI
		WHEN Mercato.codMercato = '06' THEN 
			CASE WHEN Modello.idSegmentoLCV IS NOT NULL THEN 
					CASE WHEN omm.denominazione_commerciale LIKE '% VAN' THEN 'VAN'
					ELSE LCV.desSegmentoLCV END
				ELSE 'DERIVATA DA AUTOVETTURE' END
			ELSE LCV.desSegmentoLCV
		END, 'ND') AS 'SEGMENTO',	  
    Alimentazione.desAlimentazione AS 'ALIMENTAZIONE', 
    Uso.desTipoUso AS 'USO', 
    Acquistone.desPurchaseGruppoI AS 'MODALITA ACQUISTO', 
    CASE WHEN Titolarità.codTitolarita IS NULL THEN 'PRI' ELSE Titolarità.codTitolarita END AS 'SIGLA TITOLARITA', 
    CASE WHEN Titolarità.desTitolarita IS NULL THEN 'PRIVATO' ELSE Titolarità.desTitolarita END AS 'TITOLARITA', 
    CASE WHEN Comune.codComune IS NULL THEN 'ROMA' ELSE Comune.desComune END AS 'COMUNE', 
    CASE WHEN Comune.codComune IS NULL THEN '058091' ELSE Comune.codComune END AS 'COD. COMUNE', 
    CASE WHEN Comune.codComune IS NULL THEN '00118' ELSE imm.cap_intestatario END AS 'CAP', 
    CASE WHEN Comune.codComune IS NULL THEN 'ROMA' ELSE Provincia.desProvincia END AS 'PROVINCIA', 
    CASE WHEN Comune.codComune IS NULL THEN 'LAZIO' ELSE Regione.desRegione END AS 'REGIONE', 
    CASE WHEN Comune.codComune IS NULL THEN 'CENTRO' ELSE Area.desArea END AS 'AREA', 
    TRY_CAST(omm.tara AS INT) AS 'TARA', 
    TRY_CAST(COALESCE(imm.peso_complessivo_del_veicolo, omm.peso_complessivo) AS INT) AS 'PESO COMPLESSIVO', 
    CASE WHEN Cambio.desCambio IS NULL THEN 'ND' ELSE Cambio.desCambio END AS 'CAMBIO', 
    TRY_CAST(CASE 
        WHEN omm.cilindrata IS NOT NULL AND omm.cilindrata <> 0 THEN LEFT(omm.cilindrata, LEN(omm.cilindrata) - 2)
        ELSE 0
    END AS INT) AS 'CILINDRATA', 
    CAST(omm.valore_potenza_massima AS INT) / 100 AS 'KW MOTORE', 
    CASE WHEN imm.colore IS NULL THEN 'ND' ELSE imm.colore END AS 'COLORE', 
    MONTH(imm.data_immatricolazione_del_veicolo) AS 'MESE_IMM', 
    DAY(imm.data_immatricolazione_del_veicolo) AS 'GIORNO_IMM', 
    YEAR(imm.data_immatricolazione_del_veicolo) AS 'ANNO_IMM', 
    CONCAT(YEAR(imm.data_immatricolazione_del_veicolo), '-', RIGHT(CONCAT('0', MONTH(imm.data_immatricolazione_del_veicolo)), 2)) AS 'ANNO-MESE', 
    CONVERT(varchar, TRY_CAST(imm.data_immatricolazione_del_veicolo AS DATE), 23) AS 'ANNO-MM-GG', 
    (CASE 
		WHEN NULLIF(omm.WLTP_Misto_Calcolato, 0) IS NULL THEN NULLIF(omm.WLTP_Ponderato_Calcolato, 0)
		WHEN NULLIF(omm.WLTP_Ponderato_Calcolato, 0) IS NULL THEN NULLIF(omm.WLTP_Misto_Calcolato, 0)
		ELSE IIF(omm.WLTP_Misto_Calcolato > omm.WLTP_Ponderato_Calcolato, omm.WLTP_Ponderato_Calcolato, omm.WLTP_Misto_Calcolato)
	   END) AS 'CO2 EMIX WLTP',    
    ISNULL(
    CASE WHEN Alimentazione.codAlimentazione = 'ELE' THEN 'ND' ELSE CEE.codRaggSiglaEuropea END, 'ND') AS 'DIRETTIVA CEE'
    FROM dbo.ImmatricolazioniVeicoliLeggeri imm 
INNER JOIN dbo.Omologazioni omm ON omm.codice_omologazione = imm.omologazione_cuc 
LEFT JOIN dbo.Mercato Mercato ON Mercato.codMercato = imm.codMercato 
LEFT JOIN dbo.Marca Marca ON Marca.idMarca = imm.idMarca 
LEFT JOIN dbo.GruppoMarcaAutomobili GruppoMarca ON GruppoMarca.idGruppoMarcaAutomobili = Marca.idGruppoMarcaAutomobili 
LEFT JOIN dbo.Modello Modello ON Modello.idModello = omm.idModello 
LEFT JOIN dbo.Allestimento Allestimento ON Allestimento.codAllestimento = imm.codAllestimento 
LEFT JOIN dbo.Alimentazione Alimentazione ON Alimentazione.codAlimentazione = omm.codAlimentazione 
LEFT JOIN dbo.TipoUso Uso ON Uso.codTipoUso = SUBSTRING(imm.codCategoriaUso, 2, 1) 
LEFT JOIN dbo.Purchase Acquisto ON Acquisto.idPurchase = imm.idPurchase 
LEFT JOIN dbo.PurchaseGruppoI Acquistone ON Acquistone.idPurchaseGruppoI = Acquisto.idPurchaseGruppoI 
LEFT JOIN dbo.Segmento Segmento ON Segmento.codSegmento = imm.codSegmento 
LEFT JOIN dbo.SegmentoGruppoI Segmento1 ON Segmento1.idSegmentoGruppoI = Segmento.idSegmentoGruppoI 
LEFT JOIN dbo.TitolaritaDenominazioneProprietario TitolaritàDenom ON TitolaritàDenom.codDenominazioneProprietario = imm.codice_denominazione_proprietario 
LEFT JOIN dbo.Titolarita Titolarità ON Titolarità.idTitolarita = TitolaritàDenom.idTitolarita 
LEFT JOIN dbo.Comune Comune ON Comune.codComune = imm.codComune 
LEFT JOIN dbo.Provincia Provincia ON Provincia.codProvincia = Comune.codProvincia 
LEFT JOIN dbo.Regione Regione ON Regione.codRegione = Provincia.codRegione 
LEFT JOIN dbo.Area Area ON Area.codArea = Regione.codArea 
LEFT JOIN dbo.TipoCambio Cambio ON Cambio.codCambio = omm.codTipoCambio 
LEFT JOIN dbo.TargheAutoAnnullateReporting AnnullamentiReporting ON AnnullamentiReporting.immatricolazioniId = imm.immatricolazioniId
LEFT JOIN dbo.SegmentoLCV LCV ON LCV.idSegmentoLCV = Modello.idSegmentoLCV
LEFT JOIN dbo.DirettivaCee CEE ON CEE.codDirettivaCee = omm.codDirettivaCee
WHERE imm.data_immatricolazione_del_veicolo >= @start_date
AND imm.data_immatricolazione_del_veicolo < @end_date
AND imm.numero_targa IS NOT NULL
AND CASE WHEN AnnullamentiReporting.immatricolazioniId IS NOT NULL and flannullato = 1 THEN 0 ELSE flannullato END = 0
AND imm.flNuovo = 1
AND omm.codDirettivaCee IS NOT NULL
AND omm.codice_omologazione IS NOT NULL
"""

query_hy2 = """
DECLARE @start_date AS DATE = DATEFROMPARTS(YEAR(GETDATE()),7,1)
DECLARE @end_date AS DATE = DATEFROMPARTS(YEAR(GETDATE())+1,1,1)

SELECT 
    CASE WHEN Mercato.desMercato = 'Autovetture' THEN 'PC' 
         WHEN Mercato.desMercato = 'Veicoli Commerciali Leggeri' THEN 'LCV'
         ELSE 'ND' END AS 'CAT_VEICOLO', 
    ISNULL(codCategoriaInternazionale, 'ND') AS 'CAT_INTERNAZIONALE', 
    CAST(codice_vin AS VARCHAR(3)) AS 'WMI', 
    CAST(numero_telaio AS VARCHAR(17)) AS 'VIN', 
    CAST(imm.numero_targa AS VARCHAR) AS 'TARGA', 
    CAST(imm.omologazione_cuc AS VARCHAR) AS 'OMOL/CUC', 
    ISNULL(GruppoMarca.desGruppoMarcaAutomobili, 'ND') AS 'GRUPPO', 
    ISNULL(Marca.desMarca, 'ND') AS 'MARCA', 
    ISNULL(Modello.desModello, 'ND') AS 'MODELLO',
    Allestimento.desAllestimento AS 'ALLESTIMENTO', 
    ISNULL(
	CASE 
		WHEN Mercato.codMercato = '07' THEN Segmento1.desSegmentoGruppoI
		WHEN Mercato.codMercato = '06' THEN 
			CASE WHEN Modello.idSegmentoLCV IS NOT NULL THEN 
					CASE WHEN omm.denominazione_commerciale LIKE '% VAN' THEN 'VAN'
					ELSE LCV.desSegmentoLCV END
				ELSE 'DERIVATA DA AUTOVETTURE' END
			ELSE LCV.desSegmentoLCV
		END, 'ND') AS 'SEGMENTO',	  
    Alimentazione.desAlimentazione AS 'ALIMENTAZIONE', 
    Uso.desTipoUso AS 'USO', 
    Acquistone.desPurchaseGruppoI AS 'MODALITA ACQUISTO', 
    CASE WHEN Titolarità.codTitolarita IS NULL THEN 'PRI' ELSE Titolarità.codTitolarita END AS 'SIGLA TITOLARITA', 
    CASE WHEN Titolarità.desTitolarita IS NULL THEN 'PRIVATO' ELSE Titolarità.desTitolarita END AS 'TITOLARITA', 
    CASE WHEN Comune.codComune IS NULL THEN 'ROMA' ELSE Comune.desComune END AS 'COMUNE', 
    CASE WHEN Comune.codComune IS NULL THEN '058091' ELSE Comune.codComune END AS 'COD. COMUNE', 
    CASE WHEN Comune.codComune IS NULL THEN '00118' ELSE imm.cap_intestatario END AS 'CAP', 
    CASE WHEN Comune.codComune IS NULL THEN 'ROMA' ELSE Provincia.desProvincia END AS 'PROVINCIA', 
    CASE WHEN Comune.codComune IS NULL THEN 'LAZIO' ELSE Regione.desRegione END AS 'REGIONE', 
    CASE WHEN Comune.codComune IS NULL THEN 'CENTRO' ELSE Area.desArea END AS 'AREA', 
    TRY_CAST(omm.tara AS INT) AS 'TARA', 
    TRY_CAST(COALESCE(imm.peso_complessivo_del_veicolo, omm.peso_complessivo) AS INT) AS 'PESO COMPLESSIVO', 
    CASE WHEN Cambio.desCambio IS NULL THEN 'ND' ELSE Cambio.desCambio END AS 'CAMBIO', 
    TRY_CAST(CASE 
        WHEN omm.cilindrata IS NOT NULL AND omm.cilindrata <> 0 THEN LEFT(omm.cilindrata, LEN(omm.cilindrata) - 2)
        ELSE 0
    END AS INT) AS 'CILINDRATA', 
    CAST(omm.valore_potenza_massima AS INT) / 100 AS 'KW MOTORE', 
    CASE WHEN imm.colore IS NULL THEN 'ND' ELSE imm.colore END AS 'COLORE', 
    MONTH(imm.data_immatricolazione_del_veicolo) AS 'MESE_IMM', 
    DAY(imm.data_immatricolazione_del_veicolo) AS 'GIORNO_IMM', 
    YEAR(imm.data_immatricolazione_del_veicolo) AS 'ANNO_IMM', 
    CONCAT(YEAR(imm.data_immatricolazione_del_veicolo), '-', RIGHT(CONCAT('0', MONTH(imm.data_immatricolazione_del_veicolo)), 2)) AS 'ANNO-MESE', 
    CONVERT(varchar, TRY_CAST(imm.data_immatricolazione_del_veicolo AS DATE), 23) AS 'ANNO-MM-GG', 
    (CASE 
		WHEN NULLIF(omm.WLTP_Misto_Calcolato, 0) IS NULL THEN NULLIF(omm.WLTP_Ponderato_Calcolato, 0)
		WHEN NULLIF(omm.WLTP_Ponderato_Calcolato, 0) IS NULL THEN NULLIF(omm.WLTP_Misto_Calcolato, 0)
		ELSE IIF(omm.WLTP_Misto_Calcolato > omm.WLTP_Ponderato_Calcolato, omm.WLTP_Ponderato_Calcolato, omm.WLTP_Misto_Calcolato)
	   END) AS 'CO2 EMIX WLTP',    
    ISNULL(
    CASE WHEN Alimentazione.codAlimentazione = 'ELE' THEN 'ND' ELSE CEE.codRaggSiglaEuropea END, 'ND') AS 'DIRETTIVA CEE'
    FROM dbo.ImmatricolazioniVeicoliLeggeri imm 
INNER JOIN dbo.Omologazioni omm ON omm.codice_omologazione = imm.omologazione_cuc 
LEFT JOIN dbo.Mercato Mercato ON Mercato.codMercato = imm.codMercato 
LEFT JOIN dbo.Marca Marca ON Marca.idMarca = imm.idMarca 
LEFT JOIN dbo.GruppoMarcaAutomobili GruppoMarca ON GruppoMarca.idGruppoMarcaAutomobili = Marca.idGruppoMarcaAutomobili 
LEFT JOIN dbo.Modello Modello ON Modello.idModello = omm.idModello 
LEFT JOIN dbo.Allestimento Allestimento ON Allestimento.codAllestimento = imm.codAllestimento 
LEFT JOIN dbo.Alimentazione Alimentazione ON Alimentazione.codAlimentazione = omm.codAlimentazione 
LEFT JOIN dbo.TipoUso Uso ON Uso.codTipoUso = SUBSTRING(imm.codCategoriaUso, 2, 1) 
LEFT JOIN dbo.Purchase Acquisto ON Acquisto.idPurchase = imm.idPurchase 
LEFT JOIN dbo.PurchaseGruppoI Acquistone ON Acquistone.idPurchaseGruppoI = Acquisto.idPurchaseGruppoI 
LEFT JOIN dbo.Segmento Segmento ON Segmento.codSegmento = imm.codSegmento 
LEFT JOIN dbo.SegmentoGruppoI Segmento1 ON Segmento1.idSegmentoGruppoI = Segmento.idSegmentoGruppoI 
LEFT JOIN dbo.TitolaritaDenominazioneProprietario TitolaritàDenom ON TitolaritàDenom.codDenominazioneProprietario = imm.codice_denominazione_proprietario 
LEFT JOIN dbo.Titolarita Titolarità ON Titolarità.idTitolarita = TitolaritàDenom.idTitolarita 
LEFT JOIN dbo.Comune Comune ON Comune.codComune = imm.codComune 
LEFT JOIN dbo.Provincia Provincia ON Provincia.codProvincia = Comune.codProvincia 
LEFT JOIN dbo.Regione Regione ON Regione.codRegione = Provincia.codRegione 
LEFT JOIN dbo.Area Area ON Area.codArea = Regione.codArea 
LEFT JOIN dbo.TipoCambio Cambio ON Cambio.codCambio = omm.codTipoCambio 
LEFT JOIN dbo.TargheAutoAnnullateReporting AnnullamentiReporting ON AnnullamentiReporting.immatricolazioniId = imm.immatricolazioniId
LEFT JOIN dbo.SegmentoLCV LCV ON LCV.idSegmentoLCV = Modello.idSegmentoLCV
LEFT JOIN dbo.DirettivaCee CEE ON CEE.codDirettivaCee = omm.codDirettivaCee
WHERE imm.data_immatricolazione_del_veicolo >= @start_date
AND imm.data_immatricolazione_del_veicolo < @end_date
AND imm.numero_targa IS NOT NULL
AND CASE WHEN AnnullamentiReporting.immatricolazioniId IS NOT NULL and flannullato = 1 THEN 0 ELSE flannullato END = 0
AND imm.flNuovo = 1
AND omm.codDirettivaCee IS NOT NULL
AND omm.codice_omologazione IS NOT NULL
"""

query_hy2_previous_year = """
DECLARE @start_date AS DATE = DATEFROMPARTS(YEAR(GETDATE())-1,7,1)
DECLARE @end_date AS DATE = DATEFROMPARTS(YEAR(GETDATE()),1,1)

SELECT 
    CASE WHEN Mercato.desMercato = 'Autovetture' THEN 'PC' 
         WHEN Mercato.desMercato = 'Veicoli Commerciali Leggeri' THEN 'LCV'
         ELSE 'ND' END AS 'CAT_VEICOLO', 
    ISNULL(codCategoriaInternazionale, 'ND') AS 'CAT_INTERNAZIONALE', 
    CAST(codice_vin AS VARCHAR(3)) AS 'WMI', 
    CAST(numero_telaio AS VARCHAR(17)) AS 'VIN', 
    CAST(imm.numero_targa AS VARCHAR) AS 'TARGA', 
    CAST(imm.omologazione_cuc AS VARCHAR) AS 'OMOL/CUC', 
    ISNULL(GruppoMarca.desGruppoMarcaAutomobili, 'ND') AS 'GRUPPO', 
    ISNULL(Marca.desMarca, 'ND') AS 'MARCA', 
    ISNULL(Modello.desModello, 'ND') AS 'MODELLO',
    Allestimento.desAllestimento AS 'ALLESTIMENTO', 
    ISNULL(
	CASE 
		WHEN Mercato.codMercato = '07' THEN Segmento1.desSegmentoGruppoI
		WHEN Mercato.codMercato = '06' THEN 
			CASE WHEN Modello.idSegmentoLCV IS NOT NULL THEN 
					CASE WHEN omm.denominazione_commerciale LIKE '% VAN' THEN 'VAN'
					ELSE LCV.desSegmentoLCV END
				ELSE 'DERIVATA DA AUTOVETTURE' END
			ELSE LCV.desSegmentoLCV
		END, 'ND') AS 'SEGMENTO',	  
    ISNULL(Alimentazione.desAlimentazione, 'ND') AS 'ALIMENTAZIONE', 
    Uso.desTipoUso AS 'USO', 
    Acquistone.desPurchaseGruppoI AS 'MODALITA ACQUISTO', 
    CASE WHEN Titolarità.codTitolarita IS NULL THEN 'PRI' ELSE Titolarità.codTitolarita END AS 'SIGLA TITOLARITA', 
    CASE WHEN Titolarità.desTitolarita IS NULL THEN 'PRIVATO' ELSE Titolarità.desTitolarita END AS 'TITOLARITA', 
    CASE WHEN Comune.codComune IS NULL THEN 'ROMA' ELSE Comune.desComune END AS 'COMUNE', 
    CASE WHEN Comune.codComune IS NULL THEN '058091' ELSE Comune.codComune END AS 'COD. COMUNE', 
    CASE WHEN Comune.codComune IS NULL THEN '00118' ELSE imm.cap_intestatario END AS 'CAP', 
    CASE WHEN Comune.codComune IS NULL THEN 'ROMA' ELSE Provincia.desProvincia END AS 'PROVINCIA', 
    CASE WHEN Comune.codComune IS NULL THEN 'LAZIO' ELSE Regione.desRegione END AS 'REGIONE', 
    CASE WHEN Comune.codComune IS NULL THEN 'CENTRO' ELSE Area.desArea END AS 'AREA', 
    TRY_CAST(omm.tara AS INT) AS 'TARA', 
    TRY_CAST(COALESCE(imm.peso_complessivo_del_veicolo, omm.peso_complessivo) AS INT) AS 'PESO COMPLESSIVO', 
    CASE WHEN Cambio.desCambio IS NULL THEN 'ND' ELSE Cambio.desCambio END AS 'CAMBIO', 
    TRY_CAST(CASE 
        WHEN omm.cilindrata IS NOT NULL AND omm.cilindrata <> 0 THEN LEFT(omm.cilindrata, LEN(omm.cilindrata) - 2)
        ELSE 0
    END AS INT) AS 'CILINDRATA', 
    CAST(omm.valore_potenza_massima AS INT) / 100 AS 'KW MOTORE', 
    CASE WHEN imm.colore IS NULL THEN 'ND' ELSE imm.colore END AS 'COLORE', 
    MONTH(imm.data_immatricolazione_del_veicolo) AS 'MESE_IMM', 
    DAY(imm.data_immatricolazione_del_veicolo) AS 'GIORNO_IMM', 
    YEAR(imm.data_immatricolazione_del_veicolo) AS 'ANNO_IMM', 
    CONCAT(YEAR(imm.data_immatricolazione_del_veicolo), '-', RIGHT(CONCAT('0', MONTH(imm.data_immatricolazione_del_veicolo)), 2)) AS 'ANNO-MESE', 
    CONVERT(varchar, TRY_CAST(imm.data_immatricolazione_del_veicolo AS DATE), 23) AS 'ANNO-MM-GG', 
    (CASE 
		WHEN NULLIF(omm.WLTP_Misto_Calcolato, 0) IS NULL THEN NULLIF(omm.WLTP_Ponderato_Calcolato, 0)
		WHEN NULLIF(omm.WLTP_Ponderato_Calcolato, 0) IS NULL THEN NULLIF(omm.WLTP_Misto_Calcolato, 0)
		ELSE IIF(omm.WLTP_Misto_Calcolato > omm.WLTP_Ponderato_Calcolato, omm.WLTP_Ponderato_Calcolato, omm.WLTP_Misto_Calcolato)
	   END) AS 'CO2 EMIX WLTP',    
    ISNULL(
    CASE WHEN Alimentazione.codAlimentazione = 'ELE' THEN 'ND' ELSE CEE.codRaggSiglaEuropea END, 'ND') AS 'DIRETTIVA CEE'
FROM dbo.ImmatricolazioniVeicoliLeggeri imm 
INNER JOIN dbo.Omologazioni omm ON omm.codice_omologazione = imm.omologazione_cuc 
LEFT JOIN dbo.Mercato Mercato ON Mercato.codMercato = imm.codMercato 
LEFT JOIN dbo.Marca Marca ON Marca.idMarca = imm.idMarca 
LEFT JOIN dbo.GruppoMarcaAutomobili GruppoMarca ON GruppoMarca.idGruppoMarcaAutomobili = Marca.idGruppoMarcaAutomobili 
LEFT JOIN dbo.Modello Modello ON Modello.idModello = omm.idModello 
LEFT JOIN dbo.Allestimento Allestimento ON Allestimento.codAllestimento = imm.codAllestimento 
LEFT JOIN dbo.Alimentazione Alimentazione ON Alimentazione.codAlimentazione = omm.codAlimentazione 
LEFT JOIN dbo.TipoUso Uso ON Uso.codTipoUso = SUBSTRING(imm.codCategoriaUso, 2, 1) 
LEFT JOIN dbo.Purchase Acquisto ON Acquisto.idPurchase = imm.idPurchase 
LEFT JOIN dbo.PurchaseGruppoI Acquistone ON Acquistone.idPurchaseGruppoI = Acquisto.idPurchaseGruppoI 
LEFT JOIN dbo.Segmento Segmento ON Segmento.codSegmento = imm.codSegmento 
LEFT JOIN dbo.SegmentoGruppoI Segmento1 ON Segmento1.idSegmentoGruppoI = Segmento.idSegmentoGruppoI 
LEFT JOIN dbo.TitolaritaDenominazioneProprietario TitolaritàDenom ON TitolaritàDenom.codDenominazioneProprietario = imm.codice_denominazione_proprietario 
LEFT JOIN dbo.Titolarita Titolarità ON Titolarità.idTitolarita = TitolaritàDenom.idTitolarita 
LEFT JOIN dbo.Comune Comune ON Comune.codComune = imm.codComune 
LEFT JOIN dbo.Provincia Provincia ON Provincia.codProvincia = Comune.codProvincia 
LEFT JOIN dbo.Regione Regione ON Regione.codRegione = Provincia.codRegione 
LEFT JOIN dbo.Area Area ON Area.codArea = Regione.codArea 
LEFT JOIN dbo.TipoCambio Cambio ON Cambio.codCambio = omm.codTipoCambio 
LEFT JOIN dbo.TargheAutoAnnullateReporting AnnullamentiReporting ON AnnullamentiReporting.immatricolazioniId = imm.immatricolazioniId
LEFT JOIN dbo.SegmentoLCV LCV ON LCV.idSegmentoLCV = Modello.idSegmentoLCV
LEFT JOIN dbo.DirettivaCee CEE ON CEE.codDirettivaCee = omm.codDirettivaCee
WHERE imm.data_immatricolazione_del_veicolo >= @start_date
AND imm.data_immatricolazione_del_veicolo < @end_date
AND imm.numero_targa IS NOT NULL
AND CASE WHEN AnnullamentiReporting.immatricolazioniId IS NOT NULL and flannullato = 1 THEN 0 ELSE flannullato END = 0
AND imm.flNuovo = 1 
AND omm.codDirettivaCee IS NOT NULL
AND omm.codice_omologazione IS NOT NULL
"""

# Possible new filters for a new ETL rule about new and used vehicles:
#
# AND codDirettivaCee IS NOT NULL AND IS IN ['EURO6D', 'EURO6E', 'EURO6C', 'EURO6B', 'EURO6A', 'EURO6', 'EURO0']
# CASE WHEN anno_prima_immatricolazione IS NULL OR anno_prima_immatricolazione = anno_immatricolazione THEN flNuovo = 1 END

# Date variables to properly select the correct queries



# Connection data
user, password, server, database = get_login_info_from_config()
SERVER = server
DATABASE = database

connection_string = (
    "mssql+pyodbc:///?odbc_connect="
    + urllib.parse.quote_plus(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={SERVER};"
        f"DATABASE={DATABASE};"
        "Authentication=ActiveDirectoryInteractive"
    )
)


# Extract correct data based on current date: if current month is 1 or 2, then data from previous year
# must still be updated.
try:
    auth_thread = Thread(target=simulate_user_login, args=(user, password))
    auth_thread.start()

    engine = create_engine(connection_string)
    with engine.connect() as connection:
        if current_month <= 3:
            hy1 = pd.DataFrame(connection.execute(text(query_hy1)).fetchall())
            hy2 = pd.DataFrame(connection.execute(text(query_hy2_previous_year)).fetchall())
        elif 4 <= current_month < 7:
            hy1 = pd.DataFrame(connection.execute(text(query_hy1)).fetchall())
            hy2 = pd.DataFrame()
        elif 10 <= current_month <= 12:
            hy2 = pd.DataFrame(connection.execute(text(query_hy2)).fetchall())
            hy1 = pd.DataFrame()
        else:
            hy1 = pd.DataFrame(connection.execute(text(query_hy1)).fetchall())
            hy2 = pd.DataFrame(connection.execute(text(query_hy2)).fetchall())
            
            
        # # Test for full year
        # hy1 = pd.DataFrame(connection.execute(text(query_hy1)).fetchall())
        # # hy1.columns = column_names

        # hy2 = pd.DataFrame(connection.execute(text(query_hy2)).fetchall())
        # # hy2.columns = column_names   
        
        # Save as CSV and then convert to .xlsx to maximise performance
        output_folder = r"L:\03.Articoli_Analisi_(exUtenti)\Day-by-day"
        if not hy1.empty:
            verify_df_pairs(hy1)
            hy1.to_csv(f"{current_year}_HY1.csv", index=False, encoding='utf-8')
            convert_csv_to_xlsx(f"{current_year}_HY1.csv", f"{current_year}_HY1.xlsx", output_folder)

        if not hy2.empty:
            hy2_filename = f"{previous_year}_HY2.xlsx" if current_month in [1, 2, 3] else f"{current_year}_HY2.xlsx"
            hy2 = verify_df_pairs(hy2)
            hy2.to_csv(f"{previous_year}_HY2.csv" if current_month in [1, 2, 3] else f"{current_year}_HY2.csv", index=False, encoding='utf-8')            
            convert_csv_to_xlsx(f"{previous_year}_HY2.csv" if current_month in [1, 2, 3] else f"{current_year}_HY2.csv", hy2_filename, output_folder)

except OperationalError as e:
    print("Errore di connessione:", e)



end_time = time.time()
elapsed_time = end_time - start_time
print(f"Tempo totale: {elapsed_time:.2f} secondi ({elapsed_time / 60:.2f} minuti). Brum brum!")