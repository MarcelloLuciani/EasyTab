import clingo
import os
import pywintypes
import win32com.client

from sys import exit, argv
from pathlib import Path

# Exit code:
    # 0: corretta esecuzione
    # 2: errore excel
    # 3: nessuna soluzione possibile
    # 4: nessuna riga show trovata
    # 99: altri errori

def make_on_model(mappa):
    def on_model(model):
        elementi = model.symbols(shown=True)
        for elemento in elementi:
            nome = elemento.name
            if nome not in mappa:
                mappa[nome] = []
            valori = [str(a).replace('"', '') for a in elemento.arguments]
            mappa[nome].append(valori)
    return on_model

def indice_colonna_to_lettera(n):
    lettere = ''
    while n > 0:
        n, resto = divmod(n - 1, 26)
        lettere = chr(65 + resto) + lettere
    return lettere


cartella = Path.home() / "Documents" / "EasyPlan"
cartella.mkdir(parents=True, exist_ok=True)

lpFiles = argv[1:]

mappaRisultati = {}

soluzioneTrovata = False
showTrovato = False

# Creo un file temporaneo che conterrà il contenuto di tutti i file passati come
# paramentro
percorsoFileTemp = str(cartella) + "\\temp.lp"
fileTemp = open(percorsoFileTemp, "w")

# Leggo tutti i file passati come parametro e li aggiungo al file temporaneo

for file in lpFiles:
    with open(file, "r") as f:
        while line := f.readline():
            fileTemp.write(line)
            if line.strip().startswith("#show"):
                showTrovato = True
    fileTemp.write("\n\n")



fileTemp.close()

if not showTrovato:
    print("Nessuna riga nel formato \"#show\" trovata! Correggere e avviare nuovamente la funzione!")
    input("Premere invio per continuare...")
    exit("4")

file = os.path.abspath(percorsoFileTemp)  # percorso assoluto

# Clingo
ctl = clingo.Control()
ctl.load(file)
ctl.ground([("base", [])])
risultato = ctl.solve(on_model=make_on_model(mappaRisultati))

soluzioneTrovata = risultato.satisfiable or soluzioneTrovata

if os.path.exists(percorsoFileTemp):
    os.remove(percorsoFileTemp)

if (soluzioneTrovata):
    
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        wb = excel.ActiveWorkbook

        # Disattivo temporaneamente gli avvisi di Excel in modo
        # da cancellare i file già esistenti senza dover interagire
        # con l'utente
        excel.DisplayAlerts = False


        for chiave in mappaRisultati:

            for sheet in wb.Sheets:
                if sheet.Name == chiave:
                    sheet.Delete()
                    break

            nuovoFoglio = wb.Sheets.Add()
            nuovoFoglio.Name = chiave
            nuovoFoglio.Cells.Clear()
            ws = nuovoFoglio

            dati = mappaRisultati[chiave]
            numeroRighe = len(dati)
            numeroColonne = len(dati[0]) if numeroRighe > 0 else 0
            

            # Scrittura dati
            for riga_idx, riga in enumerate(dati):
                for col_idx, valore in enumerate(riga):
                    ws.Cells(riga_idx + 3, col_idx + 2).Value = valore

            # Definizione range tabella
            col_iniziale = 2
            col_fine = col_iniziale + numeroColonne - 1
            prima_colonna_lettera = indice_colonna_to_lettera(col_iniziale)
            ultima_colonna_lettera = indice_colonna_to_lettera(col_fine)
            ultima_riga = 2 + numeroRighe
            table_range = f"{prima_colonna_lettera}2:{ultima_colonna_lettera}{ultima_riga}"

            # Creazione tabella
            lista = ws.ListObjects.Add(1, ws.Range(table_range), 0, 1)
            lista.Name = f"Tabella_{chiave}"
            lista.TableStyle = "TableStyleMedium9"

        # Riattivo gli avvisi
        excel.DisplayAlerts = True
        print("Esecuzione completata!")
        input("Premere invio per continuare...")
        exit("0")

    except pywintypes.com_error as e:
        
        # Riattivo gli avvisi
        excel.DisplayAlerts = True

        print("Errore COM:", e)
        input("Premere invio per continuare...")
        exit("2")


    except Exception as ex:
        
        # Riattivo gli avvisi
        excel.DisplayAlerts = True
        
        template = "An exception of type {0} occurred. Arguments:\n{1!r}"
        message = template.format(type(ex).__name__, ex.args)
        print(message)
        input("Premere invio per continuare...")
        exit("99")

else:
    print("Nessuna soluzione trovata!")
    input("Premere invio per continuare...")
    exit("3")