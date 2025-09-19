import win32com.client
import pythoncom

from pathlib import Path
from sys import exit
from datetime import date

    #
    # Per ogni Foglio devo creare il codice corrispondente in formato clingo e salvarlo su un file di uscita:
    # xls_output(nomeFoglio).
    # xls_output_param(nomeFoglio, "Parametro", valore). 
    #
    # Esempio:
    # xls_output("pazienti").
    # xls_output_param("pazienti", .... ) -> paziente(nome, eta, sesso).
    #

try:

    pythoncom.CoInitialize()

    # Percorso della cartella Documenti > DoctorPlan
    cartella = Path.home() / "Documents" / "EasyTab"
    cartella.mkdir(parents=True, exist_ok=True)

    excel = win32com.client.GetActiveObject("Excel.Application")
    wb = excel.ActiveWorkbook

    lista = []

    # Recupero i nomi di tutti i fogli disponibili nel Foglio Excel
    for foglio in wb.Sheets:
        # Ignoro eventuali fogli nascosti
        if foglio.Visible == -1:
            print(foglio.Name)
            # Non aggiungo fogli che nel nome contengono "Foglio"
            if "Foglio" not in foglio.Name:
                lista.append(foglio.name)

    
    for chiave in lista:
        
        foglio_selezionato = wb.Sheets(str(chiave))
        
        percorso_completo_file = str(cartella) + "\\" + str(chiave) + ".lp"

        with open(percorso_completo_file, "w") as f:

            f.write("xls_output(\"" + str(chiave) + "\").\n\n")

            for tabella in foglio_selezionato.ListObjects:
                intervallo = tabella.Range
                dati = intervallo.Value
                lista_dati = [list(riga) for riga in dati]
                
                intestazione = lista_dati[0]
                
                stringa_commento = "% " + str(chiave) + "( " + ", ".join(str(v) for v in intestazione) + " ).\n\n"

                f.write(stringa_commento)

                for riga in lista_dati[1:]:
                    
                    stringa = str(chiave) + "( "

                    if all(cell is None for cell in riga):
                        continue  # Salta righe completamente vuote

                    valoriFormattati = []

                    for valore in riga:
                        if valore is None:
                            valoriFormattati.append("none")
                        elif isinstance(valore, float):
                            if valore.is_integer():
                                valoriFormattati.append(str(int(valore)))
                            else:
                                valoriFormattati.append(str(valore).replace(",", "."))
                        elif isinstance(valore, date):
                            # converto solo la parte data in formato ISO YYYY-MM-DD
                            data = "\"" + valore.strftime("%Y-%m-%d") + "\""
                            valoriFormattati.append(data)
                        else: 
                            valoriFormattati.append(f'"{valore}"')

                    stringa = f'{chiave}(' + ", ".join(valoriFormattati) + ').'

                    f.write(stringa)
                    f.write("\n")
                    
    pythoncom.CoUninitialize()
    print("File Excel convertito correttamente!")
    print("I file sono disponibili nella cartella Documenti.")
    input("Premere invio per continuare...")
    exit(0)

except Exception as ex:
    
    pythoncom.CoUninitialize()

    template = f"Eccezione di tipo {type(ex).__name__}. Argomenti: {ex.args}."
    print(template) 
    print("Premere invio per continuare...")
    input()
    exit(99)
    