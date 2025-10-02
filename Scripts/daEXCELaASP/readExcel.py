import win32com.client
import pythoncom
from pathlib import Path
from sys import exit
from datetime import date
from decimal import Decimal, ROUND_HALF_UP

try:
    pythoncom.CoInitialize()

    cartella = Path.home() / "Documents" / "EasyTab"
    cartella.mkdir(parents=True, exist_ok=True)

    excel = win32com.client.GetActiveObject("Excel.Application")
    wb = excel.ActiveWorkbook

    lista_fogli = []

    for foglio in wb.Sheets:
        if foglio.Visible == -1 and "Foglio" not in foglio.Name:
            lista_fogli.append((foglio.Name).lower())

    for nome_foglio in lista_fogli:
        foglio_selezionato = wb.Sheets(nome_foglio)
        percorso_file = str(cartella / f"{nome_foglio}.lp")

        with open(percorso_file, "w") as f:
            f.write(f'xls_output("{nome_foglio}").\n\n')

            for tabella in foglio_selezionato.ListObjects:
                intervallo = tabella.Range
                dati = intervallo.Value
                listaDati = [list(riga) for riga in dati]

                intestazione = listaDati[0]
                stringa_commento = "% " + str(nome_foglio) + "( " + ", ".join(str(v) for v in intestazione) + " ).\n\n"
                f.write(stringa_commento)

                colonneDecimali = []
                for indice, col in enumerate(tabella.ListColumns, start=1):
                    try:
                        if col.DataBodyRange.Validation.Type == 2:  # xlValidateDecimal
                            colonneDecimali.append(indice - 1)  
                    except:
                        pass
                
                fattoriColonna = {}
                for indice in colonneDecimali:
                    max_dec = 0
                    
                    for riga in listaDati[1:]:
                        if riga[indice] is not None:
                            valore = Decimal(str(riga[indice]))
                            
                            # Calcolo delle cifre decimali
                            if valore % 1 != 0:
                                numDecimali = -(valore.as_tuple().exponent)
                                if numDecimali > max_dec:
                                    max_dec = numDecimali
                    
                    # Determino il fattore di conversione
                    if max_dec > 0:
                        fattoriColonna[indice] = 10 ** max_dec
                    else:
                        fattoriColonna[indice] = 1

                # Scrivo i dati formattati
                for riga in listaDati[1:]:
                    if all(cell is None for cell in riga):
                        continue

                    valoriFormattati = []
                    for indice, valore in enumerate(riga):
                        if valore is None:
                            valoriFormattati.append("none")
                        elif isinstance(valore, (float, Decimal)):
                            val = Decimal(str(valore))
                            if indice in fattoriColonna:
                                val *= fattoriColonna[indice]
                            intero_val = int(val.to_integral_value(rounding=ROUND_HALF_UP))
                            valoriFormattati.append(str(intero_val))
                        else:
                            valoriFormattati.append(f'"{valore}"')

                    stringa = f'{nome_foglio}({", ".join(valoriFormattati)}).'
                    f.write(stringa + "\n")

                # Commenti sui fattori di conversione
                for indice, fattore in fattoriColonna.items():
                    f.write(f"% Colonna {intestazione[indice]} (campo {indice+1}): fattore di conversione {fattore}\n")
                f.write("\n")

    pythoncom.CoUninitialize()
    print("File Excel convertito correttamente!")
    print("I file sono disponibili nella cartella Documenti.")
    input("Premere invio per continuare...")
    exit(0)

except Exception as ex:
    pythoncom.CoUninitialize()
    print(f"Eccezione di tipo {type(ex).__name__}. Argomenti: {ex.args}.")
    input("Premere invio per continuare...")
    exit(99)

