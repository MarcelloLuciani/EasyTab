import win32com.client
import pythoncom
import os

from sys import exit, argv

def calc_correct(lista1, lista2, index):
    lunghezza = 1
    for x in range(index-2,-1,-1):
        lunghezza += len(lista1[lista2[x]][0])
    return lunghezza

def indice_colonna_to_lettera(n):
    #
    # 1 -> A
    # 27 -> AA
    # 
    lettere = ""
    while n > 0:
        n, resto = divmod(n - 1, 26)
        lettere = chr(65 + resto) + lettere
    return lettere

status_file = os.path.join(os.environ["TEMP"], "")

try:
    
    pythoncom.CoInitialize()

    excel = win32com.client.GetActiveObject("Excel.Application")
    wb = excel.ActiveWorkbook
    if wb == None:
        raise Exception()
    

except Exception as ex:
    print("Errore nell'apertura del file excel")
    input("Premi invio per continuare...")
    exit(1)


try:

    fogli = []
    vincoli = []

    for filepath in argv[1:]:
        with open(filepath, "r", encoding="utf-8") as f:
            for riga in f:
                riga = riga.strip()
                if not riga or riga[0] == "%" or riga == "\n":
                    continue

                # Controllo con match preciso
                if riga.startswith("xls_input_types(") and riga.endswith("."):
                    vincoli.append(riga)
                elif riga.startswith("xls_input(") and riga.endswith("."):
                    fogli.append(riga)
                else:
                    raise Exception("Errore Riga ", f'Riga non conforme: \"{riga.replace("\n", "")}\" nel file {filepath}')

except Exception as ex:
    
    print("! ".join(ex.args))
    input("Premi invio per continuare...")
    exit(2)

# Creazione del dizionario che conterrà i nomi dei vari fogli, le intestazioni delle tabelle ed eventuali vincoli
fogliVincoliDict = {}

# Manipolazione stringhe Fogli e Intestazioni
for foglio in fogli:
    rigaModificata = foglio[10:len(foglio)-3]

    nomeFoglio = rigaModificata.split("(")[0]
    intestazioni = rigaModificata.split("(")[1].split(",")

    for x in range(len(intestazioni)):
        intestazioni[x] = intestazioni[x].strip().replace("\"","")
        
    if nomeFoglio not in fogliVincoliDict:
        fogliVincoliDict[nomeFoglio] = {}

    for intestazione in intestazioni:
        if intestazione not in fogliVincoliDict[nomeFoglio]:
            fogliVincoliDict[nomeFoglio][intestazione] = []

try:

    # Manipolazione stringhe per i Vincoli
    for vincolo in vincoli:
        rigaModificata = vincolo[16:len(vincolo)-3]
        
        nomeFoglio = rigaModificata.split("(")[0]
        
        listaVincoli = rigaModificata.split("(")[1].split(",")

        for x in range(len(listaVincoli)):
            listaVincoli[x] = listaVincoli[x].strip().replace("\"","")
        
        if nomeFoglio in fogliVincoliDict:
            if listaVincoli[0] in fogliVincoliDict[nomeFoglio]:
                fogliVincoliDict[nomeFoglio][listaVincoli[0]].append(listaVincoli[1])
        else:
            raise Exception("Probabile errore di battitura! ", f'Vincolo non riconosciuto: \"{vincolo.replace("\n", "")}\"')

except Exception as ex:
    
    print("! ".join(ex.args))
    input("Premi invio per continuare...")
    exit(3)


# Lista per le validazioni da inserire nel foglio "Liste"
liste_per_intestazione = {}
for foglio, colonne in fogliVincoliDict.items():
    for intestazione, vincoli in colonne.items():
        if vincoli:
            chiave_lista = (foglio, intestazione)
            if chiave_lista not in liste_per_intestazione:
                liste_per_intestazione[chiave_lista] = []
            liste_per_intestazione[chiave_lista].append(vincoli)

# Creazione del foglio Liste con tutti i valori per i drop-down
# Cancella foglio Liste se esiste

    # Disattivo temporaneamente gli avvisi di Excel in modo
    # da cancellare i file già esistenti senza dover interagire
    # con l'utente
    excel.DisplayAlerts = False

try:
    ws_liste = wb.Sheets("Liste")
    ws_liste.Delete()
except:
    pass

# Crea nuovo foglio Liste
ws_liste = wb.Sheets.Add()
ws_liste.Name = "Liste"
ws_liste.Visible = False
ws_liste.Cells.Clear()


try:


    lista_chiavi = list(liste_per_intestazione.keys())
    cols = 1
    for x in range(len(lista_chiavi)):
        nome_foglio = lista_chiavi[x][0]
        intestazione = lista_chiavi[x][1]
        
        for y in range(len(liste_per_intestazione[lista_chiavi[x]][0])):
            ws_liste.Cells(1, cols).Value = f"{nome_foglio}_{intestazione}"
            ws_liste.Cells(2, cols).Value = liste_per_intestazione[lista_chiavi[x]][0][y]
            cols += 1

    

    # Creazione dei vari fogli nel documento Excel
    for nomeFoglio in fogliVincoliDict:

        # Se il foglio esiste già, lo cancello
        for sheet in wb.Sheets:
            if sheet.Name == nomeFoglio:
                sheet.Delete()
                break

        # Creo un nuovo foglio con quel nome
        ws = wb.Sheets.Add()
        ws.Name = nomeFoglio


        # Inizio la costruzione della Tabella
        
        headers = list(fogliVincoliDict[nomeFoglio].keys())

        for x in range(len(headers)):
            ws.Cells(2, x+2).Value = headers[x]
            cell = ws.Cells(2, x+2)
            cell.Font.Bold = True
            cell.Font.Color = -1
            cell.Interior.Color = 15132390
            cell.HorizontalAlignment = -4108
            cell.VerticalAlignment = -4108

        num_colonne = len(headers)
        colonna_iniziale = 2
        colonna_fine = colonna_iniziale + len(headers) - 1

        prima_colonna_lettera = indice_colonna_to_lettera(colonna_iniziale)
        ultima_colonna_lettera = indice_colonna_to_lettera(colonna_fine)

        riga_iniziale = 2
        ultima_riga = riga_iniziale + 10

        table_range = f"{prima_colonna_lettera}{colonna_iniziale}:{ultima_colonna_lettera}{ultima_riga}"

        lista = ws.ListObjects.Add(1, ws.Range(table_range), 0, 1)
        lista.Name = f"Tabella_{nomeFoglio}"
        lista.TableStyle = "TableStyleMedium9"

        for i in range(num_colonne):
            lettera_colonna = indice_colonna_to_lettera(colonna_iniziale + i)
            ws.Columns(lettera_colonna).ColumnWidth = 20
        
        for idx in range(len(headers)):
            intestazione = headers[idx]
            
            vincoli_colonna = fogliVincoliDict[nomeFoglio][intestazione]
            vincoli_colonna_lower = str(vincoli_colonna).lower().translate(str.maketrans("", "", "'[] "))
            col_index_excel = colonna_iniziale + idx
            lettera_colonna = indice_colonna_to_lettera(col_index_excel)
            range_validazione = ws.Range(f"{lettera_colonna}{riga_iniziale+1}:{lettera_colonna}{ultima_riga}")

            range_validazione.Validation.Delete()           #Per sicurezza

            chiave = (nomeFoglio, intestazione)

            
            if vincoli_colonna_lower == "" or vincoli_colonna_lower == "string":
                #print("Sono una stringa")
                continue

            elif vincoli_colonna_lower == "integer":
                #print("Sono integer")
                try:
                    range_validazione.Validation.Add(
                        Type=1,       # xlValidateWholeNumber
                        AlertStyle=1, # xlValidAlertStop
                        Operator=1,   # xlBetween
                        Formula1="-999999",
                        Formula2="999999"
                    )
                    range_validazione.Validation.ErrorTitle = "Numero richiesto"
                    range_validazione.Validation.ErrorMessage = "Inserisci un numero intero valido"
                except:
                    pythoncom.CoUninitialize()
                    print("Errore primo vincolo")
                    input("Premere invio per continuare...")
                    exit(3)
            
            elif vincoli_colonna_lower == "decimal":
                #print("Sono decimal")
                try:
                    # Numeri decimali
                    range_validazione.Validation.Add(
                    Type=2,        # xlValidateDecimal
                    AlertStyle=1,  # xlValidAlertStop
                    Operator=1,    # xlBetween
                    Formula1=-1e307,
                    Formula2= 1e307
                    )
                    range_validazione.Validation.ErrorTitle = "Numero decimale richiesto"
                    range_validazione.Validation.ErrorMessage = "Inserisci un numero decimale valido"
                
                except Exception as ex:
                    pythoncom.CoUninitialize()
                    print("Errore secondo vincolo")
                    input("Premere invio per continuare...")
                    exit(3)
            
            else:
                #print("Sono multivalore")
                try:
                    if chiave in liste_per_intestazione:
                        col_liste_idx = calc_correct(liste_per_intestazione, lista_chiavi, lista_chiavi.index(chiave) + 1)
                        num_valori = len(vincoli_colonna)
                        col_iniziale = indice_colonna_to_lettera(col_liste_idx)
                        col_finale = indice_colonna_to_lettera(col_liste_idx+num_valori-1)
                        formula = f"=Liste!${col_iniziale}$2:${col_finale}$2"
                        range_validazione.Validation.Add(
                            Type=3,
                            AlertStyle=1,
                            Operator=3,
                            Formula1=formula
                        )
                        range_validazione.Validation.ErrorTitle = "Valore non valido"
                        range_validazione.Validation.ErrorMessage = f"Seleziona un valore valido dalla lista"
                except:
                    pythoncom.CoUninitialize()
                    print("Errore terzo vincolo")
                    input("Premere invio per continuare...")
                    exit(3)

    # Riattivo gli avvisi
    excel.DisplayAlerts = True
    pythoncom.CoUninitialize()
    print("Esecuzione completata!")
    input("Premere invio per continuare...")
    exit("0")

except Exception as ex:

    # Riattivo gli avvisi
    excel.DisplayAlerts = True
    
    pythoncom.CoUninitialize()

    template = f"Eccezione di tipo {type(ex).__name__}. Argomenti: {ex.args}."
    print(template)
    print("Premere invio per continuare...")
    input(99)

    
