EasyTab Add-in per Microsoft Excel
====================================

<p align="center">
  <img src="docs/Icona.png" alt="Logo" width="150"/>
</p>

<p align="center">
================== <b>A cosa serve</b> ==================
</p>


EasyTab è un **Add-in per Microsoft Excel** che estende le funzionalità di Excel integrando Python.  
Ti permette di:

- **Leggere file in formato logico** e trasformarli automaticamente in tabelle su Excel.
- **Convertire tabelle Excel** in file scritti in sintassi logica.
- **Risolvere programmi logici (ASP)** e consultare i risultati direttamente nei fogli di lavoro.

<p align="center">
================== <b>Installazione</b> ==================
</p>

1. Scaricare l’ultima versione dell’installer da [Release](../../releases/latest).  
2. Avviare `EasyTab_Installer.exe` e attendere il completamento delle operazioni.  
3. Al termine, l’Add-in verrà copiato nella cartella predefinita: `[LetteraUnità]:\Program Files\EasyTab`
(ad es. C:\Program Files\EasyTab se Windows è stato installato sul disco C:)

<p align="center">
================== <b>Disinstallazione</b> ==================
</p>

1. Disattivare l’Add-in da Excel:  
- Aprire Excel → File → Opzioni → Componenti aggiuntivi → [Vai...]  
- Togliere la spunta su “EasyTab” e confermare con **OK**  
2. Aprire il **Pannello di Controllo → Programmi → Disinstalla un programma**  
3. Selezionare dalla lista “EasyTab Add-in per Excel” e procedere con la disinstallazione.

<p align="center">
================== <b>Come attivare l'Add-in</b> ==================
</p>

1. Aprire Microsoft Excel.  
2. Selezionare: `File → Opzioni → Componenti aggiuntivi`.  
3. In basso, dove è presente la voce “Gestisci”, selezionare **Componenti aggiuntivi di Excel** e cliccare su **Vai...**.  
4. Selezionare `EasyTab` dalla lista, oppure se non è presente cliccare su **Sfoglia…** e cercare: `[LetteraUnità]:\Program Files\EasyTab\EasyTab.xlam`
5. Assicurarsi che la casella sia spuntata, quindi cliccare su **OK**.

A volte è possibile che nonostante l'Add-in sia stato caricato Excel blocchi l'esecuzione delle macro. Per risolvere questo problema seguire i seguenti passaggi:

## Abilitare le macro

1. Aprire Microsoft Excel.  
2. Selezionare **File → Opzioni → Centro protezione → Impostazioni Centro protezione**.  
3. Cliccare su **Impostazioni delle macro**.  
4. Selezionare “Attivare tutte le macro”.  
5. Confermare con **OK**.

================== **Come utilizzare l'Add-in** ==================

### 1️⃣ Lettura file

- Utilizzare file con estensione `.ini`, `.cfg`, `.json`, `.asp`, `.lp`, `.txt`, `.pl`. I dati devono essere formattati nel seguente modo:

```prolog
xls_input(nome_foglio(parametri)).
```

Esempio:

```prolog
xls_input(pazienti("nome", "eta", "sesso")).
```

- Mentre per quanto riguarda la definizione dei tipi di dato bisogna seguire la seguente sintassi:

```prolog
xls_input_types(nome_foglio(parametro, tipo)).
```

- I tipi disponibili non sono case sensitive e sono:

1. integer: accetta valori interi (-999999 … 999999)

2. decimal: accetta numeri reali (-1e307 … 1e307)

3. string: accetta qualsiasi stringa

4. multivalued: specificare i valori ammessi ripetendo la scrittura

Se non specificato, il tipo predefinito è string.

Esempio:

```prolog
xls_input_types(pazienti("eta", "integer")).
xls_input_types(pazienti("peso", "decimal")).
xls_input_types(pazienti("nome", "string")).
xls_input_types(pazienti("sesso", "maschio")).
xls_input_types(pazienti("sesso", "femmina")).
```

### 2️⃣ Convertitore

- Non ha regole particolari di utilizzo.

### 3️⃣ Risoluzione ASP

- Utilizzare file .ini, .cfg, .json, .asp, .lp, .txt, .pl. All'interno di questi file è necessario che ci sia la direttiva:

```
#show
```
per standardizzare il formato della risposta.

Esempio:

```prolog
xls_output("pazienti").

% pazienti(nome, eta).

pazienti("Luigi", 23). 
pazienti("Giovanni", 9).
pazienti("Marco", 45).
pazienti("Luca", 3).
pazienti("Giovanna", 32).
pazienti("Lucia", 9).

minorenne(Nome) :- pazienti(Nome, Eta), Eta < 18.

#show minorenne/1.
```

Questo genererà un foglio minorenne con una tabella contenente: Giovanni, Luca, Lucia.

<p align="center">
================== <b>Risoluzione dei Problemi</b> ==================
</p>

1. Excel non trova l’Add-in
→ Verificare il percorso corretto nella cartella:
    C:\Users\<nomeUtente>\AppData\Roaming\Microsoft\AddIns

2. Le macro non funzionano
→ Controllare che le macro siano abilitate (vedi sezione Abilitare le macro).

<p align="center">
================== <b>Supporto</b> ==================
</p>

Segnalare bug o proporre migliorie tramite Issue su GitHub o scrivere una mail al seguente indirizzo: <b> marcello.luciani2@gmail.com

====================================

Nota: EasyTab è stato sviluppato per scopi accademici e potrebbe non essere compatibile con tutte le versioni di Microsoft Excel.