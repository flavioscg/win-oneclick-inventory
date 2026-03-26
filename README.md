# Win One‑Click Inventory

Script PowerShell **one‑file** e **portabile** per inventariare rapidamente un PC Windows.  
Mostra a video e salva su file **CPU, RAM, dischi, versione/build Windows (con rilevamento affidabile di Windows 11), Office (build/canale), profili Windows, account Office/Teams/WAM, software installato (Uninstall) e, opzionalmente, app MSIX/Store**.

> Output: crea `NOME-log.txt` accanto allo script (utile su chiavetta USB).  
> Opzionale: CSV con la lista software.

---

## Contenuto del repo

```
win-oneclick-inventory/
├─ inventory.ps1
└─ README.md
```

> Il file principale è **`inventory.ps1`**. Nessuna dipendenza esterna.

---

## Caratteristiche

- **Hardware**
  - CPU (nome / cores / threads / max MHz)
  - RAM totale (GB)
  - Dischi fisici (modello + capacità in GB)
  - Volumi/partizioni (size/free in GB)

- **Windows**
  - Edizione (es. *Windows 11 Pro*), EditionID
  - **Rilevamento affidabile di Windows 11 tramite build**: `CurrentBuild >= 22000` ⇒ Windows 11  
    (utile se `ProductName` resta “Windows 10 …” dopo upgrade)
  - Feature release (es. *24H2*)
  - Build completa (es. *26100.4946* da `CurrentBuild.UBR`)

- **Office**
  - Click‑to‑Run (Microsoft 365 Apps) o MSI/legacy
  - Versione/build, canale/SKU, architettura (x64/x86)

- **Account**
  - Profili Windows presenti (User, SID, LastUse, Loaded)
  - **Office** (dal registry HKU\…\Office.0\Common\Identity\Identities) — funziona anche se lo script è eseguito elevato
  - **Teams** (best‑effort): classico + nuovo (MSIX), ricerca e‑mail nei JSON/log
  - **WAM/AAD**: account moderni (UPN/e‑mail) da registry (IdentityCRL + AAD Cache)

- **Software installato**
  - Lettura da **Registry Uninstall**: HKLM (x64/x86) + HKCU  
    (no `Win32_Product`: è lento e può attivare “self‑heal” MSI)
  - **App MSIX/Store** opzionali con `Get-AppxPackage`

- **Output**
  - Chiede un nome e salva `NOME-log.txt` **nella stessa cartella** dello script
  - (Opzionale) CSV con elenco software/app Store

- **Compatibilità**
  - Windows 10/11
  - PowerShell **5.1** (preinstallato) e PowerShell 7

---

## Requisiti

- Windows 10 o Windows 11
- PowerShell 5.1 (o PowerShell 7)
- Nessun modulo aggiuntivo
- Per vedere **App Store di tutti gli utenti**: avvia PowerShell **come Amministratore**

---

## Esecuzione rapida

### Metodo consigliato (per‑process, non modifica il sistema)
```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\inventory.ps1
```

Oppure: tasto destro sul file → **Esegui con PowerShell**.  
Alla partenza chiede un nome (es. `MarioRossi`) e genera `MarioRossi-log.txt`.

> Se il file è “bloccato” (proviene da Internet), sbloccalo:
> ```powershell
> Unblock-File .\inventory.ps1
> ```

---

## Execution Policy (quando gli script sono bloccati)

Windows può impedire l’esecuzione di script non firmati. Hai più opzioni:

- **Solo per questa sessione** (temporaneo, sicuro):
  ```powershell
  Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
  .\inventory.ps1
  ```
  Chiusa la console, la policy torna com’era.

- **Per l’utente corrente** (consigliato se lo usi spesso):
  ```powershell
  Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
  .\inventory.ps1

  # Dopo l’uso, per ripristinare:
  Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Restricted
  ```

- **One‑liner, senza toccare la policy globale**:
  ```powershell
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\inventory.ps1
  ```

> Evita `Unrestricted` se non necessario. Se lo usi, **ripristina** a `Restricted` o `RemoteSigned` al termine.

---

## Opzioni (nel file `inventory.ps1`)

In testa alla sezione **Collect** puoi attivare/disattivare:

```powershell
$IncludeSystemComponents = $false  # true: include componenti di sistema
$IncludeUpdates          = $false  # true: include Update/Hotfix nell’elenco software
$IncludeStoreApps        = $false  # true: include App MSIX/Store (Get-AppxPackage)
$ExportCSV               = $false  # true: esporta CSV accanto al TXT
```

- Se abiliti `$IncludeStoreApps`, con PowerShell **elevato** puoi usare `-AllUsers` per elencare le app Store di tutti gli utenti.

---

## Esempio di output (estratto)


## Troubleshooting

- **Windows 11 mostrato come Windows 10**
  - Lo script usa la **build** (`CurrentBuild >= 22000`) per determinare la famiglia; se vedi incongruenze, assicurati di eseguire **PowerShell 64 bit** (`C:\Windows\System32\WindowsPowerShell
1.0\powershell.exe`) e non la versione 32 bit.

- **Nessun account Office/Teams**
  - Esegui lo script **come l’utente interessato** e assicurati che il profilo sia **Loaded** (login interattivo).
  - Apri Office/Teams, conferma il login, **chiudi e riapri** l’app, poi rilancia lo script.
  - Le versioni moderne possono non esporre e‑mail in chiaro nei file/registry classici; la sezione **WAM/AAD** cerca di coprirlo.

- **App Store non complete**
  - Avvia PowerShell **come Amministratore** e abilita `$IncludeStoreApps = $true`.

- **“L’esecuzione di script è disabilitata”**
  - Vedi la sezione **Execution Policy** (usa `Process Bypass` o `RemoteSigned` per CurrentUser e poi ripristina).

---

## Sicurezza e privacy

- Il report può contenere **UPN/e‑mail** (account aziendali/personali). Trattalo come **confidenziale**.
- Nessun dato viene inviato in rete: tutto resta **in locale**.
- Se alleghi log a ticket o repo, considera di **anonimizzare** le e‑mail.

---