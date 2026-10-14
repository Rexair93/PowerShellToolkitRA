# Linee Guida PowerShell (CREDIT-SAVING MODE)

## 1. Regola d'oro: Solo Codice Puro

- **Output essenziale:** Fornisci SOLO il codice PowerShell richiesto.

- **Nessuna spiegazione:** Non aggiungere introduzioni, conclusioni, commenti descrittivi o spiegazioni testuali, a meno che non sia io a richiederlo esplicitamente.

## 2. Architettura & Anti-Ridondanza

- **Verifica moduli:** Controlla sempre se le funzioni possono essere integrate nei file già esistenti in `/Modules` invece di creare nuovi script.

- **Helper riutilizzabili:** Estrai la logica ripetitiva in funzioni helper anziché duplicare il codice.

## 3. Standard Minimi (Sintetici)

- **Documentazione:** Includi il Comment-Based Help inserendo SOLO `.SYNOPSIS` e `.DESCRIPTION` in forma ultra-concisa (massimo 1 riga ciascuno). Ometti gli altri tag se non richiesti.

- **Robustezza e Naming:** Usa strutture `Try/Catch` essenziali per gli errori e utilizza rigorosamente i verbi approvati di PowerShell (es. `Get-`, `Set-`).