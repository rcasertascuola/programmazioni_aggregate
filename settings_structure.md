# Struttura del Foglio di Configurazione "Settings"

Questo documento descrive la struttura del foglio `Settings`, che guida il funzionamento dello script di generazione dei documenti. Tutta la logica di popolamento delle tabelle è definita qui, rendendo lo script astratto e configurabile.

Il foglio `Settings` deve contenere una riga per ogni tabella che si desidera inserire nel documento. Le colonne sono le seguenti:

---

### Colonne Obbligatorie

| Nome Colonna          | Descrizione                                                                                                                                                                     | Esempio                                                                |
| --------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------- |
| **ID**                | Un identificatore unico per questa regola di configurazione. Utile per il debug.                                                                                                | `COMP_CITT`                                                            |
| **Attivo**            | Indica se la regola deve essere eseguita. Scrivere `SI` per attivarla. Qualsiasi altro valore la disattiva.                                                                       | `SI`                                                                   |
| **NomeTabellaTemplate** | Il testo esatto (senza parentesi) presente nella prima cella della tabella template nel documento. Lo script userà questo testo per identificare quale tabella popolare.             | `CITTADINANZA`                                                         |
| **NomeFoglioDati**    | Il nome del foglio di calcolo da cui leggere i dati per questa tabella.                                                                                                          | `competenze`                                                           |
| **Colonne**           | Un elenco di colonne da estrarre dal `NomeFoglioDati` e inserire nel documento, separate da virgola. L'ordine definisce l'ordine di inserimento nelle celle della tabella.       | `codice,nome`                                                          |
| **TipoLogica**        | Definisce quale logica di elaborazione usare. I valori possibili sono: `TabellaSemplice` o `LogicaModuli`.                                                                       | `TabellaSemplice`                                                      |

---

### Colonne Opzionali

| Nome Colonna   | Descrizione                                                                                                                                                                                                                                                                              | Esempio                                                                    |
| -------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------- |
| **Filtri**     | Permette di filtrare le righe da visualizzare. La sintassi è:<ul><li>**AND**: `chiave1=valore1;chiave2=valore2`</li><li>**OR**: `$or(chiave=valoreA|chiave=valoreB)`</li><li>**Variabili**: Usa `$` per valori dinamici dalla riga corrente di `programmazioni` (es. `$periodo`).</li></ul> | `tipo=cittadinanza;$or(nome_periodo=tutti|nome_periodo=$periodo)`            |
| **Ordinamento**  | Specifica come ordinare le righe prima di inserirle. La sintassi è `nome_colonna:asc` per ascendente o `nome_colonna:desc` per discendente.                                                                                                                                              | `ordine:asc`                                                               |
| **Note**       | Un campo libero per aggiungere commenti o descrizioni sulla regola, per aiutare la manutenzione.                                                                                                                                                                                        | `Tabella delle competenze di cittadinanza per il primo o secondo periodo.` |

---

### Tipi di Logica (`TipoLogica`)

#### 1. `TabellaSemplice`

È la logica più comune. Prende i dati da `NomeFoglioDati`, li filtra secondo i `Filtri`, li ordina e popola la tabella corrispondente a `NomeTabellaTemplate` con le `Colonne` specificate.

#### 2. `LogicaModuli`

È una logica specializzata per gestire la creazione del blocco completo dei moduli, che include le tabelle "MODULO", "Unità Didattiche" e "Abilità". Quando si usa questo tipo:
- `NomeFoglioDati` deve puntare al foglio dei **moduli** (es. `moduli`).
- Lo script cercherà automaticamente i dati correlati nel foglio `ud`, collegandoli tramite le colonne `titolo` (in `moduli`) e `titolo_modulo` (in `ud`).
- I `Filtri` si applicheranno solo ai dati del foglio `moduli`.
- `NomeTabellaTemplate` deve corrispondere al segnaposto della prima tabella del blocco (es. `MODULO`). Lo script troverà e gestirà autonomamente anche le tabelle "Unità Didattiche" e "Abilità".
- La colonna `Colonne` viene ignorata, poiché la struttura di queste tabelle è complessa e gestita internamente dalla logica.
