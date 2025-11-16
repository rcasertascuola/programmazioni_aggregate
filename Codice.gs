/*
 * IDs da configurare (Queste sono solo impostazioni predefinite, 
 * verranno sovrascritte dalla scheda 'templates' se trovate)
 */

/**
 * 1. CREARE UNA VOCE DI MENU
 * (Cambiato il nome della funzione chiamata in 'funzionePrincipale')
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Azioni Sinossi')
      .addItem('Crea Sinossi', 'funzionePrincipale') // Corrisponde alla funzione principale
      .addToUi();
}

/**
 * Funzione principale chiamata dal menu, ora con gestione degli errori.
 */

function funzionePrincipale() {
  try {
    // 1. CARICA CONFIGURAZIONE
    var config = analizzaEstraiDati("templates");
    if (!config || typeof config !== 'object' || Array.isArray(config)) {
      throw new Error("Formato dati 'templates' non valido. Mi aspetto un oggetto Chiave/Valore.");
    }
    const idCartella = config['cartella_destinazione'];
    const idTemplate = config['id_template'];
    if (!idCartella || !idTemplate) {
      throw new Error("Configurazione 'templates' incompleta.");
    }
    
    var i = 0
    
    // 2. CARICA DATI PER MERGE
    var datiMerge = analizzaEstraiDati("programmazioni");
    Logger.log(datiMerge[i]);

    if (!datiMerge[i] || typeof datiMerge[i] !== 'object' || Array.isArray(datiMerge[i])) {
      throw new Error("Formato dati 'programmazioni' non valido. Mi aspetto un oggetto Chiave/Valore.");
    }
    const nomeNuovoFile = datiMerge[i]['alias_disciplina'];
    if (!nomeNuovoFile) {
      throw new Error("Manca 'alias_disciplina' in 'programmazioni'.");
    }
    delete datiMerge['alias']; // Rimuovilo così non cerca di sostituire {{nome_file}}
    
    
    // 3. USA IL GESTORE DOCUMENTO
    Logger.log("Avvio GestoreDocumento...");
    
    // Inizializza l'oggetto con la configurazione
    var gestore = new GestoreDocumento(idTemplate, idCartella);
    
    // Esegui i metodi in sequenza
    var nuovoDocumento = gestore
      .crea(nomeNuovoFile)
      .sostituisciPlaceholder(datiMerge[i])
      .inserisciTabella('COMPETENZE DI INDIRIZZO', analizzaEstraiDati("competenze"), ['codice', 'nome', ])
      .finalizza(); // Salva e chiude

    Logger.log("PROCESSO COMPLETATO.");
    Logger.log("Nuovo documento disponibile a: " + nuovoDocumento.url);

  } catch (e) {
    // Un singolo blocco catch intercetterà qualsiasi errore
    // lanciato dalla classe o dalla logica principale.
    Logger.log("ERRORE FATALE in funzionePrincipale: " + e.message);
    Logger.log("Stack: " + e.stack);
  }
}



/**
 * Classe per gestire la creazione e la manipolazione di un documento
 * a partire da un template.
 */
class GestoreDocumento {
  
  /**
   * @param {string} idTemplate L'ID del file template di Google Docs.
   * @param {string} idCartella L'ID della cartella Drive di destinazione.
   */
  constructor(idTemplate, idCartella) {
    if (!idTemplate || !idCartella) {
      throw new Error("ID Template e ID Cartella sono obbligatori per il costruttore.");
    }
    this.idTemplate = idTemplate;
    this.idCartella = idCartella;
    
    // Proprietà che verranno valorizzate dai metodi
    this.fileCopia = null; // Il file Drive clonato
    this.doc = null;       // Il documento DocumentApp aperto
    this.body = null;      // Il corpo (body) del documento
  }

  /**
   * 1. Clona il template e apre il nuovo documento.
   * @param {string} nomeNuovoFile Il nome da dare al file clonato.
   */
  crea(nomeNuovoFile) {
    if (!nomeNuovoFile) {
      throw new Error("Il 'nomeNuovoFile' è obbligatorio per il metodo crea().");
    }
    try {
      var cartellaDestinazione = DriveApp.getFolderById(this.idCartella);
      var templateFile = DriveApp.getFileById(this.idTemplate);
      
      this.fileCopia = templateFile.makeCopy(nomeNuovoFile, cartellaDestinazione);
      this.doc = DocumentApp.openById(this.fileCopia.getId());
      this.body = this.doc.getBody();
      
      Logger.log("Documento clonato e aperto. ID: " + this.doc.getId());
      return this; // Permette il "chaining" (es. gestore.crea().sostituisci())
      
    } catch (e) {
      Logger.log("ERRORE in crea(): " + e.message);
      // Rilancia l'errore per fermare l'esecuzione in funzionePrincipale
      throw new Error("Fallimento clonazione: " + e.message); 
    }
  }

  /**
   * 2. Sostituisce tutti i placeholder {{chiave}} con i valori di un oggetto.
   * @param {Object} dati Oggetto {chiave: valore} per le sostituzioni.
   */
  sostituisciPlaceholder(dati) {
    if (!this.body) {
      throw new Error("Documento non inizializzato. Chiamare prima il metodo crea().");
    }
    try {
      Logger.log("Avvio sostituzione placeholder...");
      for (var chiave in dati) {
        if (dati.hasOwnProperty(chiave)) {
          // Usiamo replaceText per sostituire tutte le occorrenze
          this.body.replaceText('{{' + chiave + '}}', dati[chiave]);
        }
      }
      Logger.log("Sostituzione completata.");
      return this; // Permette il chaining
      
    } catch (e) {
      Logger.log("ERRORE in sostituisciPlaceholder(): " + e.message);
      throw e;
    }
  }

  /**
   * 3. Trova una tabella, ne usa l'ultima riga come template e la popola.
   * @param {string} tagTabella - La stringa (es. '{{TABELLA_DATI}}') da cercare nella *prima riga* della tabella.
   * @param {Object[]} datiTabella - L'array di oggetti da inserire.
   * @param {string[]} colonneDaInserire - Array di stringhe (es. ['nome', 'email']) che definiscono
   * quali colonne estrarre dagli oggetti e in quale ordine.
   */
  inserisciTabella(tagTabella, datiTabella, colonneDaInserire) {
    if (!this.body) {
      throw new Error("Documento non inizializzato. Chiamare prima il metodo crea().");
    }
    if (!datiTabella || datiTabella.length === 0) {
      Logger.log("Nessun dato fornito per la tabella, salto inserimento.");
      return this; // Non è un errore, salta l'operazione.
    }

    try {
      Logger.log("Ricerca tabella con tag: " + tagTabella);
      
      // 1. Cerca la tabella
      var targetTable = null;
      var firstRow = null;
      var tables = this.body.getTables();
      
      for (var i = 0; i < tables.length; i++) {
        var table = tables[i];
        if (table.getNumRows() > 0) {
          var row = table.getRow(0);
          if (row.getText().includes(tagTabella)) {
            targetTable = table;
            firstRow = row;
            break;
          }
        }
      }

      if (targetTable === null) {
        throw new Error("Tabella con tag '" + tagTabella + "' non trovata nella prima riga.");
      }

      Logger.log("Tabella trovata.");

      // 2. Copia il formato dell'ultima riga (template)
      if (targetTable.getNumRows() < 2) {
        throw new Error("La tabella deve avere almeno 2 righe (intestazione/tag e riga template).");
      }
      var templateRow = targetTable.getRow(targetTable.getNumRows() - 1);
      var numCellsTemplate = templateRow.getNumCells();
      
      // Valida la coerenza delle colonne
      if (numCellsTemplate !== colonneDaInserire.length) {
        throw new Error("Il numero di colonne nel template (" + numCellsTemplate + ") non corrisponde al numero di 'colonneDaInserire' (" + colonneDaInserire.length + ").");
      }
      
      Logger.log("Formato riga template individuato.");
      
      // 3. Cancella la riga template
      targetTable.removeRow(targetTable.getNumRows() - 1);
      Logger.log("Riga template cancellata.");

      // 4. Inserisce i nuovi dati, clonando la riga template per mantenere la formattazione
      datiTabella.forEach(function(dataObject) {
        // Clona la riga template e la aggiunge alla tabella. Questo preserva tutta la formattazione delle celle (inclusi i bordi).
        var newRow = targetTable.appendTableRow(templateRow.copy());
        
        colonneDaInserire.forEach(function(chiave, index) {
          var valore = String(dataObject[chiave] || ''); // Converte in stringa, gestisce null/undefined
          
          // Usa replaceText per sostituire i placeholder nella cella clonata, preservando la formattazione.
          newRow.getCell(index).replaceText('{{' + chiave + '}}', valore);
        });
      });
      
      Logger.log("Inserite " + datiTabella.length + " righe di dati nella tabella.");
      return this; // Permette il chaining

    } catch (e) {
      Logger.log("ERRORE in inserisciTabella(): " + e.message);
      throw e;
    }
  }

  /**
   * 4. Salva, chiude il documento e restituisce i riferimenti.
   */
  finalizza() {
    if (!this.doc) {
      throw new Error("Documento non inizializzato. Chiamare prima il metodo crea().");
    }
    
    try {
      this.doc.saveAndClose();
      Logger.log("Documento salvato e chiuso.");
      
      return { id: this.fileCopia.getId(), url: this.fileCopia.getUrl() };
      
    } catch (e) {
      Logger.log("ERRORE in finalizza(): " + e.message);
      
      // Se il salvataggio fallisce, proviamo a cestinare il file
      // per evitare di lasciare "spazzatura" in Drive.
      if (this.fileCopia) {
        DriveApp.getFileById(this.fileCopia.getId()).setTrashed(true);
        Logger.log("File clonato parzialmente e spostato nel cestino.");
      }
      throw e;
    }
  }
}










/**
 * Analizza una tabella in un foglio e restituisce i dati in un formato specifico
 * basato sulla presenza di intestazioni 'chiave'/'valore' o 'id'.
 *
 * @param {string} sheetName Il nome della scheda da cui leggere i dati.
 * @returns {Object | Array} 
 * - Un Oggetto {k:v} se le colonne 'chiave' e 'valore' sono presenti.
 * - Un Array di Oggetti [{}, {}] se la colonna 'id' è presente.
 * - Un Array vuoto [] in tutti gli altri casi o in caso di errore.
 */
function analizzaEstraiDati(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  // --- Gestione Errori Iniziale ---
  if (sheet == null) {
    Logger.log("Errore: Scheda '" + sheetName + "' non trovata.");
    return []; // Restituisce array vuoto
  }
  
  var allData;
  try {
    allData = sheet.getDataRange().getValues();
  } catch (e) {
    Logger.log("Impossibile leggere i dati dal foglio '" + sheetName + "'. Probabilmente è vuoto.");
    return []; // Restituisce array vuoto
  }

  // Se il foglio è vuoto (0 righe) o ha solo intestazioni (1 riga)
  if (allData.length < 2) {
    Logger.log("Nessuna riga di dati trovata in '" + sheetName + "'.");
    // Determiniamo cosa restituire in base alle sole intestazioni (se presenti)
    if (allData.length === 1) {
      var onlyHeaders = allData[0].map(h => h.toString().trim().toLowerCase());
      var hasChiave = onlyHeaders.includes('chiave');
      var hasValore = onlyHeaders.includes('valore');
      var hasId = onlyHeaders.includes('id');
      
      if (hasChiave && hasValore) return {}; // Restituisce oggetto vuoto
      if (hasId) return []; // Restituisce array vuoto
    }
    return []; // Default: array vuoto
  }

  // --- Elaborazione Dati ---
  
  // Estrae le intestazioni e le "pulisce"
  var headers = allData.shift().map(h => h.toString().trim().toLowerCase());
  
  var hasChiave = headers.includes('chiave');
  var hasValore = headers.includes('valore');
  var hasId = headers.includes('id');
  
  // --- LOGICA DI FORMATTAZIONE ---

  // CASO 1: Formato Chiave/Valore
  if (hasChiave && hasValore) {
    Logger.log("Rilevato formato 'chiave/valore' in '" + sheetName + "'.");
    var chiaveIndex = headers.indexOf('chiave');
    var valoreIndex = headers.indexOf('valore');
    var resultObject = {};
    
    allData.forEach(function(row) {
      var key = row[chiaveIndex];
      if (key && key.toString().trim() !== "") { // Assicura che la chiave esista
        resultObject[key] = row[valoreIndex];
      }
    });
    return resultObject;
  }
  
  // CASO 2: Formato Tabella con ID (Array di Oggetti)
  if (hasId) {
    Logger.log("Rilevato formato tabella con 'id' in '" + sheetName + "'.");
    var resultArray = allData.map(function(row) {
      var rowObject = {};
      headers.forEach(function(header, index) {
        rowObject[header] = row[index];
      });
      return rowObject;
    });
    return resultArray;
  }
  
  // CASO 3: Formato non riconosciuto
  Logger.log("Formato non riconosciuto per '" + sheetName + "'. La tabella non ha né 'chiave'/'valore' né 'id'. Restituisco array vuoto.");
  return []; // "vuoto se non si può"
}


// --- ESEMPIO DI UTILIZZO ---

function testAnalisi() {
  // Supponendo tu abbia una scheda "templates" formattata così:
  // | chiave                | valore                |
  // | cartella_destinazione | 12345ABC              |
  // | id_template           | 67890XYZ              |
  var config = analizzaEstraiDati("templates");
  Logger.log("--- Risultato 'templates' (Oggetto) ---");
  Logger.log(config); 
  // Output atteso: { cartella_destinazione: "12345ABC", id_template: "67890XYZ" }
  // Puoi accedere a: config['cartella_destinazione']


  // Supponendo tu abbia una scheda "utenti" formattata così:
  // | id    | nome  | email               |
  // | 1     | Mario | mario@example.com   |
  // | 2     | Laura | laura@example.com   |
  var utenti = analizzaEstraiDati("programmazioni");
  Logger.log("--- Risultato 'utenti' (Array) ---");
  Logger.log(utenti);
  // Output atteso: [ {id: 1, nome: "Mario", ...}, {id: 2, nome: "Laura", ...} ]
  // Puoi accedere a: utenti[0].nome
}


