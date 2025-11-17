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
    var dataManager = new DataManager();
    var config = dataManager.getSheetData("templates");

    if (!config || typeof config !== 'object' || Array.isArray(config)) {
      throw new Error("Formato dati 'templates' non valido. Mi aspetto un oggetto Chiave/Valore.");
    }
    const idCartella = config['cartella_destinazione'];
    const idTemplate = config['id_template'];
    if (!idCartella || !idTemplate) {
      throw new Error("Configurazione 'templates' incompleta.");
    }
  
    var parametri_elenchi = dataManager.getSheetData("parametri_elenchi");
    if (!config || typeof parametri_elenchi !== 'object' || Array.isArray(config)) {
      throw new Error("Formato dati 'parametri_elenchi' non valido. Mi aspetto un oggetto Chiave/Valore.");
    }
    var i = 0
    
    // 2. CARICA DATI PER MERGE
    var datiMerge = dataManager.getSheetData("programmazioni");
    
    if (!datiMerge[i] || typeof datiMerge[i] !== 'object' || Array.isArray(datiMerge[i])) {
      throw new Error("Formato dati 'programmazioni' non valido. Mi aspetto un oggetto Chiave/Valore.");
    }
    const nomeNuovoFile = datiMerge[i]['anno_scolastico'] + " " + datiMerge[i]['classe'] + datiMerge[i]['corso'] + " " + datiMerge[i]['alias_disciplina'];
    if (!datiMerge[i]['anno_scolastico'] || !datiMerge[i]['classe'] || !datiMerge[i]['corso'] || !datiMerge[i]['alias_disciplina']) {
      throw new Error("Dati insufficienti in 'programmazioni' per creare il nome del file. Sono richiesti 'anno_scolastico', 'classe', 'corso', e 'alias_disciplina'.");
    }
    delete datiMerge['alias']; // Rimuovilo così non cerca di sostituire {{nome_file}}
  
    // Carica i dati per le nuove tabelle
    var datiModuli = dataManager.getSheetData("moduli");
    var datiUd = dataManager.getSheetData("ud");

    // Ordina i moduli in base alla colonna 'ordine'
    datiModuli.sort(function(a, b) {
      return a.ordine - b.ordine;
    });

    // 3. USA IL GESTORE DOCUMENTO
    Logger.log("Avvio GestoreDocumento...");
    
    // Inizializza l'oggetto con la configurazione
    var gestore = new GestoreDocumento(idTemplate, idCartella);
    
    // Esegui i metodi in sequenza
    var nuovoDocumento = gestore
      .crea(nomeNuovoFile)
      .sostituisciPlaceholder(datiMerge[i])
      .sostituisciPlaceholder(parametri_elenchi)
      .inserisciTabella('eqf', dataManager.getSheetData("eqf"), ['periodo',	'livello','conoscenze',	'abilità','competenze'], { 'periodo': datiMerge[i]['periodo'] })
      .inserisciTabella('PERMANENTE', dataManager.getSheetData("competenze"), ['codice', 'nome', ], { 'tipo': 'apprendimento permanente', '$or': [{ 'nome_periodo': 'tutti' }, { 'nome_periodo': datiMerge[i]['periodo'] }] })
      .inserisciTabella('CITTADINANZA', dataManager.getSheetData("competenze"), ['codice', 'nome', ], { 'tipo': 'cittadinanza', '$or': [{ 'nome_periodo': 'tutti' }, { 'nome_periodo': datiMerge[i]['periodo'] }] })
      .inserisciTabella('INDIRIZZO', dataManager.getSheetData("competenze"), ['codice', 'nome', ], { 'tipo': 'indirizzo', '$or': [{ 'nome_periodo': 'tutti' }, { 'nome_periodo': datiMerge[i]['periodo'] }] })
      .inserisciTabella('DISCIPLINARI', dataManager.getSheetData("competenze"), ['codice', 'nome', ], { 'tipo': 'disciplinari', '$or': [{ 'nome_periodo': 'tutti' }, { 'nome_periodo': datiMerge[i]['periodo'] }] })
      .creaTabelleDeiModuli(datiModuli, datiUd)
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
   * 3. Trova una tabella, ne usa l'ultima riga come template e la popola, con un filtro opzionale.
   * @param {string} tagTabella - La stringa (es. '{{TABELLA_DATI}}') da cercare nella *prima riga* della tabella.
   * @param {Object[]} datiTabella - L'array di oggetti da inserire.
   * @param {string[]} colonneDaInserire - Array di stringhe (es. ['nome', 'email']) che definiscono quali colonne estrarre.
   * @param {Object} [filtro=null] - Un oggetto opzionale per filtrare i dati (es. {'colonna': 'valore'}).
   */
  inserisciTabella(tagTabella, datiTabella, colonneDaInserire, filtro = null) {
    if (!this.body) {
      throw new Error("Documento non inizializzato. Chiamare prima il metodo crea().");
    }

    // Applica il filtro se fornito
    var datiFiltrati = datiTabella;
    if (filtro) {
      datiFiltrati = datiTabella.filter(function(riga) {
        var andMatch = true;
        var orMatch = false;

        // Controlla prima le condizioni AND
        for (var chiave in filtro) {
          if (chiave !== '$or') {
            if (!riga.hasOwnProperty(chiave) || String(riga[chiave]) !== String(filtro[chiave])) {
              andMatch = false;
              break;
            }
          }
        }

        if (!andMatch) return false; // Se la parte AND fallisce, la riga è esclusa

        // Se ci sono condizioni OR, controllale
        if (filtro['$or']) {
          var orConditions = filtro['$or'];
          for (var i = 0; i < orConditions.length; i++) {
            var condition = orConditions[i];
            for (var chiave in condition) {
              if (riga.hasOwnProperty(chiave) && String(riga[chiave]) === String(condition[chiave])) {
                orMatch = true;
                break;
              }
            }
            if (orMatch) break;
          }
          return orMatch; // Il risultato finale dipende dalla corrispondenza OR
        }

        return true; // Se c'erano solo condizioni AND e sono state superate
      });
      Logger.log("Dati filtrati. Righe rimanenti: " + datiFiltrati.length);
    }

    if (!datiFiltrati || datiFiltrati.length === 0) {
      Logger.log("Nessun dato da inserire nella tabella dopo il filtraggio.");
      return this;
    }

    try {
      var targetTable = null;
      var tables = this.body.getTables();
      for (var i = 0; i < tables.length; i++) {
        if (tables[i].getNumRows() > 0 && tables[i].getRow(0).getText().includes(tagTabella)) {
          targetTable = tables[i];
          break;
        }
      }

      if (!targetTable) throw new Error("Tabella con tag '" + tagTabella + "' non trovata.");
      Logger.log("Tabella trovata.");

      if (targetTable.getNumRows() < 2) throw new Error("La tabella deve avere almeno 2 righe.");
      
      var templateRow = targetTable.getRow(targetTable.getNumRows() - 1);
      
      datiFiltrati.forEach(function(dataObject) {
        var newRow = targetTable.appendTableRow(templateRow.copy());
        colonneDaInserire.forEach(function(chiave, index) {
            var valore = String(dataObject[chiave] || '');
            var cella = newRow.getCell(index);
            var textElement = cella.getChild(0).asParagraph().getChild(0);
            if (textElement && textElement.getType() == DocumentApp.ElementType.TEXT) {
              var attributi = textElement.asText().getAttributes();
              cella.setText(valore);
              cella.getChild(0).asParagraph().getChild(0).setAttributes(attributi);
            } else {
              cella.setText(valore);
            }
        });
      });

      targetTable.removeRow(targetTable.getNumRows() - datiFiltrati.length -1);
      
      Logger.log("Inserite " + datiFiltrati.length + " righe di dati nella tabella.");
      return this;

    } catch (e) {
      Logger.log("ERRORE in inserisciTabella(): " + e.message + " Stack: " + e.stack);
      throw e;
    }
  }

  /**
   * Crea e popola le tabelle dei moduli, unità didattiche e abilità.
   * @param {Object[]} datiModuli L'array di oggetti dei moduli.
   * @param {Object[]} datiUd L'array di oggetti delle unità didattiche.
   */
  creaTabelleDeiModuli(datiModuli, datiUd) {
    if (!this.body) {
      throw new Error("Documento non inizializzato. Chiamare prima il metodo crea().");
    }

    try {
      var infoTabelle = this._trovaERimuoviTabelleTemplate(['MODULO', 'Unità Didattiche', 'Abilità']);
      var indiceInserimento = infoTabelle.posizione;

      datiModuli.forEach(function(modulo) {
        // --- Crea e popola la tabella MODULO ---
        var nuovaTabellaModulo = infoTabelle.templates['MODULO'].copy();
        var cellaModulo = nuovaTabellaModulo.getRow(0).getCell(0);
        var attributiModulo = cellaModulo.getChild(0).asParagraph().getChild(0).asText().getAttributes();
        cellaModulo.setText('MODULO ' + modulo.ordine);
        cellaModulo.getChild(0).asParagraph().getChild(0).setAttributes(attributiModulo);

        nuovaTabellaModulo.getRow(0).getCell(1).setText('UDA - ' + modulo.titolo_uda + ": " + modulo.titolo);
        nuovaTabellaModulo.getRow(1).getCell(1).setText(modulo.tempi_modulo);
        this.body.insertTable(indiceInserimento++, nuovaTabellaModulo);
        this.body.insertParagraph(indiceInserimento++, "");

        // --- Crea e popola la tabella Unità Didattiche ---
        var datiUdFiltrati = datiUd.filter(function(ud) { return ud.id_modulo === modulo.id; });
        datiUdFiltrati.sort(function(a, b) { return a.ordinale - b.ordinale; });

        var nuovaTabellaUd = infoTabelle.templates['Unità Didattiche'].copy();
        var templateRowUd = nuovaTabellaUd.getRow(nuovaTabellaUd.getNumRows() - 1);
        datiUdFiltrati.forEach(function(ud) {
          var newRow = nuovaTabellaUd.appendTableRow(templateRowUd.copy());
          newRow.getCell(0).setText(ud.titolo);
          newRow.getCell(1).setText(ud.conoscenze);
        });
        nuovaTabellaUd.removeRow(nuovaTabellaUd.getNumRows() - datiUdFiltrati.length - 1);
        this.body.insertTable(indiceInserimento++, nuovaTabellaUd);
        this.body.insertParagraph(indiceInserimento++, "");

        // --- Crea e popola la tabella Abilità ---
        var nuovaTabellaAbilita = infoTabelle.templates['Abilità'].copy();
        var templateRowAbilita = nuovaTabellaAbilita.getRow(nuovaTabellaAbilita.getNumRows() - 1);
        var newRowAbilita = nuovaTabellaAbilita.appendTableRow(templateRowAbilita.copy());

        var cellaAbilita = newRowAbilita.getCell(0);
        cellaAbilita.clear(); // Questo lascia un paragrafo vuoto

        var primoParagrafo = cellaAbilita.getParagraphs()[0];
        primoParagrafo.appendText("abilità cognitive: ").setBold(true);
        primoParagrafo.appendText(modulo.abilità_specifiche_cognitive).setBold(false);

        var secondoParagrafo = cellaAbilita.appendParagraph('');
        secondoParagrafo.appendText("abilità teoriche: ").setBold(true);
        secondoParagrafo.appendText(modulo.abilità_specifiche_pratiche + ' (' + modulo.abilità + ')').setBold(false);

        var competenzeTesto = modulo.competenze_specifiche + ' (' + modulo.competenze + ')';
        newRowAbilita.getCell(1).setText(competenzeTesto);
        nuovaTabellaAbilita.removeRow(nuovaTabellaAbilita.getNumRows() - 2);
        this.body.insertTable(indiceInserimento++, nuovaTabellaAbilita);
        this.body.insertParagraph(indiceInserimento++, "");

      }, this);

    } catch (e) {
      Logger.log("ERRORE in creaTabelleDeiModuli(): " + e.message + " Stack: " + e.stack);
      throw e;
    }
    return this;
  }

  _trovaERimuoviTabelleTemplate(tags) {
    var templates = {};
    var posizione = -1;
    var tabelle = this.body.getTables();
    var tagNormalizzati = tags.map(function(t) { return t.replace(/\s/g, '').toLowerCase(); });

    tags.forEach(function(tag, index) {
      for (var i = 0; i < tabelle.length; i++) {
        var testoCella = tabelle[i].getRow(0).getCell(0).getText().replace(/\s/g, '').toLowerCase();
        if (testoCella === tagNormalizzati[index]) {
          templates[tag] = tabelle[i].copy();
          var indiceTabella = this.body.getChildIndex(tabelle[i]);
          if (posizione === -1 || indiceTabella < posizione) {
            posizione = indiceTabella;
          }
          tabelle[i].removeFromParent();
          break;
        }
      }
    }, this);

    if (Object.keys(templates).length !== tags.length) {
      throw new Error("Non è stato possibile trovare tutte le tabelle template.");
    }
    return { templates: templates, posizione: posizione };
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

class DataManager {
  constructor() {
    this.cache = {};
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
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
  getSheetData(sheetName) {
    if (this.cache[sheetName]) {
      return this.cache[sheetName];
    }

    var sheet = this.ss.getSheetByName(sheetName);

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

    let result;
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
      result = resultObject;
    }
    // CASO 2: Formato Tabella con ID (Array di Oggetti)
    else if (hasId) {
      Logger.log("Rilevato formato tabella con 'id' in '" + sheetName + "'.");
      var resultArray = allData.map(function(row) {
        var rowObject = {};
        headers.forEach(function(header, index) {
          rowObject[header] = row[index];
        });
        return rowObject;
      });
      result = resultArray;
    }
    // CASO 3: Formato non riconosciuto
    else {
      Logger.log("Formato non riconosciuto per '" + sheetName + "'. La tabella non ha né 'chiave'/'valore' né 'id'. Restituisco array vuoto.");
      result = []; // "vuoto se non si può"
    }

    this.cache[sheetName] = result;
    return result;
  }
}
