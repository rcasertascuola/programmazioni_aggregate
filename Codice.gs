/**
 * @fileoverview Script per la generazione automatizzata di documenti Google Docs
 * basato su un sistema di configurazione flessibile definito in un foglio di calcolo.
 *
 * @version 2.0.0
 * @author Jules - AI Software Engineer
 */

// =================================================================
// 1. ENTRY POINT & MENU UI
// =================================================================

/**
 * Funzione eseguita all'apertura del foglio di calcolo.
 * Aggiunge un menu personalizzato per avviare la generazione del documento.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Genera Documento')
    .addItem('Crea Sinossi', 'main')
    .addToUi();
}

/**
 * Funzione principale che orchestra l'intero processo di generazione del documento.
 * Viene eseguita quando l'utente clicca sulla voce di menu.
 */
function main() {
  try {
    const dataManager = new DataManager();
    const settingsManager = new SettingsManager(dataManager);
    
    // Carica la configurazione e i dati di base.
    const settings = settingsManager.getSettings();
    const datiProgrammazioni = dataManager.getSheetData("programmazioni");
    const datiTemplates = dataManager.getSheetData("templates");
    
    if (!datiProgrammazioni || datiProgrammazioni.length === 0) {
      throw new Error("Nessun dato trovato nel foglio 'programmazioni'.");
    }

    // Itera su ogni riga del foglio "programmazioni".
    datiProgrammazioni.forEach((contestoCorrente, index) => {
      Logger.log(`--- Avvio elaborazione per riga ${index + 1} ---`);

      const nomeNuovoFile = `${contestoCorrente['anno_scolastico']} ${contestoCorrente['classe']} ${contestoCorrente['corso']} ${contestoCorrente['alias_disciplina']}`;

      const idTemplate = datiTemplates['id_template'];
      const idCartella = datiTemplates['cartella_destinazione'];

      if (!idTemplate || !idCartella) {
        throw new Error("ID del template o della cartella di destinazione non trovati nel foglio 'templates'.");
      }

      // Avvia il processo di costruzione del documento.
      const builder = new DocumentBuilder(idTemplate, idCartella, nomeNuovoFile);
      builder.build(settings, dataManager, contestoCorrente);

      Logger.log(`Documento per la riga ${index + 1} creato: ${builder.getFileUrl()}`);
    });

    Logger.log(`PROCESSO COMPLETATO. Elaborate ${datiProgrammazioni.length} righe.`);
    SpreadsheetApp.getUi().alert(`Processo completato con successo! Generati ${datiProgrammazioni.length} documenti.`);

  } catch (e) {
    Logger.log(`ERRORE FATALE: ${e.message}\nStack: ${e.stack}`);
    SpreadsheetApp.getUi().alert(`Errore durante l'esecuzione: ${e.message}`);
  }
}

// =================================================================
// 2. ORCHESTRAZIONE DELLA CREAZIONE DEL DOCUMENTO
// =================================================================

/**
 * Gestisce l'intero ciclo di vita della creazione di un documento:
 * clonazione, sostituzione placeholder e popolamento delle tabelle.
 * @class
 */
class DocumentBuilder {
  /**
   * @param {string} templateId L'ID del file template di Google Docs.
   * @param {string} folderId L'ID della cartella Drive di destinazione.
   * @param {string} outputFileName Il nome del file di output.
   */
  constructor(templateId, folderId, outputFileName) {
    this.templateId = templateId;
    this.folderId = folderId;
    this.outputFileName = outputFileName;
    
    this.doc = null;
    this.body = null;
    this.file = null;
  }

  /**
   * Esegue tutti i passaggi per costruire il documento.
   * @param {Array<Object>} settings Le configurazioni per le tabelle.
   * @param {DataManager} dataManager L'istanza per accedere ai dati.
   * @param {Object} context I dati della riga corrente di "programmazioni".
   */
  build(settings, dataManager, context) {
    this._cloneTemplate();
    this._replacePlaceholders(context);
    this._processTables(settings, dataManager, context);
    this.doc.saveAndClose();
    Logger.log("Documento salvato e chiuso.");
  }

  /**
   * Restituisce l'URL del documento generato.
   * @returns {string}
   */
  getFileUrl() {
    return this.file ? this.file.getUrl() : '';
  }

  /**
   * Clona il documento template.
   * @private
   */
  _cloneTemplate() {
    const templateFile = DriveApp.getFileById(this.templateId);
    const destinationFolder = DriveApp.getFolderById(this.folderId);
    this.file = templateFile.makeCopy(this.outputFileName, destinationFolder);
    this.doc = DocumentApp.openById(this.file.getId());
    this.body = this.doc.getBody();
    Logger.log(`Template clonato. Nuovo ID: ${this.doc.getId()}`);
  }

  /**
   * Sostituisce i placeholder globali (es. {{anno_scolastico}}).
   * @param {Object} context L'oggetto contenente i dati per la sostituzione.
   * @private
   */
  _replacePlaceholders(context) {
    Logger.log("Sostituzione dei placeholder globali...");
    for (const key in context) {
      if (context.hasOwnProperty(key)) {
        this.body.replaceText(`{{${key}}}`, context[key]);
      }
    }
  }

  /**
   * Itera attraverso le tabelle del documento e le processa in base alla configurazione.
   * @param {Array<Object>} settings Le configurazioni per le tabelle.
   * @param {DataManager} dataManager L'istanza per accedere ai dati.
   * @param {Object} context I dati della riga corrente.
   * @private
   */
  _processTables(settings, dataManager, context) {
    const tableFactory = new TableFactory();
    const tables = this.body.getTables();

    tables.forEach(table => {
      if (table.getNumRows() === 0 || table.getRow(0).getNumCells() === 0) return;

      const templateName = table.getRow(0).getCell(0).getText().trim();
      const config = settings.find(s => s.NomeTabellaTemplate === templateName);

      if (config) {
        Logger.log(`Trovata corrispondenza per la tabella template: "${templateName}"`);
        try {
          const tableLogic = tableFactory.create(config, dataManager, this.body, context);
          tableLogic.execute(table);
        } catch (e) {
          Logger.log(`Errore durante l'elaborazione della tabella "${templateName}": ${e.message}`);
          // Continua con le altre tabelle
        }
      }
    });
  }
}

// =================================================================
// 3. GESTIONE DATI E CONFIGURAZIONE
// =================================================================

/**
 * Carica e memorizza nella cache i dati dai fogli di calcolo.
 * @class
 */
class DataManager {
  constructor() {
    this.cache = {};
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }

  /**
   * Recupera i dati da un foglio. Se i dati sono già in cache, restituisce la versione in cache.
   * La logica distingue tra tabelle chiave-valore e tabelle di dati standard.
   * @param {string} sheetName Il nome del foglio da cui leggere i dati.
   * @returns {Object|Array<Object>} Un oggetto se rileva colonne 'chiave'/'valore', altrimenti un array di oggetti.
   */
  getSheetData(sheetName) {
    if (this.cache[sheetName]) {
      return this.cache[sheetName];
    }

    const sheet = this.spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`Attenzione: Foglio "${sheetName}" non trovato.`);
      return [];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log(`Attenzione: Nessun dato trovato nel foglio "${sheetName}".`);
      return [];
    }

    const headers = data.shift().map(h => String(h).trim().toLowerCase());
    const hasChiave = headers.includes('chiave');
    const hasValore = headers.includes('valore');

    let result;

    if (hasChiave && hasValore) {
      // Logica Chiave-Valore
      const chiaveIndex = headers.indexOf('chiave');
      const valoreIndex = headers.indexOf('valore');
      result = data.reduce((obj, row) => {
        const key = row[chiaveIndex];
        if (key) {
          obj[key] = row[valoreIndex];
        }
        return obj;
      }, {});
      Logger.log(`Dati da "${sheetName}" caricati come oggetto Chiave/Valore.`);

    } else {
      // Logica Tabella Standard (Array di Oggetti)
      result = data.map(row => {
        const rowObject = {};
        headers.forEach((header, index) => {
          rowObject[header] = row[index];
        });
        return rowObject;
      });
      Logger.log(`Dati da "${sheetName}" caricati come Array di Oggetti.`);
    }

    this.cache[sheetName] = result;
    return result;
  }
}

/**
 * Legge e interpreta la configurazione dal foglio "Settings".
 * @class
 */
class SettingsManager {
  /**
   * @param {DataManager} dataManager Un'istanza di DataManager per leggere i dati.
   */
  constructor(dataManager) {
    this.dataManager = dataManager;
    this.settings = [];
  }

  /**
   * Restituisce le configurazioni attive.
   * @returns {Array<Object>}
   */
  getSettings() {
    if (this.settings.length > 0) {
      return this.settings;
    }

    const configData = this.dataManager.getSheetData('Settings');
    this.settings = configData.filter(row => String(row['Attivo']).trim().toUpperCase() === 'SI');
    Logger.log(`Caricate ${this.settings.length} configurazioni attive dal foglio "Settings".`);

    return this.settings;
  }
}


// =================================================================
// 4. LOGICHE DI POPOLAMENTO TABELLE (FACTORY & STRATEGIES)
// =================================================================

/**
 * Factory per creare l'oggetto di logica corretto per un dato tipo di tabella.
 * @class
 */
class TableFactory {
  /**
   * Crea un'istanza della classe di logica appropriata.
   * @param {Object} config La riga di configurazione dal foglio "Settings".
   * @param {DataManager} dataManager L'istanza per accedere ai dati.
   * @param {GoogleAppsScript.Document.Body} docBody Il corpo del documento.
   * @param {Object} context I dati della riga corrente.
   * @returns {BaseTableLogic} Un'istanza di una classe che estende BaseTableLogic.
   */
  create(config, dataManager, docBody, context) {
    const data = dataManager.getSheetData(config.NomeFoglioDati);
    switch (config.TipoLogica) {
      case 'TabellaSemplice':
        return new SimpleTableLogic(config, data, docBody, context);
      case 'LogicaModuli':
        const udData = dataManager.getSheetData('ud'); // Dati specifici per questa logica
        return new MasterDetailLogic(config, data, docBody, context, udData);
      default:
        throw new Error(`TipoLogica non riconosciuto: "${config.TipoLogica}"`);
    }
  }
}

/**
 * Classe base astratta per tutte le logiche di tabella.
 * @class
 */
class BaseTableLogic {
  /**
   * @param {Object} config La riga di configurazione.
   * @param {Array<Object>|Object} data I dati da inserire.
   * @param {GoogleAppsScript.Document.Body} docBody Il corpo del documento.
   * @param {Object} context Il contesto di esecuzione.
   */
  constructor(config, data, docBody, context) {
    this.config = config;
    this.data = data;
    this.docBody = docBody;
    this.context = context;
  }

  /**
   * Metodo che deve essere implementato dalle classi figlie.
   * @param {GoogleAppsScript.Document.Table} table La tabella da popolare.
   */
  execute(table) {
    throw new Error("Il metodo execute() deve essere implementato.");
  }
}

/**
 * Logica per popolare una tabella standard con dati filtrati e ordinati.
 * @class
 * @extends BaseTableLogic
 */
class SimpleTableLogic extends BaseTableLogic {

  /**
   * Esegue il popolamento della tabella.
   * @param {GoogleAppsScript.Document.Table} table La tabella fisica nel documento.
   */
  execute(table) {
    const filteredData = this._filterData(this.data);
    const sortedData = this._sortData(filteredData);

    if (sortedData.length === 0) {
      Logger.log(`Nessun dato per la tabella "${this.config.NomeTabellaTemplate}" dopo i filtri. La tabella verrà lasciata vuota o rimossa se necessario.`);
      // Opcionale: rimuovere la tabella se vuota
      // table.removeFromParent();
      return;
    }

    if (table.getNumRows() < 2) {
      throw new Error(`La tabella template "${this.config.NomeTabellaTemplate}" deve avere almeno 2 righe (intestazione + riga template).`);
    }

    const templateRow = table.getRow(table.getNumRows() - 1);
    const columns = this.config.Colonne.split(',').map(c => c.trim());

    sortedData.forEach(rowData => {
      const newRow = table.appendTableRow(templateRow.copy());
      columns.forEach((colName, index) => {
        if (newRow.getNumCells() > index) {
          const cell = newRow.getCell(index);
          const value = String(rowData[colName] || '');
          this._formatCell(cell, value);
        }
      });
    });

    // Rimuovi la riga template originale
    table.removeRow(table.getNumRows() - sortedData.length - 1);
    Logger.log(`Popolate ${sortedData.length} righe nella tabella "${this.config.NomeTabellaTemplate}".`);
  }

  /**
   * Formatta una cella preservando lo stile del template.
   * @param {GoogleAppsScript.Document.TableCell} cell La cella da formattare.
   * @param {string} text Il testo da inserire.
   * @private
   */
  _formatCell(cell, text) {
      const paragraph = cell.getChild(0).asParagraph();
      let attributes = {};

      if (paragraph.getNumChildren() > 0 && paragraph.getChild(0).getType() == DocumentApp.ElementType.TEXT) {
          const originalAttributes = paragraph.getChild(0).asText().getAttributes();
          for (const attr in originalAttributes) {
              if (originalAttributes[attr] !== null) {
                  attributes[attr] = originalAttributes[attr];
              }
          }
      }

      paragraph.clear();
      const newTextElement = paragraph.appendText(text);

      if (Object.keys(attributes).length > 0) {
          newTextElement.setAttributes(attributes);
      }
  }

  /**
   * Filtra i dati in base alla configurazione.
   * Supporta AND (separati da ';'), OR (sintassi $or(...|...)) e variabili di contesto (es. $periodo).
   * @param {Array<Object>} data I dati da filtrare.
   * @returns {Array<Object>} I dati filtrati.
   * @private
   */
  _filterData(data) {
    if (!this.config.Filtri) return data;
    
    const filters = this.config.Filtri.split(';').map(f => f.trim());

    return data.filter(row => {
      return filters.every(filter => {
        // Logica OR
        if (filter.toLowerCase().startsWith('$or(')) {
          const orConditions = filter.substring(4, filter.length - 1).split('|');
          return orConditions.some(orCond => this._evaluateCondition(row, orCond));
        }
        // Logica AND
        return this._evaluateCondition(row, filter);
      });
    });
  }

  /**
   * Valuta una singola condizione di filtro.
   * @param {Object} row La riga di dati.
   * @param {string} condition La condizione (es. "tipo=cittadinanza").
   * @returns {boolean}
   * @private
   */
  _evaluateCondition(row, condition) {
    const parts = condition.split('=');
    if (parts.length !== 2) return true; // Condizione malformata, la ignoriamo

    const key = parts[0].trim();
    let value = parts[1].trim();

    // Sostituisci variabili di contesto
    if (value.startsWith('$')) {
      const contextKey = value.substring(1);
      value = this.context[contextKey] || '';
    }

    const rowValue = String(row[key] || '').trim().toLowerCase();
    const filterValue = String(value).trim().toLowerCase();

    return rowValue === filterValue;
  }

  /**
   * Ordina i dati in base alla configurazione.
   * @param {Array<Object>} data I dati da ordinare.
   * @returns {Array<Object>} I dati ordinati.
   * @private
   */
  _sortData(data) {
    if (!this.config.Ordinamento) return data;

    const [column, direction] = this.config.Ordinamento.split(':').map(p => p.trim());
    const desc = direction.toLowerCase() === 'desc';

    return data.sort((a, b) => {
      if (a[column] < b[column]) return desc ? 1 : -1;
      if (a[column] > b[column]) return desc ? -1 : 1;
      return 0;
    });
  }
}

/**
 * Logica specializzata per creare il blocco di tabelle Modulo, UD e Abilità.
 * @class
 * @extends BaseTableLogic
 */
class MasterDetailLogic extends BaseTableLogic {

  /**
   * @param {Object} config La riga di configurazione.
   * @param {Array<Object>} data I dati dei moduli.
   * @param {GoogleAppsScript.Document.Body} docBody Il corpo del documento.
   * @param {Object} context Il contesto di esecuzione.
   * @param {Array<Object>} udData I dati delle unità didattiche.
   */
  constructor(config, data, docBody, context, udData) {
    super(config, data, docBody, context);
    this.udData = udData; // Dati aggiuntivi per le sotto-tabelle
  }

  /**
   * Esegue la creazione del blocco di tabelle.
   * Questo metodo ignora il parametro `table` perché deve gestire più tabelle.
   * @override
   */
  execute() {
    // Applica filtri e ordinamento ai dati principali (moduli)
    const simpleFilter = new SimpleTableLogic(this.config, this.data, this.docBody, this.context);
    const filteredModules = simpleFilter._filterData(this.data);
    const sortedModules = simpleFilter._sortData(filteredModules);

    if (sortedModules.length === 0) {
      Logger.log("Nessun modulo trovato dopo il filtro. Blocco moduli non creato.");
      return;
    }

    if (!this.config.TabelleCorrelate) {
      throw new Error(`Configurazione "${this.config.ID}" di tipo LogicaModuli non valida: la colonna 'TabelleCorrelate' è obbligatoria.`);
    }
    const templateTags = this.config.TabelleCorrelate.split(',').map(t => t.trim());
    if (templateTags.length < 3) {
      throw new Error(`La colonna 'TabelleCorrelate' deve contenere almeno 3 nomi di tabella separati da virgola.`);
    }

    const { templates, position } = this._findAndRemoveTemplateTables(templateTags);

    let insertionIndex = position;

    sortedModules.forEach(modulo => {
      // 1. Popola e inserisci la tabella MODULO (la prima della lista)
      this._insertModuloTable(templates[templateTags[0]], modulo, insertionIndex++);

      // 2. Filtra, popola e inserisci la tabella Unità Didattiche (la seconda)
      const udFiltrate = this.udData
        .filter(ud => String(ud.titolo_modulo).trim() === String(modulo.titolo).trim())
        .sort((a,b) => a.ordinale - b.ordinale);
      this._insertUdTable(templates[templateTags[1]], udFiltrate, insertionIndex++);

      // 3. Popola e inserisci la tabella Abilità (la terza)
      this._insertAbilitaTable(templates[templateTags[2]], modulo, insertionIndex++);
    });
  }

  /**
   * Trova le tabelle template, le copia e le rimuove dal documento.
   * @param {Array<string>} tags I tag da cercare nella prima cella di ogni tabella.
   * @returns {{templates: Object, position: number}} Un oggetto con le tabelle copiate e la posizione di inserimento.
   * @private
   */
  _findAndRemoveTemplateTables(tags) {
    const templates = {};
    let position = -1;
    const allTables = this.docBody.getTables();

    tags.forEach(tag => {
      const foundTable = allTables.find(t => t.getNumRows() > 0 && t.getRow(0).getCell(0).getText().trim() === tag);
      if (foundTable) {
        templates[tag] = foundTable.copy();
        const tableIndex = this.docBody.getChildIndex(foundTable);
        if (position === -1 || tableIndex < position) {
          position = tableIndex;
        }
        foundTable.removeFromParent();
      } else {
        throw new Error(`Tabella template con tag "${tag}" non trovata.`);
      }
    });

    if (Object.keys(templates).length !== tags.length) {
      throw new Error("Non è stato possibile trovare tutte le tabelle template per LogicaModuli.");
    }

    Logger.log(`Trovate e rimosse le tabelle template per LogicaModuli. Posizione di inserimento: ${position}`);
    return { templates, position };
  }

  _insertModuloTable(templateTable, moduloData, index) {
    const newTable = templateTable.copy();
    // Esempio di popolamento - adattare alle proprie esigenze
    newTable.getRow(0).getCell(0).setText(`MODULO ${moduloData.ordine}`);
    newTable.getRow(0).getCell(1).setText(`UDA - ${moduloData.titolo_uda}: ${moduloData.titolo}`);
    newTable.getRow(1).getCell(1).setText(moduloData.tempi_modulo);
    this.docBody.insertTable(index, newTable);
  }

  _insertUdTable(templateTable, udData, index) {
    const newTable = templateTable.copy();
    const templateRow = newTable.getRow(newTable.getNumRows() - 1);

    udData.forEach(ud => {
      const newRow = newTable.appendTableRow(templateRow.copy());
      newRow.getCell(0).setText(ud.titolo || '');
      newRow.getCell(1).setText(ud.conoscenze || '');
    });

    newTable.removeRow(newTable.getNumRows() - udData.length - 1);
    this.docBody.insertTable(index, newTable);
  }

  _insertAbilitaTable(templateTable, moduloData, index) {
    const newTable = templateTable.copy();
    const templateRow = newTable.getRow(newTable.getNumRows() - 1);
    const newRow = newTable.appendTableRow(templateRow.copy());

    const cell = newRow.getCell(0);
    cell.clear(); // Lascia un paragrafo vuoto

    const p1 = cell.getChild(0).asParagraph();
    p1.appendText("abilità cognitive: ").setBold(true);
    p1.appendText(String(moduloData.abilità_specifiche_cognitive || '')).setBold(false);

    const p2 = cell.appendParagraph('');
    p2.appendText("abilità teoriche: ").setBold(true);
    const pratiche = `${moduloData.abilità_specifiche_pratiche || ''} (${moduloData.abilità || ''})`.trim();
    p2.appendText(pratiche).setBold(false);

    const competenze = `${moduloData.competenze_specifiche || ''} (${moduloData.competenze || ''})`.trim();
    newRow.getCell(1).setText(competenze);

    newTable.removeRow(newTable.getNumRows() - 2);
    this.docBody.insertTable(index, newTable);
  }
}
