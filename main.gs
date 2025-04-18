/**
 * SheetSense - AI Assistant for Google Sheets
 * 
 * This file contains the main entry points for the SheetSense addon.
 */

/**
 * Creates a custom menu when the spreadsheet is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('SheetSense')
    .addItem('Open Chat', 'showSidebar')
    .addSeparator()
    .addItem('Settings', 'showSettings')
    .addItem('Help', 'showHelp')
    .addToUi();
}

/**
 * Shows the SheetSense sidebar.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('chat')
    .setTitle('SheetSense Chat')
    .setWidth(450);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Shows the settings dialog.
 */
function showSettings() {
  const html = HtmlService.createHtmlOutputFromFile('settings')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'SheetSense Settings');
}

/**
 * Shows the help dialog.
 */
function showHelp() {
  const html = HtmlService.createHtmlOutputFromFile('help')
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'SheetSense Help');
}

/**
 * Saves user settings to Script Properties.
 */
function saveSettings(settings) {
  const userProperties = PropertiesService.getUserProperties();
  
  // Store each setting in user properties
  for (const key in settings) {
    if (settings.hasOwnProperty(key)) {
      userProperties.setProperty(key, settings[key]);
    }
  }
  
  return { success: true };
}

/**
 * Gets user settings from Script Properties.
 */
function getSettings() {
  const userProperties = PropertiesService.getUserProperties();
  const properties = userProperties.getProperties();
  
  // Default settings if none exist
  const settings = {
    geminiApiKey: properties.geminiApiKey || '',
    customInstructions: properties.customInstructions || 'You are SheetSense, an AI assistant that helps with Google Sheets data analysis.',
    saveHistory: properties.saveHistory === 'true' || false
  };
  
  return settings;
}

/**
 * Gets the active spreadsheet data for the AI to analyze.
 */
function getSpreadsheetData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const sheetName = sheet.getName();
  
  // Check if spreadsheet has data
  if (!data || data.length === 0) {
    throw new Error("No data found in this sheet. Please add data to your spreadsheet.");
  }
  
  // Get column headers (assuming first row contains headers)
  const headers = data[0];
  
  // Check if headers exist
  if (!headers || headers.length === 0 || headers.every(header => header === "")) {
    throw new Error("No headers found in the first row. Please add column headers to your spreadsheet.");
  }
  
  // Get actual data (excluding headers)
  const rows = data.slice(1);
  
  return {
    sheetName: sheetName,
    headers: headers,
    rows: rows,
    rowCount: rows.length,
    columnCount: headers.length
  };
}

/**
 * Creates a chart based on the specified parameters.
 */
function createChart(chartParams) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const dataRange = sheet.getRange(chartParams.dataRange);
    
    // Get chart type - Convert string to ChartType enum
    let chartType;
    switch(chartParams.type?.toLowerCase() || 'line') {
      case 'bar':
        chartType = Charts.ChartType.BAR;
        break;
      case 'column':
        chartType = Charts.ChartType.COLUMN;
        break;
      case 'pie':
        chartType = Charts.ChartType.PIE;
        break;
      case 'scatter':
        chartType = Charts.ChartType.SCATTER;
        break;
      case 'area':
        chartType = Charts.ChartType.AREA;
        break;
      default:
        chartType = Charts.ChartType.LINE;
    }
    
    // Create chart using the correct API for Google Sheets
    const chart = sheet.newChart()
      .setChartType(chartType)
      .addRange(dataRange)
      .setPosition(5, 5, 0, 0)
      .setOption('title', chartParams.title || 'Chart')
      .setOption('legend', {position: 'bottom'})
      .build();
    
    sheet.insertChart(chart);
    
    return { success: true, message: 'Chart created successfully.' };
  } catch (error) {
    return { success: false, message: 'Error creating chart: ' + error.toString() };
  }
}

/**
 * Executes a formula or command in the active sheet.
 * @param {string} command - The formula or command to execute
 * @param {string} cellReference - Optional cell reference where to insert the formula
 * @return {object} The result of the operation
 */
function executeCommand(command, cellReference) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    
    // Check if this is a formula to insert
    if (command.startsWith('=')) {
      if (!cellReference) {
        // If no cell reference is provided, use the currently selected cell
        const activeRange = sheet.getActiveRange();
        if (!activeRange) {
          return {
            success: false,
            message: "No cell selected. Please select a cell where to insert the formula."
          };
        }
        cellReference = activeRange.getA1Notation();
      }
      
      // Insert the formula
      const cell = sheet.getRange(cellReference);
      cell.setFormula(command);
      
      return {
        success: true,
        message: `Formula "${command}" inserted in cell ${cellReference}.`
      };
    }
    
    // Handle specific commands
    if (command.toLowerCase().startsWith('/create chart')) {
      // Basic chart creation
      const activeRange = sheet.getActiveRange();
      if (!activeRange) {
        return {
          success: false,
          message: "No data selected. Please select data range for the chart."
        };
      }
      
      const chart = sheet.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(activeRange)
        .setPosition(5, 5, 0, 0)
        .setOption('title', 'Chart from SheetSense')
        .build();
      
      sheet.insertChart(chart);
      
      return {
        success: true,
        message: "Chart created successfully."
      };
    }
    
    // Add more command types as needed
    
    return {
      success: false,
      message: "Command not recognized. Please use '=' for formulas or a supported command."
    };
  } catch (error) {
    console.error("Error executing command: " + error.toString());
    return {
      success: false,
      message: "Error executing command: " + error.toString()
    };
  }
}
