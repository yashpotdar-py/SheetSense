/**
 * SheetSense - Tools and Utilities
 * 
 * This file contains utility functions for data analysis and spreadsheet manipulation.
 */

/**
 * Performs data analysis on the specified range.
 * @param {string} range - The range to analyze (e.g., "A1:C10")
 * @param {string} analysisType - The type of analysis to perform (e.g., "summary", "correlation")
 * @return {object} The analysis results
 */
function analyzeData(range, analysisType = "summary") {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const rangeObj = sheet.getRange(range);
    const values = rangeObj.getValues();
    
    // Extract headers (first row) and data (remaining rows)
    const headers = values[0];
    const data = values.slice(1);
    
    let result = {};
    
    switch (analysisType.toLowerCase()) {
      case "summary":
        result = calculateSummaryStatistics(data, headers);
        break;
      case "correlation":
        result = calculateCorrelations(data, headers);
        break;
      default:
        result = {
          success: false,
          error: `Unknown analysis type: ${analysisType}`
        };
    }
    
    return {
      success: true,
      analysisType: analysisType,
      results: result
    };
  } catch (error) {
    console.error("Error in analyzeData: " + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Calculates summary statistics for each column in the data.
 */
function calculateSummaryStatistics(data, headers) {
  const results = {};
  
  // For each column
  for (let colIndex = 0; colIndex < headers.length; colIndex++) {
    const header = headers[colIndex];
    const columnData = data.map(row => row[colIndex]);
    
    // Filter out non-numeric values for statistical calculations
    const numericValues = columnData.filter(value => 
      typeof value === 'number' && !isNaN(value));
    
    const stats = {
      count: columnData.length,
      numericCount: numericValues.length,
      uniqueCount: [...new Set(columnData)].length,
    };
    
    // Only calculate numeric statistics if we have numeric values
    if (numericValues.length > 0) {
      // Mean
      const sum = numericValues.reduce((acc, val) => acc + val, 0);
      stats.mean = sum / numericValues.length;
      
      // Min, Max
      stats.min = Math.min(...numericValues);
      stats.max = Math.max(...numericValues);
      
      // Calculate median
      const sorted = [...numericValues].sort((a, b) => a - b);
      const middle = Math.floor(sorted.length / 2);
      stats.median = sorted.length % 2 === 0
        ? (sorted[middle - 1] + sorted[middle]) / 2
        : sorted[middle];
        
      // Calculate standard deviation
      const meanDiffs = numericValues.map(val => Math.pow(val - stats.mean, 2));
      const variance = meanDiffs.reduce((acc, val) => acc + val, 0) / numericValues.length;
      stats.stdDev = Math.sqrt(variance);
    }
    
    results[header] = stats;
  }
  
  return results;
}

/**
 * Calculates correlations between numeric columns.
 */
function calculateCorrelations(data, headers) {
  const results = {};
  
  // Only include numeric columns
  const numericColumns = [];
  const numericHeaders = [];
  
  for (let colIndex = 0; colIndex < headers.length; colIndex++) {
    const header = headers[colIndex];
    const columnData = data.map(row => row[colIndex]);
    
    // Check if column has numeric values
    const numericValues = columnData.filter(value => 
      typeof value === 'number' && !isNaN(value));
    
    if (numericValues.length > 0 && numericValues.length === data.length) {
      numericColumns.push(columnData);
      numericHeaders.push(header);
    }
  }
  
  // Calculate correlations between numeric columns
  for (let i = 0; i < numericHeaders.length; i++) {
    const xHeader = numericHeaders[i];
    results[xHeader] = {};
    
    for (let j = 0; j < numericHeaders.length; j++) {
      const yHeader = numericHeaders[j];
      if (i === j) {
        results[xHeader][yHeader] = 1;
      } else {
        const correlation = calculatePearsonCorrelation(
          numericColumns[i],
          numericColumns[j]
        );
        results[xHeader][yHeader] = correlation;
      }
    }
  }
  
  return results;
}

/**
 * Calculates Pearson correlation coefficient between two arrays.
 */
function calculatePearsonCorrelation(x, y) {
  const n = x.length;
  let sumX = 0;
  let sumY = 0;
  let sumXY = 0;
  let sumX2 = 0;
  let sumY2 = 0;
  
  for (let i = 0; i < n; i++) {
    sumX += x[i];
    sumY += y[i];
    sumXY += x[i] * y[i];
    sumX2 += x[i] * x[i];
    sumY2 += y[i] * y[i];
  }
  
  const numerator = n * sumXY - sumX * sumY;
  const denominator = Math.sqrt((n * sumX2 - sumX * sumX) * (n * sumY2 - sumY * sumY));
  
  if (denominator === 0) return 0;
  return numerator / denominator;
}

/**
 * Creates a pivot table with the provided parameters.
 */
function createPivotTable(params) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = sheet.getSheetByName(params.sourceSheet);
    
    if (!sourceSheet) {
      return {
        success: false,
        error: `Source sheet '${params.sourceSheet}' not found`
      };
    }
    
    // Create a new sheet for the pivot table
    let pivotSheet = sheet.getSheetByName(params.pivotSheetName || "Pivot Table");
    if (!pivotSheet) {
      pivotSheet = sheet.insertSheet(params.pivotSheetName || "Pivot Table");
    }
    
    // Get source data range
    const sourceRange = sourceSheet.getRange(params.sourceRange);
    
    // Create the pivot table
    const pivotTable = pivotSheet.getRange("A1").createPivotTable(sourceRange);
    
    // Add rows, columns, and values based on params
    if (params.rows && params.rows.length > 0) {
      params.rows.forEach(rowField => {
        pivotTable.addRowGroup(rowField);
      });
    }
    
    if (params.columns && params.columns.length > 0) {
      params.columns.forEach(columnField => {
        pivotTable.addColumnGroup(columnField);
      });
    }
    
    if (params.values && params.values.length > 0) {
      params.values.forEach(valueField => {
        pivotTable.addPivotValue(
          valueField.field, 
          valueField.summarizeFunction || SpreadsheetApp.PivotTableSummarizeFunction.SUM
        );
      });
    }
    
    return {
      success: true,
      message: `Pivot table created in sheet '${pivotSheet.getName()}'`
    };
  } catch (error) {
    console.error("Error in createPivotTable: " + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Creates a formula in the specified cell.
 */
function createFormula(cellReference, formulaText) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const cell = sheet.getRange(cellReference);
    cell.setFormula(formulaText);
    
    return {
      success: true,
      message: `Formula created in cell ${cellReference}`
    };
  } catch (error) {
    console.error("Error in createFormula: " + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Gets sample formulas for common tasks.
 */
function getSampleFormulas(category) {
  const formulas = {
    "lookup": [
      { name: "VLOOKUP", formula: "=VLOOKUP(search_key, range, index, [is_sorted])" },
      { name: "HLOOKUP", formula: "=HLOOKUP(search_key, range, index, [is_sorted])" },
      { name: "INDEX/MATCH", formula: "=INDEX(return_range, MATCH(lookup_value, lookup_range, 0))" }
    ],
    "date": [
      { name: "Current Date", formula: "=TODAY()" },
      { name: "Current Date and Time", formula: "=NOW()" },
      { name: "Extract Year", formula: "=YEAR(date)" },
      { name: "Extract Month", formula: "=MONTH(date)" },
      { name: "Extract Day", formula: "=DAY(date)" }
    ],
    "text": [
      { name: "Concatenate Text", formula: "=CONCATENATE(text1, text2, ...)" },
      { name: "Convert to Uppercase", formula: "=UPPER(text)" },
      { name: "Convert to Lowercase", formula: "=LOWER(text)" },
      { name: "Extract Substring", formula: "=MID(text, start_position, number_of_characters)" }
    ],
    "statistical": [
      { name: "Average", formula: "=AVERAGE(range)" },
      { name: "Median", formula: "=MEDIAN(range)" },
      { name: "Standard Deviation", formula: "=STDEV(range)" },
      { name: "Correlation", formula: "=CORREL(range1, range2)" }
    ],
    "conditional": [
      { name: "IF Statement", formula: "=IF(logical_expression, value_if_true, value_if_false)" },
      { name: "Nested IF", formula: "=IF(condition1, value1, IF(condition2, value2, value3))" },
      { name: "SUMIF", formula: "=SUMIF(range, criteria, [sum_range])" },
      { name: "COUNTIF", formula: "=COUNTIF(range, criteria)" }
    ]
  };
  
  return category ? formulas[category] || [] : formulas;
}
