/**
 * SheetSense - Predefined Prompts
 * 
 * This file contains predefined prompts for common data analysis tasks.
 */

/**
 * Returns a list of predefined prompt categories and their prompts.
 * @return {object} Categories and prompts
 */
function getPredefinedPrompts() {
  return {
    "Data Analysis": [
      {
        title: "Summarize Data",
        prompt: "Please analyze this data and provide a summary of the key trends and insights."
      },
      {
        title: "Find Outliers",
        prompt: "Can you identify any outliers in this data and explain why they might be occurring?"
      },
      {
        title: "Compare Columns",
        prompt: "Compare these two columns and tell me the main differences and similarities."
      },
      {
        title: "Growth Analysis",
        prompt: "Calculate the growth rates between these time periods and identify any patterns."
      }
    ],
    "Reporting": [
      {
        title: "Executive Summary",
        prompt: "Generate an executive summary of this data highlighting the most important insights."
      },
      {
        title: "Weekly Report",
        prompt: "Create a weekly report based on this data that I can share with my team."
      },
      {
        title: "Chart Recommendation",
        prompt: "What type of chart would best visualize this data and why?"
      },
      {
        title: "KPI Dashboard",
        prompt: "Help me create a KPI dashboard using this data. What metrics should I include?"
      }
    ],
    "Formula Help": [
      {
        title: "Create Formula",
        prompt: "I need a formula that can calculate the following: "
      },
      {
        title: "Fix Formula Error",
        prompt: "This formula is giving me an error. Can you help me fix it? "
      },
      {
        title: "Complex Calculation",
        prompt: "I need to perform this complex calculation across multiple rows. What's the best formula to use?"
      },
      {
        title: "Conditional Formatting",
        prompt: "Help me create a conditional formatting rule to highlight cells based on the following condition: "
      }
    ],
    "Data Cleaning": [
      {
        title: "Remove Duplicates",
        prompt: "Help me find and remove duplicate entries in this data."
      },
      {
        title: "Standardize Format",
        prompt: "I need to standardize the format of these values. How can I do that?"
      },
      {
        title: "Fix Text Issues",
        prompt: "Some text entries have inconsistent formatting. How can I clean them up?"
      },
      {
        title: "Merge Data",
        prompt: "I need to merge these two datasets. What's the best way to do it?"
      }
    ],
    "Advanced Analysis": [
      {
        title: "Forecasting",
        prompt: "Based on this historical data, can you help me forecast the values for the next 3 periods?"
      },
      {
        title: "Correlation Analysis",
        prompt: "Analyze the correlation between these variables and explain what it means."
      },
      {
        title: "Segment Analysis",
        prompt: "Help me segment this data into meaningful groups for analysis."
      },
      {
        title: "Anomaly Detection",
        prompt: "Can you identify any anomalies or unusual patterns in this dataset?"
      }
    ]
  };
}

/**
 * Returns a list of sample data analysis questions based on the current sheet data.
 * @return {array} List of sample questions
 */
function getSampleQuestionsForCurrentData() {
  try {
    const sheetData = getSpreadsheetData();
    const { sheetName, headers, rows, rowCount } = sheetData;
    
    // Basic questions that work for most datasets
    const genericQuestions = [
      `What's a summary of the data in the "${sheetName}" sheet?`,
      `What are the key insights from this data?`,
      `How many unique values are there in each column?`,
      `Are there any outliers in this dataset?`
    ];
    
    // Questions based on the headers
    const headerBasedQuestions = [];
    
    // Check for date columns to suggest time-based analysis
    const potentialDateColumns = headers.filter(header => 
      header.toLowerCase().includes('date') || 
      header.toLowerCase().includes('time') ||
      header.toLowerCase().includes('year') ||
      header.toLowerCase().includes('month')
    );
    
    if (potentialDateColumns.length > 0) {
      headerBasedQuestions.push(
        `What are the trends over time in this data?`,
        `How has the data changed from the earliest to the latest date?`
      );
    }
    
    // Check for numeric columns to suggest statistical analysis
    let hasNumericColumns = false;
    for (let i = 0; i < rows.length && i < 5; i++) {
      for (let j = 0; j < headers.length; j++) {
        if (typeof rows[i][j] === 'number') {
          hasNumericColumns = true;
          break;
        }
      }
      if (hasNumericColumns) break;
    }
    
    if (hasNumericColumns) {
      headerBasedQuestions.push(
        `What are the average, minimum, and maximum values for each numeric column?`,
        `Which columns have the strongest correlation with each other?`
      );
    }
    
    // Check for potential categorical columns
    const potentialCategoricalColumns = headers.filter(header =>
      header.toLowerCase().includes('category') ||
      header.toLowerCase().includes('type') ||
      header.toLowerCase().includes('group') ||
      header.toLowerCase().includes('status') ||
      header.toLowerCase().includes('region') ||
      header.toLowerCase().includes('country')
    );
    
    if (potentialCategoricalColumns.length > 0) {
      const catColumn = potentialCategoricalColumns[0];
      headerBasedQuestions.push(
        `How is the data distributed across different ${catColumn} categories?`,
        `What are the top 3 ${catColumn} categories by count?`
      );
    }
    
    return [...genericQuestions, ...headerBasedQuestions];
  } catch (error) {
    console.error("Error generating sample questions: " + error.toString());
    
    // Return generic questions if there's an error
    return [
      "Can you summarize this data for me?",
      "What are the main trends in this data?",
      "How can I visualize this data effectively?",
      "What formula would you recommend for analyzing this data?"
    ];
  }
}

/**
 * Gets a contextual prompt based on the current spreadsheet state.
 * @param {string} promptType - The type of prompt to generate
 * @return {string} The generated prompt
 */
function getContextualPrompt(promptType) {
  try {
    const sheetData = getSpreadsheetData();
    const { sheetName, headers, rowCount } = sheetData;
    
    switch (promptType) {
      case "analysis":
        return `Analyze the data in the "${sheetName}" sheet with ${rowCount} rows and the following columns: ${headers.join(", ")}. What are the key insights?`;
        
      case "visualization":
        return `What would be the best way to visualize the data in the "${sheetName}" sheet? It has the following columns: ${headers.join(", ")}.`;
        
      case "formula":
        return `I need help creating a formula to analyze the data in the "${sheetName}" sheet. It has the following columns: ${headers.join(", ")}.`;
        
      case "cleaning":
        return `I need to clean the data in the "${sheetName}" sheet. It has ${rowCount} rows and the following columns: ${headers.join(", ")}. What steps should I take?`;
        
      default:
        return `Help me analyze the data in the "${sheetName}" sheet with ${rowCount} rows and the following columns: ${headers.join(", ")}.`;
    }
  } catch (error) {
    console.error("Error generating contextual prompt: " + error.toString());
    return "How can I analyze the data in this spreadsheet?";
  }
}
