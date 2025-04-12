/*
 * SheetSense - AI Agent
 * 
 * This file contains functions for interacting with the Gemini AI model.
 */

/**
 * Sends a user query to the Gemini API and gets a response.
 * @param {string} query - The user's question or request
 * @param {object} context - Contextual data to help the AI understand the request
 * @param {array} history - Previous chat history for context
 * @return {object} The AI response
 */
function callGeminiAPI(query, context, history = []) {
  try {
    // Get user settings
    const settings = getSettings();
    const apiKey = settings.geminiApiKey;
    
    // Check if API key is provided
    if (!apiKey) {
      return {
        success: false,
        text: "Please set up your Gemini API key in the settings. Get one from Google AI Studio at https://makersuite.google.com/app/apikey",
        error: "API key not found"
      };
    }
    
    // Prepare prompt with context
    const customInstructions = settings.customInstructions || 
      "You are SheetSense, an AI assistant for Google Sheets. Help users analyze data and create visualizations.";
    
    // Format history into proper context
    const formattedHistory = history.map(msg => {
      return {
        role: msg.role,
        parts: [{ text: msg.content }]
      };
    });
    
    // Prepare sheet data to provide context
    let sheetContext = "";
    if (context && context.sheetData) {
      const { sheetName, headers, rowCount, columnCount } = context.sheetData;
      sheetContext = `
        You're analyzing a Google Sheet named "${sheetName}" with ${rowCount} rows of data and ${columnCount} columns.
        The column headers are: ${headers.join(", ")}.
        
        First few rows of data (sample):
        ${formatSampleData(context.sheetData.rows.slice(0, 5), headers)}
      `;
    }
    
    // Build the complete prompt
    const prompt = `${customInstructions}
    
    ${sheetContext}
    
    User query: ${query}
    `;
    
    // Make API request to Gemini
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;
    
    const payload = {
      contents: [
        {
          parts: [{ text: prompt }]
        }
      ],
      generationConfig: {
        temperature: 0.7,
        maxOutputTokens: 1024,
        topP: 0.8,
        topK: 40
      }
    };
    
    const options = {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    console.log("Sending request to Gemini API...");
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    console.log("Response code: " + responseCode);
    
    if (responseCode === 200) {
      try {
        const parsedResponse = JSON.parse(responseText);
        console.log("Response parsed successfully");
        
        // Check if we have candidates in the response
        if (!parsedResponse.candidates || parsedResponse.candidates.length === 0) {
          console.error("No candidates in response: " + JSON.stringify(parsedResponse));
          return {
            success: false,
            text: "The AI model didn't generate a response. Please try again with a different question.",
            error: "No candidates in response"
          };
        }
        
        // Extract the text from Gemini's response
        const candidate = parsedResponse.candidates[0];
        if (!candidate.content || !candidate.content.parts || candidate.content.parts.length === 0) {
          console.error("No content in response candidate: " + JSON.stringify(candidate));
          return {
            success: false,
            text: "The AI model response was empty. Please try again.",
            error: "No content in response"
          };
        }
        
        const aiResponseText = candidate.content.parts[0].text;
        
        return {
          success: true,
          text: aiResponseText
        };
      } catch (parseError) {
        console.error("Error parsing response: " + parseError.toString());
        console.error("Response text: " + responseText);
        return {
          success: false,
          text: "Error parsing the AI response. Please check your API key or try again later.",
          error: parseError.toString()
        };
      }
    } else {
      console.error("API Error: " + responseText);
      
      // Attempt to parse error for more helpful message
      try {
        const errorJson = JSON.parse(responseText);
        const errorMessage = errorJson.error?.message || "Unknown API error";
        const errorStatus = errorJson.error?.status || "UNKNOWN";
        
        // Handle specific error types
        if (errorStatus.includes("INVALID_ARGUMENT")) {
          return {
            success: false,
            text: "Your API key may be invalid or there's an issue with the request. Please check your settings.",
            error: errorMessage
          };
        } else if (errorStatus.includes("PERMISSION_DENIED")) {
          return {
            success: false,
            text: "Your API key doesn't have permission to use this model. Please check your Google AI Studio account.",
            error: errorMessage
          };
        } else if (errorStatus.includes("RESOURCE_EXHAUSTED")) {
          return {
            success: false,
            text: "You've reached your API usage limit. Please try again later or check your quota in Google AI Studio.",
            error: errorMessage
          };
        } else {
          return {
            success: false,
            text: "Sorry, I encountered an error: " + errorMessage,
            error: errorMessage
          };
        }
      } catch (e) {
        return {
          success: false,
          text: "Sorry, there was an error contacting the AI service. Please try again later.",
          error: responseText
        };
      }
    }
    
  } catch (error) {
    console.error("Error in callGeminiAPI: " + error.toString());
    return {
      success: false,
      text: "Sorry, something went wrong when connecting to the AI service. Please try again later.",
      error: error.toString()
    };
  }
}

/**
 * Helper function to format sample data in a readable way
 */
function formatSampleData(rows, headers) {
  if (!rows || rows.length === 0) return "No data available";
  
  let result = "";
  rows.forEach((row, index) => {
    result += `Row ${index + 1}: `;
    headers.forEach((header, colIndex) => {
      result += `${header}: ${row[colIndex]}, `;
    });
    result = result.slice(0, -2); // Remove trailing comma and space
    result += "\n";
  });
  
  return result;
}

/**
 * Processes the user's message and returns an appropriate response.
 * This function handles command detection and tool usage.
 */
function processUserMessage(message, history = []) {
  try {
    // Check for special commands
    if (message.toLowerCase().startsWith("/help")) {
      return {
        success: true,
        text: "SheetSense Help:\n\n" +
              "- Ask me about your spreadsheet data\n" +
              "- Request charts and visualizations\n" +
              "- Get formulas for calculations\n" +
              "- Analyze trends in your data\n" +
              "- Use /clear to clear the chat history"
      };
    }
    
    if (message.toLowerCase() === "/clear") {
      // Clear chat history handled by the frontend
      return {
        success: true,
        text: "Chat history cleared.",
        action: "clear_history"
      };
    }
    
    // Get spreadsheet data for context
    try {
      const sheetData = getSpreadsheetData();
      
      // Call Gemini API with message and context
      const context = {
        sheetData: sheetData,
        currentTime: new Date().toString()
      };
      
      return callGeminiAPI(message, context, history);
    } catch (dataError) {
      console.error("Error getting spreadsheet data: " + dataError.toString());
      return {
        success: false,
        text: "I couldn't access your spreadsheet data. Please make sure your sheet contains data with headers in the first row.",
        error: dataError.toString()
      };
    }
    
  } catch (error) {
    console.error("Error in processUserMessage: " + error.toString());
    return {
      success: false,
      text: "Sorry, I encountered an error processing your message. Please try again.",
      error: error.toString()
    };
  }
}

/**
 * Saves chat history to user properties
 */
function saveChatHistory(history) {
  try {
    const settings = getSettings();
    if (!settings.saveHistory) return { success: true };
    
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('chatHistory', JSON.stringify(history));
    return { success: true };
  } catch (error) {
    console.error("Error saving chat history: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Retrieves saved chat history from user properties
 */
function getChatHistory() {
  try {
    const settings = getSettings();
    if (!settings.saveHistory) return { success: true, history: [] };
    
    const userProperties = PropertiesService.getUserProperties();
    const historyString = userProperties.getProperty('chatHistory');
    
    if (!historyString) return { success: true, history: [] };
    
    return { 
      success: true, 
      history: JSON.parse(historyString)
    };
  } catch (error) {
    console.error("Error retrieving chat history: " + error.toString());
    return { success: false, history: [], error: error.toString() };
  }
}
