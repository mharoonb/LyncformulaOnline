/**
 * LyncFormula Task Pane JavaScript
 * Handles UI interactions, service communication, and Excel integration
 */

// Global variables
let serviceUrl = 'http://127.0.0.1:8700';
let connectionStatus = 'connecting';
let selectedFiles = [];
let operationHistory = [];
let currentOperation = null;

// Initialize the task pane
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.addEventListener('DOMContentLoaded', initializeTaskPane);
    }
});

/**
 * Initialize the task pane interface
 */
function initializeTaskPane() {
    console.log('🚀 LyncFormula Task Pane initializing...');
    
    // Set up event listeners
    setupEventListeners();
    
    // Check service connection
    checkServiceConnection();
    
    // Load settings from storage
    loadSettings();
    
    // Load operation history
    loadHistory();
    
    // Enable input monitoring
    setupInputMonitoring();
    
    console.log('✅ Task pane initialized');
}

/**
 * Set up all event listeners
 */
function setupEventListeners() {
    // Main action button
    document.getElementById('btn-submit').addEventListener('click', handleSubmitRequest);
    
    // Quick action buttons
    document.getElementById('btn-audit-sheet').addEventListener('click', () => {
        document.getElementById('natural-language-input').value = 'Audit all formulas in the current sheet for errors and issues';
        handleSubmitRequest();
    });
    
    document.getElementById('btn-explain-formulas').addEventListener('click', () => {
        document.getElementById('natural-language-input').value = 'Explain all formulas in the current sheet in plain English';
        handleSubmitRequest();
    });
    
    document.getElementById('btn-consolidate-files').addEventListener('click', () => {
        document.getElementById('natural-language-input').value = 'Consolidate multiple Excel files into a single summary report';
        showFileSelection();
    });
    
    // File selection
    document.getElementById('btn-browse-files').addEventListener('click', browseFiles);
    
    // History management
    document.getElementById('btn-clear-history').addEventListener('click', clearHistory);
    
    // Settings
    document.getElementById('btn-settings').addEventListener('click', showSettings);
    document.getElementById('btn-close-settings').addEventListener('click', hideSettings);
    document.getElementById('btn-save-settings').addEventListener('click', saveSettings);
    
    // Help
    document.getElementById('btn-help').addEventListener('click', showHelp);
    
    // Input validation
    document.getElementById('natural-language-input').addEventListener('input', validateInput);
    document.getElementById('natural-language-input').addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && e.ctrlKey) {
            handleSubmitRequest();
        }
    });
}

/**
 * Set up input monitoring for better UX
 */
function setupInputMonitoring() {
    const input = document.getElementById('natural-language-input');
    let timeout;
    
    input.addEventListener('input', () => {
        clearTimeout(timeout);
        timeout = setTimeout(() => {
            const text = input.value.trim();
            
            // Show file selection if user mentions files
            if (text.toLowerCase().includes('file') || text.toLowerCase().includes('consolidate') || text.toLowerCase().includes('merge')) {
                if (selectedFiles.length === 0) {
                    showFileSelection();
                }
            }
            
            // Auto-suggest improvements
            if (text.length > 50) {
                suggestImprovements(text);
            }
        }, 1000);
    });
}

/**
 * Check service connection status
 */
async function checkServiceConnection() {
    const statusIndicator = document.getElementById('status-indicator');
    const statusText = document.getElementById('status-text');
    
    try {
        const response = await fetch(`${serviceUrl}/health`, {
            method: 'GET',
            timeout: 5000
        });
        
        if (response.ok) {
            const data = await response.json();
            connectionStatus = 'connected';
            statusIndicator.className = 'status-indicator connected';
            statusText.textContent = `Connected to ${data.service} v${data.version}`;
            
            // Enable submit button if there's input
            validateInput();
        } else {
            throw new Error(`Service returned ${response.status}`);
        }
    } catch (error) {
        console.error('Service connection failed:', error);
        connectionStatus = 'error';
        statusIndicator.className = 'status-indicator error';
        statusText.textContent = 'Service unavailable - Check if Python service is running';
        
        // Disable submit button
        document.getElementById('btn-submit').disabled = true;
    }
}

/**
 * Validate input and enable/disable submit button
 */
function validateInput() {
    const input = document.getElementById('natural-language-input');
    const submitBtn = document.getElementById('btn-submit');
    
    const hasInput = input.value.trim().length > 0;
    const isConnected = connectionStatus === 'connected';
    
    submitBtn.disabled = !hasInput || !isConnected;
}

/**
 * Handle main submit request
 */
async function handleSubmitRequest() {
    const input = document.getElementById('natural-language-input');
    const query = input.value.trim();
    
    if (!query) {
        showError('Please enter a request');
        return;
    }
    
    showLoading(true);
    
    try {
        // Prepare request data
        const requestData = {
            query: query,
            files: selectedFiles,
            context: await getCurrentExcelContext()
        };
        
        // Determine the appropriate endpoint based on query
        const endpoint = determineEndpoint(query);
        
        // Send request
        const response = await fetch(`${serviceUrl}${endpoint}`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(requestData)
        });
        
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        
        const result = await response.json();
        
        // Handle the response
        if (result.success) {
            showResults(result, query);
            addToHistory(query, result);
            
            // Clear input on successful operation
            input.value = '';
            validateInput();
            
            // Hide file selection if it was shown
            hideFileSelection();
        } else {
            showError(result.error || 'Operation failed');
        }
        
    } catch (error) {
        console.error('Request failed:', error);
        showError(`Request failed: ${error.message}`);
    } finally {
        showLoading(false);
    }
}

/**
 * Determine the appropriate endpoint based on the query
 */
function determineEndpoint(query) {
    const lowerQuery = query.toLowerCase();
    
    if (lowerQuery.includes('audit') || lowerQuery.includes('check') || lowerQuery.includes('analyze')) {
        if (selectedFiles.length > 0 || lowerQuery.includes('file')) {
            return '/smart_analysis';
        }
        return '/audit_formulas';
    }
    
    if (lowerQuery.includes('explain') || lowerQuery.includes('describe')) {
        return '/explain_sheet';
    }
    
    if (lowerQuery.includes('consolidate') || lowerQuery.includes('merge') || lowerQuery.includes('combine')) {
        return '/consolidate_files';
    }
    
    if (lowerQuery.includes('fix') || lowerQuery.includes('repair')) {
        return '/apply_fixes';
    }
    
    // Default to natural language processing
    return '/natural_language';
}

/**
 * Get current Excel context for better AI understanding
 */
async function getCurrentExcelContext() {
    return new Promise((resolve) => {
        Excel.run(async (context) => {
            try {
                const worksheet = context.workbook.worksheets.getActiveWorksheet();
                worksheet.load('name');
                
                const usedRange = worksheet.getUsedRange();
                usedRange.load('address, rowCount, columnCount');
                
                await context.sync();
                
                const contextInfo = {
                    worksheetName: worksheet.name,
                    usedRange: usedRange.address,
                    rowCount: usedRange.rowCount,
                    columnCount: usedRange.columnCount,
                    hasSelection: false
                };
                
                // Try to get selected range
                try {
                    const selectedRange = context.workbook.getSelectedRange();
                    selectedRange.load('address');
                    await context.sync();
                    
                    contextInfo.selectedRange = selectedRange.address;
                    contextInfo.hasSelection = true;
                } catch (e) {
                    // No selection, which is fine
                }
                
                resolve(contextInfo);
            } catch (error) {
                console.error('Error getting Excel context:', error);
                resolve({
                    worksheetName: 'Unknown',
                    error: error.message
                });
            }
        });
    });
}

/**
 * Show/hide loading overlay
 */
function showLoading(show) {
    const overlay = document.getElementById('loading-overlay');
    overlay.style.display = show ? 'flex' : 'none';
}

/**
 * Display results in the results section
 */
function showResults(result, originalQuery) {
    const resultsSection = document.getElementById('results-section');
    const resultsContent = document.getElementById('results-content');
    
    // Clear previous results
    resultsContent.innerHTML = '';
    
    // Create results display based on result type
    if (result.explanations) {
        showFormulaExplanations(result.explanations, resultsContent);
    } else if (result.issues) {
        showAuditResults(result, resultsContent);
    } else if (result.analysis_results) {
        showSmartAnalysis(result.analysis_results, resultsContent);
    } else if (result.consolidated_file) {
        showConsolidationResults(result, resultsContent);
    } else if (result.operation_type) {
        showNaturalLanguageResults(result, resultsContent);
    } else {
        showGenericResults(result, resultsContent);
    }
    
    // Show the results section
    resultsSection.scrollIntoView({ behavior: 'smooth' });
}

/**
 * Show formula explanations
 */
function showFormulaExplanations(explanations, container) {
    const title = document.createElement('h4');
    title.textContent = 'Formula Explanations';
    title.className = 'result-title';
    container.appendChild(title);
    
    explanations.forEach(explanation => {
        const item = document.createElement('div');
        item.className = 'result-item result-success';
        
        item.innerHTML = `
            <div class="result-header">
                <strong>${explanation.address}:</strong>
                <span class="result-complexity">${explanation.complexity}</span>
            </div>
            <div class="result-formula"><code>${explanation.formula}</code></div>
            <div class="result-content">${explanation.explanation}</div>
            ${explanation.functions && explanation.functions.length > 0 ? 
                `<div class="result-functions"><small>Functions: ${explanation.functions.join(', ')}</small></div>` : 
                ''}
        `;
        
        container.appendChild(item);
    });
}

/**
 * Show audit results
 */
function showAuditResults(result, container) {
    const summary = result.audit_summary || {};
    
    const title = document.createElement('h4');
    title.textContent = 'Formula Audit Results';
    container.appendChild(title);
    
    // Summary
    const summaryDiv = document.createElement('div');
    summaryDiv.className = 'result-item';
    summaryDiv.innerHTML = `
        <div class="result-header">
            <strong>Audit Summary</strong>
        </div>
        <div class="result-content">
            <p>Total Issues: ${summary.total_issues || 0}</p>
            <p>High Priority: ${summary.severity_breakdown?.high || 0}</p>
            <p>Medium Priority: ${summary.severity_breakdown?.medium || 0}</p>
            <p>Low Priority: ${summary.severity_breakdown?.low || 0}</p>
        </div>
    `;
    container.appendChild(summaryDiv);
    
    // Issues
    if (result.issues && result.issues.length > 0) {
        result.issues.slice(0, 5).forEach(issue => {
            const item = document.createElement('div');
            item.className = `result-item ${issue.severity === 'high' ? 'result-error' : 
                                            issue.severity === 'medium' ? 'result-warning' : 
                                            'result-success'}`;
            
            item.innerHTML = `
                <div class="result-header">
                    <strong>${issue.title}</strong>
                    <span class="severity-badge">${issue.severity}</span>
                </div>
                <div class="result-content">
                    <p>${issue.description}</p>
                    <p><strong>Cell:</strong> ${issue.cell_address}</p>
                    <p><strong>Suggested Fix:</strong> ${issue.suggested_fix}</p>
                </div>
            `;
            
            container.appendChild(item);
        });
    }
}

/**
 * Show smart analysis results
 */
function showSmartAnalysis(analysisResults, container) {
    const title = document.createElement('h4');
    title.textContent = 'Smart Formula Analysis';
    container.appendChild(title);
    
    analysisResults.forEach(analysis => {
        const item = document.createElement('div');
        item.className = `result-item ${analysis.has_errors ? 'result-error' : 'result-success'}`;
        
        item.innerHTML = `
            <div class="result-header">
                <strong>${analysis.address}</strong>
                <span class="confidence-badge">${analysis.confidence}% confidence</span>
            </div>
            <div class="result-content">
                <p><strong>Formula:</strong> <code>${analysis.original_formula}</code></p>
                <p><strong>Analysis:</strong> ${analysis.explanation}</p>
                ${analysis.has_errors ? `<p><strong>Issues:</strong> ${analysis.error_details}</p>` : ''}
                ${analysis.suggestions && analysis.suggestions.length > 0 ? 
                    `<p><strong>Suggestions:</strong> ${analysis.suggestions.join(', ')}</p>` : ''}
                ${analysis.improved_formula ? 
                    `<p><strong>Improved Formula:</strong> <code>${analysis.improved_formula}</code></p>` : ''}
            </div>
        `;
        
        container.appendChild(item);
    });
}

/**
 * Show consolidation results
 */
function showConsolidationResults(result, container) {
    const item = document.createElement('div');
    item.className = 'result-item result-success';
    
    item.innerHTML = `
        <div class="result-header">
            <strong>File Consolidation Complete</strong>
        </div>
        <div class="result-content">
            <p><strong>Output File:</strong> ${result.consolidated_file}</p>
            <p><strong>Rows Merged:</strong> ${result.rows_merged}</p>
            <p><strong>Files Processed:</strong> ${result.files_processed}</p>
        </div>
    `;
    
    container.appendChild(item);
}

/**
 * Show natural language operation results
 */
function showNaturalLanguageResults(result, container) {
    const item = document.createElement('div');
    item.className = 'result-item result-success';
    
    item.innerHTML = `
        <div class="result-header">
            <strong>${result.operation_type} Operation Complete</strong>
        </div>
        <div class="result-content">
            <p>${result.summary}</p>
            ${result.results && result.results.length > 0 ? 
                `<ul>${result.results.map(r => `<li>${r.result || r.action}</li>`).join('')}</ul>` : ''}
        </div>
    `;
    
    container.appendChild(item);
}

/**
 * Show generic results
 */
function showGenericResults(result, container) {
    const item = document.createElement('div');
    item.className = 'result-item';
    
    item.innerHTML = `
        <div class="result-header">
            <strong>Operation Result</strong>
        </div>
        <div class="result-content">
            <pre>${JSON.stringify(result, null, 2)}</pre>
        </div>
    `;
    
    container.appendChild(item);
}

/**
 * Show error message
 */
function showError(message) {
    const resultsContent = document.getElementById('results-content');
    
    resultsContent.innerHTML = `
        <div class="result-item result-error">
            <div class="result-header">
                <strong>Error</strong>
            </div>
            <div class="result-content">
                ${message}
            </div>
        </div>
    `;
    
    document.getElementById('results-section').scrollIntoView({ behavior: 'smooth' });
}

/**
 * File selection functionality
 */
function showFileSelection() {
    document.getElementById('file-selection-section').style.display = 'block';
}

function hideFileSelection() {
    document.getElementById('file-selection-section').style.display = 'none';
}

async function browseFiles() {
    // In a real implementation, this would use the Office File Picker API
    // For now, we'll simulate file selection
    
    const mockFiles = [
        { name: 'Q1_Sales.xlsx', path: 'C:\\Reports\\Q1_Sales.xlsx' },
        { name: 'Q2_Sales.xlsx', path: 'C:\\Reports\\Q2_Sales.xlsx' },
        { name: 'Budget_2024.xlsx', path: 'C:\\Finance\\Budget_2024.xlsx' }
    ];
    
    // Add mock files to selection
    mockFiles.forEach(file => {
        if (!selectedFiles.find(f => f.path === file.path)) {
            selectedFiles.push(file);
        }
    });
    
    updateSelectedFilesList();
}

function updateSelectedFilesList() {
    const filesList = document.getElementById('selected-files');
    
    filesList.innerHTML = selectedFiles.map((file, index) => `
        <div class="file-item">
            <div class="file-info">
                <div class="file-name">${file.name}</div>
                <div class="file-path">${file.path}</div>
            </div>
            <div class="file-remove" onclick="removeFile(${index})">
                <i class="ms-Icon ms-Icon--Cancel"></i>
            </div>
        </div>
    `).join('');
}

function removeFile(index) {
    selectedFiles.splice(index, 1);
    updateSelectedFilesList();
}

/**
 * History management
 */
function addToHistory(query, result) {
    const historyItem = {
        id: Date.now(),
        query: query,
        result: result,
        timestamp: new Date().toLocaleString()
    };
    
    operationHistory.unshift(historyItem);
    
    // Keep only last 20 operations
    if (operationHistory.length > 20) {
        operationHistory = operationHistory.slice(0, 20);
    }
    
    updateHistoryDisplay();
    saveHistory();
}

function updateHistoryDisplay() {
    const historyContent = document.getElementById('history-content');
    
    if (operationHistory.length === 0) {
        historyContent.innerHTML = '<p class="ms-fontColor-neutralSecondary">No operations yet</p>';
        return;
    }
    
    historyContent.innerHTML = operationHistory.map(item => `
        <div class="history-item" onclick="repeatOperation('${item.query}')">
            <div class="history-query">${item.query}</div>
            <div class="history-timestamp">${item.timestamp}</div>
        </div>
    `).join('');
}

function repeatOperation(query) {
    document.getElementById('natural-language-input').value = query;
    validateInput();
}

function clearHistory() {
    operationHistory = [];
    updateHistoryDisplay();
    saveHistory();
}

function loadHistory() {
    try {
        const stored = localStorage.getItem('lyncformula-history');
        if (stored) {
            operationHistory = JSON.parse(stored);
            updateHistoryDisplay();
        }
    } catch (error) {
        console.error('Error loading history:', error);
    }
}

function saveHistory() {
    try {
        localStorage.setItem('lyncformula-history', JSON.stringify(operationHistory));
    } catch (error) {
        console.error('Error saving history:', error);
    }
}

/**
 * Settings management
 */
function showSettings() {
    document.getElementById('settings-panel').style.display = 'block';
}

function hideSettings() {
    document.getElementById('settings-panel').style.display = 'none';
}

function loadSettings() {
    try {
        const stored = localStorage.getItem('lyncformula-settings');
        if (stored) {
            const settings = JSON.parse(stored);
            
            if (settings.serviceUrl) {
                serviceUrl = settings.serviceUrl;
                document.getElementById('service-url').value = serviceUrl;
            }
            
            if (settings.llmModel) {
                document.getElementById('llm-model').value = settings.llmModel;
            }
            
            if (settings.autoBackup !== undefined) {
                document.getElementById('auto-backup').checked = settings.autoBackup;
            }
        }
    } catch (error) {
        console.error('Error loading settings:', error);
    }
}

function saveSettings() {
    try {
        const settings = {
            serviceUrl: document.getElementById('service-url').value,
            llmModel: document.getElementById('llm-model').value,
            autoBackup: document.getElementById('auto-backup').checked
        };
        
        localStorage.setItem('lyncformula-settings', JSON.stringify(settings));
        
        // Update global serviceUrl
        serviceUrl = settings.serviceUrl;
        
        // Recheck connection with new settings
        checkServiceConnection();
        
        hideSettings();
        
        showNotification('Settings saved successfully');
    } catch (error) {
        console.error('Error saving settings:', error);
        showError('Failed to save settings');
    }
}

/**
 * Show help information
 */
function showHelp() {
    const helpContent = `
        <div class="result-item">
            <div class="result-header">
                <strong>LyncFormula Help</strong>
            </div>
            <div class="result-content">
                <h4>How to Use LyncFormula:</h4>
                <ul>
                    <li><strong>Natural Language:</strong> Type what you want to do in plain English</li>
                    <li><strong>Formula Analysis:</strong> "Audit all formulas in this sheet"</li>
                    <li><strong>File Operations:</strong> "Consolidate my quarterly sales files"</li>
                    <li><strong>Explanations:</strong> "Explain the formulas in column A"</li>
                </ul>
                
                <h4>Example Requests:</h4>
                <ul>
                    <li>"Find all errors in my formulas"</li>
                    <li>"Merge all expense reports into one file"</li>
                    <li>"Show me which formulas are causing performance issues"</li>
                    <li>"Convert my VLOOKUP formulas to XLOOKUP"</li>
                </ul>
                
                <h4>Features:</h4>
                <ul>
                    <li>🔒 100% Private - All processing happens locally</li>
                    <li>🤖 AI-Powered - Uses local LLM for intelligent analysis</li>
                    <li>📊 Smart Analysis - Detects real formula errors</li>
                    <li>💬 Conversational - Ask follow-up questions</li>
                </ul>
            </div>
        </div>
    `;
    
    document.getElementById('results-content').innerHTML = helpContent;
    document.getElementById('results-section').scrollIntoView({ behavior: 'smooth' });
}

/**
 * Show notification
 */
function showNotification(message) {
    // Create a temporary notification
    const notification = document.createElement('div');
    notification.className = 'notification';
    notification.textContent = message;
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background-color: var(--success-color);
        color: white;
        padding: 12px 20px;
        border-radius: 4px;
        z-index: 9999;
        animation: fadeIn 0.3s ease-in;
    `;
    
    document.body.appendChild(notification);
    
    setTimeout(() => {
        notification.remove();
    }, 3000);
}

/**
 * Auto-suggest improvements based on input
 */
function suggestImprovements(text) {
    // This could be enhanced with more sophisticated NLP
    const suggestions = [];
    
    if (text.toLowerCase().includes('file') && !text.toLowerCase().includes('consolidate')) {
        suggestions.push('Consider: "Consolidate these files into a summary"');
    }
    
    if (text.toLowerCase().includes('error') && !text.toLowerCase().includes('audit')) {
        suggestions.push('Consider: "Audit formulas for errors and issues"');
    }
    
    // Show suggestions if any
    if (suggestions.length > 0) {
        // Could implement a suggestions dropdown here
        console.log('Suggestions:', suggestions);
    }
}

// Utility functions
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

// Error handling
window.addEventListener('error', (event) => {
    console.error('Task pane error:', event.error);
    showError(`Unexpected error: ${event.error.message}`);
});

// Connection monitoring
setInterval(checkServiceConnection, 30000); // Check every 30 seconds