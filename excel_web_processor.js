let workbook = null;
let processedWorkbook = null;
let originalFileName = '';
let sheetColumns = {};

// File upload handling
const fileInput = document.getElementById('fileInput');
const uploadSection = document.getElementById('uploadSection');

fileInput.addEventListener('change', handleFileSelect);

// Drag and drop functionality
uploadSection.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadSection.classList.add('dragover');
});

uploadSection.addEventListener('dragleave', () => {
    uploadSection.classList.remove('dragover');
});

uploadSection.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadSection.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
});

function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) {
        handleFile(file);
    }
}

function handleFile(file) {
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
        showError('Please select a valid Excel file (.xlsx or .xls)');
        return;
    }
    originalFileName = file.name.replace(/\.(xlsx|xls)$/i, '');
    showSuccess(`File "${file.name}" uploaded successfully!`);
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            workbook = XLSX.read(data, { type: 'array' });
            displaySheetInfo();
            setupSheetMapping();
            document.getElementById('sheetMappingSection').style.display = 'block';
        } catch (error) {
            showError('Error reading Excel file: ' + error.message);
        }
    };
    reader.readAsArrayBuffer(file);
}

function displaySheetInfo() {
    const sheetInfo = document.getElementById('sheetInfo');
    let html = '<h3>📑 Available Sheets:</h3>';
    workbook.SheetNames.forEach((sheetName, index) => {
        const sheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(sheet);
        const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null;
        const rows = range ? range.e.r + 1 : 0;
        const cols = range ? range.e.c + 1 : 0;
        // Store column names for each sheet
        sheetColumns[sheetName] = data.length > 0 ? Object.keys(data[0]) : [];
        html += `
            <div class="sheet-card">
                <div class="sheet-name">📄 ${sheetName}</div>
                <div class="sheet-stats">${rows} rows × ${cols} columns | ${data.length} data rows</div>
            </div>
        `;
    });
    sheetInfo.innerHTML = html;
}

function setupSheetMapping() {
    const sheetSelects = ['ordersSheet', 'inventorySheet', 'prioritySheet'];
    // Populate sheet dropdown options
    sheetSelects.forEach(selectId => {
        const select = document.getElementById(selectId);
        select.innerHTML = '<option value="">-- Select Sheet --</option>';
        workbook.SheetNames.forEach(sheetName => {
            select.innerHTML += `<option value="${sheetName}">${sheetName}</option>`;
        });
        // Add change event listener
        select.addEventListener('change', function() {
            updateColumnOptions(selectId, this.value);
            checkAllMappingsComplete();
        });
    });
    
    // Add change event listener for global preference
    document.getElementById('globalPreference').addEventListener('change', checkAllMappingsComplete);
}

function updateColumnOptions(sheetSelectId, sheetName) {
    if (!sheetName) return;
    const columns = sheetColumns[sheetName] || [];
    let columnSelects = [];
    // Determine which column selects to update based on sheet type
    switch(sheetSelectId) {
        case 'ordersSheet':
            columnSelects = ['ordersStoreCol', 'ordersItemCol', 'ordersQuantityCol'];
            break;
        case 'inventorySheet':
            columnSelects = ['inventoryItemCol', 'inventoryRetailQuantityCol', 'inventoryReturnQuantityCol'];
            break;
        case 'prioritySheet':
            columnSelects = ['priorityStoreCol', 'priorityLevelCol'];
            break;
    }
    // Update column dropdown options
    columnSelects.forEach(selectId => {
        const select = document.getElementById(selectId);
        select.innerHTML = '<option value="">-- Select Column --</option>';
        columns.forEach(column => {
            select.innerHTML += `<option value="${column}">${column}</option>`;
        });
        // Add change event listener
        select.addEventListener('change', checkAllMappingsComplete);
    });
}

function checkAllMappingsComplete() {
    const requiredMappings = [
        'ordersSheet', 'ordersStoreCol', 'ordersItemCol', 'ordersQuantityCol',
        'inventorySheet', 'inventoryItemCol', 'inventoryRetailQuantityCol', 'inventoryReturnQuantityCol',
        'prioritySheet', 'priorityStoreCol', 'priorityLevelCol',
        'globalPreference'
    ];
    const allComplete = requiredMappings.every(id => {
        const element = document.getElementById(id);
        return element && element.value !== '';
    });
    document.getElementById('processBtn').disabled = !allComplete;
}

function processAllocation() {
    showProgress(0);
    document.getElementById('progressBar').style.display = 'block';
    setTimeout(() => {
        try {
            // Get mapping values
            const mapping = {
                orders: {
                    sheet: document.getElementById('ordersSheet').value,
                    storeCol: document.getElementById('ordersStoreCol').value,
                    itemCol: document.getElementById('ordersItemCol').value,
                    quantityCol: document.getElementById('ordersQuantityCol').value
                },
                inventory: {
                    sheet: document.getElementById('inventorySheet').value,
                    itemCol: document.getElementById('inventoryItemCol').value,
                    retailQuantityCol: document.getElementById('inventoryRetailQuantityCol').value,
                    returnQuantityCol: document.getElementById('inventoryReturnQuantityCol').value
                },
                priority: {
                    sheet: document.getElementById('prioritySheet').value,
                    storeCol: document.getElementById('priorityStoreCol').value,
                    levelCol: document.getElementById('priorityLevelCol').value
                },
                globalPreference: parseInt(document.getElementById('globalPreference').value)
            };
            showProgress(20);
            // Extract data from sheets
            const ordersData = XLSX.utils.sheet_to_json(workbook.Sheets[mapping.orders.sheet]);
            const inventoryData = XLSX.utils.sheet_to_json(workbook.Sheets[mapping.inventory.sheet]);
            const priorityData = XLSX.utils.sheet_to_json(workbook.Sheets[mapping.priority.sheet]);
            showProgress(40);
            // Process allocation
            const allocationResult = calculateAllocation(ordersData, inventoryData, priorityData, mapping);
            showProgress(80);
            // Create output workbook
            processedWorkbook = XLSX.utils.book_new();
            // Add allocation results
            const allocationSheet = XLSX.utils.json_to_sheet(allocationResult.allocations);
            XLSX.utils.book_append_sheet(processedWorkbook, allocationSheet, 'Item_Allocation');
            // Add summary sheet
            const summarySheet = XLSX.utils.json_to_sheet([allocationResult.summary]);
            XLSX.utils.book_append_sheet(processedWorkbook, summarySheet, 'Allocation_Summary');
            // Add unmet orders sheet
            if (allocationResult.unmetOrders.length > 0) {
                const unmetSheet = XLSX.utils.json_to_sheet(allocationResult.unmetOrders);
                XLSX.utils.book_append_sheet(processedWorkbook, unmetSheet, 'Unmet_Orders');
            }
            showProgress(100);
            // Display summary
            displayAllocationSummary(allocationResult.summary);
            setTimeout(() => {
                document.getElementById('progressBar').style.display = 'none';
                document.getElementById('downloadSection').style.display = 'block';
                showSuccess('Allocation calculation completed successfully!');
            }, 500);
        } catch (error) {
            showError('Error processing allocation: ' + error.message);
            document.getElementById('progressBar').style.display = 'none';
        }
    }, 100);
}

function calculateAllocation(ordersData, inventoryData, priorityData, mapping) {
    // Create inventory lookup with both quantity sources
    const inventory = {};
    inventoryData.forEach(row => {
        const itemId = row[mapping.inventory.itemCol];
        const retailQuantity = parseInt(row[mapping.inventory.retailQuantityCol]) || 0;
        const returnQuantity = parseInt(row[mapping.inventory.returnQuantityCol]) || 0;
        inventory[itemId] = {
            retailLocation: retailQuantity,
            retailReturnLocation: returnQuantity,
            total: retailQuantity + returnQuantity
        };
    });
    
    // Create priority lookup
    const storePriorities = {};
    priorityData.forEach(row => {
        const storeId = row[mapping.priority.storeCol];
        const priority = parseInt(row[mapping.priority.levelCol]) || 999;
        storePriorities[storeId] = priority;
    });
    
    // Get global preference
    const globalPreference = mapping.globalPreference;
    
    // Process orders by priority
    const ordersByPriority = ordersData
        .map(row => ({
            storeId: row[mapping.orders.storeCol],
            itemId: row[mapping.orders.itemCol],
            requestedQuantity: parseInt(row[mapping.orders.quantityCol]) || 0,
            priority: storePriorities[row[mapping.orders.storeCol]] || 999
        }))
        .sort((a, b) => a.priority - b.priority);
    
    const allocations = [];
    const unmetOrders = [];
    let totalRequested = 0;
    let totalAllocated = 0;
    let totalUnmet = 0;
    let totalFromRetailLocation = 0;
    let totalFromReturnLocation = 0;
    
    // Process each order
    ordersByPriority.forEach(order => {
        totalRequested += order.requestedQuantity;
        const itemInventory = inventory[order.itemId] || { retailLocation: 0, retailReturnLocation: 0, total: 0 };
        
        let remainingRequested = order.requestedQuantity;
        let allocatedFromRetail = 0;
        let allocatedFromReturn = 0;
        
        // Determine allocation order based on global preference
        const allocationOrder = globalPreference === 1 ? 
            ['retailLocation', 'retailReturnLocation'] : 
            ['retailReturnLocation', 'retailLocation'];
        
        // Allocate from preferred source first
        for (const source of allocationOrder) {
            if (remainingRequested <= 0) break;
            
            const availableQuantity = itemInventory[source] || 0;
            const quantityToAllocate = Math.min(remainingRequested, availableQuantity);
            
            if (quantityToAllocate > 0) {
                if (source === 'retailLocation') {
                    allocatedFromRetail += quantityToAllocate;
                    itemInventory.retailLocation -= quantityToAllocate;
                } else {
                    allocatedFromReturn += quantityToAllocate;
                    itemInventory.retailReturnLocation -= quantityToAllocate;
                }
                remainingRequested -= quantityToAllocate;
            }
        }
        
        const totalAllocatedForOrder = allocatedFromRetail + allocatedFromReturn;
        
        if (totalAllocatedForOrder > 0) {
            allocations.push({
                Store_ID: order.storeId,
                Item_ID: order.itemId,
                Requested_Quantity: order.requestedQuantity,
                Allocated_Quantity: totalAllocatedForOrder,
                Allocated_From_Retail_Location: allocatedFromRetail,
                Allocated_From_Return_Location: allocatedFromReturn,
                Store_Priority: order.priority,
                Global_Preference: globalPreference,
                Allocation_Date: new Date().toISOString().split('T')[0]
            });
            
            totalAllocated += totalAllocatedForOrder;
            totalFromRetailLocation += allocatedFromRetail;
            totalFromReturnLocation += allocatedFromReturn;
        }
        
        // Record unmet portion
        if (remainingRequested > 0) {
            unmetOrders.push({
                Store_ID: order.storeId,
                Item_ID: order.itemId,
                Unmet_Quantity: remainingRequested,
                Store_Priority: order.priority,
                Global_Preference: globalPreference,
                Reason: itemInventory.total === 0 ? 'Out of Stock' : 'Insufficient Stock'
            });
            totalUnmet += remainingRequested;
        }
    });
    
    const summary = {
        Total_Orders_Processed: ordersByPriority.length,
        Total_Items_Requested: totalRequested,
        Total_Items_Allocated: totalAllocated,
        Total_Items_Unmet: totalUnmet,
        Total_Allocated_From_Retail_Location: totalFromRetailLocation,
        Total_Allocated_From_Return_Location: totalFromReturnLocation,
        Global_Preference_Used: globalPreference,
        Allocation_Rate_Percent: totalRequested > 0 ? Math.round((totalAllocated / totalRequested) * 100) : 0,
        Unique_Stores: new Set(ordersByPriority.map(o => o.storeId)).size,
        Unique_Items: new Set(ordersByPriority.map(o => o.itemId)).size,
        Processing_Date: new Date().toISOString()
    };
    
    return {
        allocations,
        unmetOrders,
        summary
    };
}

function displayAllocationSummary(summary) {
    const summaryGrid = document.getElementById('summaryGrid');
    const preferenceText = summary.Global_Preference_Used === 1 ? 
        'Retail Location First' : 'Return Location First';
    
    summaryGrid.innerHTML = `
        <div class="summary-card">
            <div class="summary-number">${summary.Total_Orders_Processed}</div>
            <div class="summary-label">Orders Processed</div>
        </div>
        <div class="summary-card">
            <div class="summary-number">${summary.Total_Items_Allocated}</div>
            <div class="summary-label">Items Allocated</div>
        </div>
        <div class="summary-card">
            <div class="summary-number">${summary.Allocation_Rate_Percent}%</div>
            <div class="summary-label">Allocation Rate</div>
        </div>
        <div class="summary-card">
            <div class="summary-number">${summary.Unique_Stores}</div>
            <div class="summary-label">Stores Served</div>
        </div>
        <div class="summary-card">
            <div class="summary-number">${summary.Total_Allocated_From_Retail_Location}</div>
            <div class="summary-label">From Retail Location</div>
        </div>
        <div class="summary-card">
            <div class="summary-number">${summary.Total_Allocated_From_Return_Location}</div>
            <div class="summary-label">From Return Location</div>
        </div>
        <div class="summary-card">
            <div class="summary-number">${preferenceText}</div>
            <div class="summary-label">Preference Used</div>
        </div>
    `;
    document.getElementById('allocationSummary').style.display = 'block';
}

function downloadFile() {
    if (!processedWorkbook) {
        showError('No processed data available for download');
        return;
    }
    try {
        const fileName = `${originalFileName}_allocation_${new Date().toISOString().slice(0, 10)}.xlsx`;
        XLSX.writeFile(processedWorkbook, fileName);
        showSuccess(`File "${fileName}" downloaded successfully!`);
    } catch (error) {
        showError('Error downloading file: ' + error.message);
    }
}

function showProgress(percent) {
    document.getElementById('progressFill').style.width = percent + '%';
}

function showError(message) {
    const errorDiv = document.getElementById('errorMessage');
    errorDiv.textContent = message;
    errorDiv.style.display = 'block';
    document.getElementById('successMessage').style.display = 'none';
}

function showSuccess(message) {
    const successDiv = document.getElementById('successMessage');
    successDiv.textContent = message;
    successDiv.style.display = 'block';
    document.getElementById('errorMessage').style.display = 'none';
}