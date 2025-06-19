document.addEventListener('DOMContentLoaded', () => {
// This URL is for your Canvassing Data sheet. Ensure it's correct and published as CSV.
const DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=696550092&single=true&output=csv"; 
// IMPORTANT: Replace this with YOUR DEPLOYED GOOGLE APPS SCRIPT WEB APP URL
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzEYf0CKgwP0O4-z1lup1lDZImD1dQVEveLWsHwa_7T5ltndfIuRWXVZqFDj03_proD/exec"; // <-- PASTE YOUR NEWLY DEPLOYED WEB APP URL HERE
// For front-end reporting, all employee and branch data will come from Canvassing Data and predefined list.
    const EMPLOYEE_MASTER_DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=2120288173&single=true&output=csv"; // Marked as UNUSED for clarity, won't be fetched for reports

    const MONTHLY_WORKING_DAYS = 22; // Common approximation for a month's working days

    const TARGETS = {
        'Branch Manager': {
            'Visit': 10,
            'Call': 3 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 20
        },
        'Investment Staff': { // Added Investment Staff with custom Visit target
            'Visit': 15,
            'Call': 3 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 20
        },
        'Seniors': { // Added Investment Staff with custom Visit target
            'Visit': 15,
            'Call': 3 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 20
        },
        'Default': { // For all other designations not explicitly defined
            'Visit': 5,
            'Call': 3 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 20
        }
    };
    // Predefined list of branches for the dropdown and "no participation" check
    const PREDEFINED_BRANCHES = [
        "Angamaly", "Corporate Office", "Edappally", "Harippad", "Koduvayur", "Kuzhalmannam",
        "Mattanchery", "Mavelikara", "Nedumkandom", "Nenmara", "Paravoor", "Perumbavoor",
        "Thiruwillamala", "Thodupuzha", "Chengannur", "Alathur", "Kottayam", "Kattapana",
        "Muvattupuzha", "Thiruvalla", "Pathanamthitta", "HO KKM" // Corrected "Pathanamthitta" typo if it existed previously
    ].sort();

    // --- Column Headers Mapping (IMPORTANT: These must EXACTLY match the column names in your "Form Responses 2" Google Sheet) ---
    const HEADER_TIMESTAMP = 'Timestamp';
    const HEADER_DATE = 'Date';
    const HEADER_BRANCH_NAME = 'Branch Name';
    const HEADER_EMPLOYEE_NAME = 'Employee Name';
    const HEADER_EMPLOYEE_CODE = 'Employee Code';
    const HEADER_DESIGNATION = 'Designation';
    const HEADER_ACTIVITY_TYPE = 'Activity Type';
    const HEADER_TYPE_OF_CUSTOMER = 'Type of Customer'; // !!! CORRECTED TYPO HERE !!!
    const HEADER_R_LEAD_SOURCE = 'rLead Source';      // Keeping user's provided interpretation of split header
    const HEADER_HOW_CONTACTED = 'How Contacted'; // This is not in the list provided by user, but is in the original script. Keeping it.
    const HEADER_PROSPECT_NAME = 'Prospect Name';
    const HEADER_PHONE_NUMBER_WHATSAPP = 'Phone Numebr(Whatsapp)'; // Keeping user's provided typo
    const HEADER_ADDRESS = 'Address';
    const HEADER_PROFESSION = 'Profession';
    const HEADER_DOB_WD = 'DOB/WD';
    const HEADER_PRODUCT_INTERESTED = 'Prodcut Interested'; // Keeping user's provided typo
    const HEADER_REMARKS = 'Remarks';
    const HEADER_NEXT_FOLLOW_UP_DATE = 'Next Follow-up Date';
    const HEADER_RELATION_WITH_STAFF = 'Relation With Staff';
    // NEW: Customer Detail Headers as provided by user
    const HEADER_FAMILY_DETAILS_1 = 'Family Deatils -1 Name of wife/Husband';
    const HEADER_FAMILY_DETAILS_2 = 'Family Deatils -2 Job of wife/Husband';
    const HEADER_FAMILY_DETAILS_3 = 'Family Deatils -3 Names of Children';
    const HEADER_FAMILY_DETAILS_4 = 'Family Deatils -4 Deatils of Children';
    const HEADER_PROFILE_OF_CUSTOMER = 'Profile of Customer';

// ------------------------
const MASTER_HEADER_EMPLOYEE_CODE = 'Employee Code'; // Must match your Master Employee Sheet column header
const MASTER_HEADER_EMPLOYEE_NAME = 'Employee Name'; // Must match your Master Employee Sheet column header
const MASTER_HEADER_DESIGNATION = 'Designation';     // Must match your Master Employee Sheet column header
const MASTER_HEADER_BRANCH = 'Branch';               // Must match your Master Employee Sheet column header
// Core Display and Status Elements (existing)
const reportDisplay = document.getElementById('reportDisplay');
const statusMessage = document.getElementById('statusMessage');

// Main Content Sections to toggle (existing)
const reportsSection = document.getElementById('reportsSection');
const detailedCustomerViewSection = document.getElementById('detailedCustomerViewSection');
const employeeManagementSection = document.getElementById('employeeManagementSection');
// NEW: Performance Summary Dashboard Section
const performanceSummarySection = document.getElementById('performanceSummarySection');

// Tab buttons for main navigation (existing)
const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');
const nonParticipatingBranchesTabBtn = document.getElementById('nonParticipatingBranchesTabBtn');
const branchPerformanceTabBtn = document.getElementById('branchPerformanceTabBtn');
const detailedCustomerViewTabBtn = document.getElementById('detailedCustomerViewTabBtn');
const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');
// NEW: Performance Summary Tab Button
const performanceSummaryTabBtn = document.getElementById('performanceSummaryTabBtn');


// Dropdowns & Filter Panels (existing)
const branchSelect = document.getElementById('branchSelect');
const employeeSelect = document.getElementById('employeeSelect');
const employeeFilterPanel = document.getElementById('employeeFilterPanel');
const customerViewBranchSelect = document.getElementById('customerViewBranchSelect');
const customerViewEmployeeSelect = document.getElementById('customerViewEmployeeSelect');
// NEW: Dashboard Branch Select
const dashboardBranchSelect = document.getElementById('dashboardBranchSelect');

// View Option Buttons (from your provided list - ensure these are present in your script)
const viewBranchPerformanceReportBtn = document.getElementById('viewBranchPerformanceReportBtn');
const viewEmployeeSummaryBtn = document.getElementById('viewEmployeeSummaryBtn');
const viewAllEntriesBtn = document.getElementById('viewAllEntriesBtn');
const viewPerformanceReportBtn = document.getElementById('viewPerformanceReportBtn');
const viewBranchVisitLeaderboardBtn = document.getElementById('viewBranchVisitLeaderboardBtn');
const viewBranchCallLeaderboardBtn = document.getElementById('viewBranchCallLeaderboardBtn');
const viewStaffParticipationBtn = document.getElementById('viewStaffParticipationBtn');


// Detailed Customer View Specific Elements (existing)
const customerCanvassedList = document.getElementById('customerCanvassedList');
const customerDetailsContent = document.getElementById('customerDetailsContent');
const customerCard1 = document.getElementById('customerCard1');
const customerCard2 = document.getElementById('customerCard2');
const customerCard3 = document.getElementById('customerCard3');


// Employee Management Form Elements (existing)
const addEmployeeForm = document.getElementById('addEmployeeForm');
const newEmployeeNameInput = document.getElementById('newEmployeeName');
const newEmployeeCodeInput = document.getElementById('newEmployeeCode');
const newBranchNameInput = document.getElementById('newBranchName');
const newDesignationInput = document.getElementById('newDesignation');
const bulkAddEmployeeForm = document.getElementById('bulkAddEmployeeForm');
const bulkEmployeeBranchNameInput = document.getElementById('bulkEmployeeBranchName');
const bulkEmployeeDetailsTextarea = document.getElementById('bulkEmployeeDetails');
const deleteEmployeeForm = document.getElementById('deleteEmployeeForm');
const deleteEmployeeCodeInput = document.getElementById('deleteEmployeeCode');
const employeeManagementMessage = document.getElementById('employeeManagementMessage');


// Download Buttons (existing)
const downloadDetailedCustomerReportBtn = document.getElementById('downloadDetailedCustomerReportBtn'); // If you have this button
const downloadOverallStaffPerformanceReportBtn = document.getElementById('downloadOverallStaffPerformanceReportBtn');
// NEW: Performance Summary Dashboard Download Button
const downloadPerformanceSummaryReportBtn = document.getElementById('downloadPerformanceSummaryReportBtn');

// Global variables to store fetched data
let allCanvassingData = []; // Raw activity data from Form Responses 2
// NEW: Master Employee List Data (will remain empty for now as per your request, until you decide to integrate it)
let allMasterEmployees = []; // Data from the Master Employee List sheet
let allUniqueBranches = []; // Will be populated from PREDEFINED_BRANCHES or canvassing data
let allUniqueEmployees = []; // Employee codes from Canvassing Data
let employeeCodeToNameMap = {}; // {code: name} from Canvassing Data
let employeeCodeToDesignationMap = {}; // {code: designation} from Canvassing Data
let selectedBranchEntries = []; // Activity entries filtered by branch (for main reports section)
let selectedEmployeeCodeEntries = []; // Activity entries filtered by employee code (for main reports section)



    // Utility to format date to ISO-MM-DD
    const formatDate = (dateString) => {
        if (!dateString) return '';
        const date = new Date(dateString);
        if (isNaN(date.getTime())) return dateString;
        return date.toISOString().split('T')[0];
    };

    // Helper to clear and display messages in a specific div (now targets statusMessageDiv)
    // Helper to clear and display messages in a specific div
function displayMessage(message, type = 'info') {
    // Corrected: Use 'statusMessage' which is declared at the top of DOMContentLoaded
    if (statusMessage) { // Ensure the element exists
        statusMessage.innerHTML = `<div class="message ${type}">${message}</div>`;
        statusMessage.style.display = 'block';
        setTimeout(() => {
            statusMessage.innerHTML = ''; // Clear message
            statusMessage.style.display = 'none';
        }, 5000); // Hide after 5 seconds
    } else {
        console.error("Error: 'statusMessage' element not found in the DOM.");
    }
}

    // Specific message display for employee management forms
    function displayEmployeeManagementMessage(message, isError = false) {
        if (employeeManagementMessage) {
            employeeManagementMessage.textContent = message;
            employeeManagementMessage.style.color = isError ? 'red' : 'green';
            employeeManagementMessage.style.display = 'block';
            setTimeout(() => {
                employeeManagementMessage.style.display = 'none';
                employeeManagementMessage.textContent = ''; // Clear content
            }, 5000);
        }
    }

    // Function to fetch activity data from Google Sheet (Form Responses 2)
    async function fetchCanvassingData() {
        displayMessage("Fetching activity data...", 'info');
        try {
            const response = await fetch(DATA_URL);
            if (!response.ok) {
                const errorText = await response.text();
                console.error(`HTTP error fetching Canvassing Data! Status: ${response.status}. Details: ${errorText}`);
                throw new Error(`Failed to fetch canvassing data. Status: ${response.status}. Please check DATA_URL.`);
            }
            const csvText = await response.text();
            allCanvassingData = parseCSV(csvText);
            console.log('--- Fetched Canvassing Data: ---');
            console.log(allCanvassingData); // Log canvassing data for debugging
            if (allCanvassingData.length > 0) {
                console.log('Canvassing Data Headers (first entry):', Object.keys(allCanvassingData[0]));
            }
            displayMessage("Activity data loaded successfully!", 'success');
        } catch (error) {
            console.error('Error fetching canvassing data:', error);
            displayMessage(`Failed to load activity data: ${error.message}. Please ensure the sheet is published correctly to CSV and the URL is accurate.`, 'error');
            allCanvassingData = [];
        }
    }

    // CSV parsing function (handles commas within quoted strings)
    function parseCSV(csv) {
        const lines = csv.split('\n').filter(line => line.trim() !== '');
        if (lines.length === 0) return [];

        const headers = parseCSVLine(lines[0]); // Headers can also contain commas in quotes
        const data = [];

        for (let i = 1; i < lines.length; i++) {
            const values = parseCSVLine(lines[i]);
            if (values.length !== headers.length) {
                console.warn(`Skipping malformed row ${i + 1}: Expected ${headers.length} columns, got ${values.length}. Line: "${lines[i]}"`);
                continue;
            }
            const entry = {};
            headers.forEach((header, index) => {
                entry[header] = values[index];
            });
            data.push(entry);
        }
        return data;
    }

    // Helper to parse a single CSV line safely
    function parseCSVLine(line) {
        const result = [];
        let inQuote = false;
        let currentField = '';
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            if (char === '"') {
                inQuote = !inQuote;
            } else if (char === ',' && !inQuote) {
                result.push(currentField.trim());
                currentField = '';
            } else {
                currentField += char;
            }
        }
        result.push(currentField.trim());
        return result;
    }


    // Process fetched data to populate filters and prepare for reports
    async function processData() {
        // Only fetch canvassing data, ignoring MasterEmployees for front-end reports
        await fetchCanvassingData(); 

        // Re-initialize allUniqueBranches from the predefined list
        allUniqueBranches = [...PREDEFINED_BRANCHES].sort(); // Use the hardcoded list

        // Populate employeeCodeToNameMap and employeeCodeToDesignationMap ONLY from Canvassing Data
        employeeCodeToNameMap = {}; // Reset map before populating
        employeeCodeToDesignationMap = {}; // Reset map before populating
        allCanvassingData.forEach(entry => {
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            const employeeName = entry[HEADER_EMPLOYEE_NAME];
            const designation = entry[HEADER_DESIGNATION];

            if (employeeCode) {
                // If an employee code exists in canvassing data, use its name/designation
                employeeCodeToNameMap[employeeCode] = employeeName || employeeCode;
                employeeCodeToDesignationMap[employeeCode] = designation || 'Default';
            }
        });

        // Re-populate allUniqueEmployees based ONLY on canvassing data
        allUniqueEmployees = [...new Set(allCanvassingData.map(entry => entry[HEADER_EMPLOYEE_CODE]))].sort((codeA, codeB) => {
            // Use the name from the map if available, otherwise use the code for sorting and display
            const nameA = employeeCodeToNameMap[codeA] || codeA;
            const nameB = employeeCodeToNameMap[codeB] || codeB;
            return nameA.localeCompare(nameB);
        });

        populateDropdown(branchSelect, allUniqueBranches); // Populate branch dropdown with predefined branches
        populateDropdown(customerViewBranchSelect, allUniqueBranches); // Populate branch dropdown for detailed customer view
        // NEW: Populate dashboard branch select
        populateDropdown(dashboardBranchSelect, allUniqueBranches);
        console.log('Final All Unique Branches (Predefined):', allUniqueBranches);
        console.log('Final Employee Code To Name Map (from Canvassing Data):', employeeCodeToNameMap);
        console.log('Final Employee Code To Designation Map (from Canvassing Data):', employeeCodeToDesignationMap);
        console.log('Final All Unique Employees (Codes from Canvassing Data):', allUniqueEmployees);

        // After data is loaded and maps are populated, render the initial report
        renderAllBranchSnapshot(); // Render the default "All Branch Snapshot" report
    }

    // Populate dropdown utility
    function populateDropdown(selectElement, items, useCodeForValue = false) {
        selectElement.innerHTML = '<option value="">-- Select --</option>'; // Default option
        items.forEach(item => {
            const option = document.createElement('option');
            if (useCodeForValue) {
                // Display name from map or code itself
                option.value = item; // item is employeeCode
                option.textContent = employeeCodeToNameMap[item] || item;
            } else {
                option.value = item; // item is branch name
                option.textContent = item;
            }
            selectElement.appendChild(option);
        });
    }

    // Filter employees based on selected branch
    branchSelect.addEventListener('change', () => {
        const selectedBranch = branchSelect.value;
        if (selectedBranch) {
            employeeFilterPanel.style.display = 'block';

            // Get employee codes ONLY from Canvassing Data for the selected branch
            const employeeCodesInBranchFromCanvassing = allCanvassingData
                .filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch)
                .map(entry => entry[HEADER_EMPLOYEE_CODE]);

            // Combine and unique all employee codes for the selected branch
            const combinedEmployeeCodes = new Set([
                ...employeeCodesInBranchFromCanvassing
            ]);

            // Convert Set back to array and sort
            const sortedEmployeeCodesInBranch = [...combinedEmployeeCodes].sort((codeA, codeB) => {
                // Use the name from the map if available, otherwise use the code for sorting and display
                const nameA = employeeCodeToNameMap[codeA] || codeA;
                const nameB = employeeCodeToNameMap[codeB] || codeB;
                return nameA.localeCompare(nameB);
            });

            populateDropdown(employeeSelect, sortedEmployeeCodesInBranch, true);
            viewOptions.style.display = 'flex'; // Show view options
            // Reset employee selection and employee-specific display when branch changes
            employeeSelect.value = "";
            selectedEmployeeCodeEntries = []; // Clear previous activity filter
            reportDisplay.innerHTML = '<p>Select an employee or choose a report option.</p>';

            // Deactivate all buttons in viewOptions and then reactivate the appropriate ones
            document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));

        } else {
            employeeFilterPanel.style.display = 'none';
            viewOptions.style.display = 'none'; // Hide view options
            reportDisplay.innerHTML = '<p>Please select a branch from the dropdown above to view reports.</p>';
            selectedBranchEntries = []; // Clear previous activity filter
            selectedEmployeeCodeEntries = []; // Clear previous activity filter
        }
    });

    // Handle employee selection (now based on employee CODE)
    employeeSelect.addEventListener('change', () => {
        const selectedEmployeeCode = employeeSelect.value;
        if (selectedEmployeeCode) {
            // Filter activity data by employee code (from allCanvassingData)
            selectedEmployeeCodeEntries = allCanvassingData.filter(entry =>
                entry[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode &&
                entry[HEADER_BRANCH_NAME] === branchSelect.value // Filter by selected branch as well
            );
            const employeeDisplayName = employeeCodeToNameMap[selectedEmployeeCode] || selectedEmployeeCode;
            reportDisplay.innerHTML = `<p>Ready to view reports for ${employeeDisplayName}.</p>`;
            
            // Automatically trigger the Employee Summary (d4.PNG style)
            document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
            viewEmployeeSummaryBtn.classList.add('active'); // Set Employee Summary as active
            renderEmployeeSummary(selectedEmployeeCodeEntries); // Render the Employee Summary
            
        } else {
            selectedEmployeeCodeEntries = []; // Clear previous activity filter
            reportDisplay.innerHTML = '<p>Select an employee or choose a report option.S</p>';
            // Clear active button if employee selection is cleared
            document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        }
    });

    // Helper to calculate total activity from a set of activity entries based on Activity Type
    function calculateTotalActivity(entries) {
        const totalActivity = { 'Visit': 0, 'Call': 0, 'Reference': 0, 'New Customer Leads': 0 }; // Initialize counters
        const productInterests = new Set(); // To collect unique product interests
        
        console.log('Calculating total activity for entries:', entries.length); // Log entries being processed
        entries.forEach((entry, index) => {
            let activityType = entry[HEADER_ACTIVITY_TYPE];
            let typeOfCustomer = entry[HEADER_TYPE_OF_CUSTOMER];
            let productInterested = entry[HEADER_PRODUCT_INTERESTED]; // Get product interested

            // Trim and convert to lowercase for robust comparison
            const trimmedActivityType = activityType ? activityType.trim().toLowerCase() : '';
            const trimmedTypeOfCustomer = typeOfCustomer ? typeOfCustomer.trim().toLowerCase() : '';
            const trimmedProductInterested = productInterested ? productInterested.trim() : ''; // Don't lowercase products unless explicitly asked

            console.log(`--- Entry ${index + 1} Debug ---`);
            console.log(`  Processed Activity Type (trimmed, lowercase): '${trimmedActivityType}'`);
            console.log(`  Processed Type of Customer (trimmed, lowercase): '${trimmedTypeOfCustomer}'`);
            console.log(`  Processed Product Interested (trimmed): '${trimmedProductInterested}'`);


            // Direct matching to user's provided sheet values (now lowercase)
            if (trimmedActivityType === 'visit') {
                totalActivity['Visit']++;
            } else if (trimmedActivityType === 'calls') { // Matches "Calls" from sheet, now lowercase
                totalActivity['Call']++;
            } else if (trimmedActivityType === 'referance') { // Matches "Referance" (with typo) from sheet, now lowercase
                totalActivity['Reference']++;
            } else {
                // If it's not one of the direct activity types, log for debugging
                console.warn(`  Unknown or unhandled Activity Type encountered (trimmed, lowercase): '${trimmedActivityType}'.`);
            }
            
            // --- UPDATED LOGIC FOR 'New Customer Leads' ---
            // Based on the user's previously working script, New Customer Leads are counted
            // if the 'Type of Customer' (now correctly spelled) is simply 'new', regardless of 'Activity Type'.
            if (trimmedTypeOfCustomer === 'new') {
                totalActivity['New Customer Leads']++;
                console.log(`  New Customer Lead INCREMENTED based on Type of Customer === 'new'.`);
            } else {
                console.log(`  New Customer Lead NOT INCREMENTED: Type of Customer is not 'new'.`);
            }
            // --- END UPDATED LOGIC ---

            // Collect unique product interests
            if (trimmedProductInterested) {
                productInterests.add(trimmedProductInterested);
            }
            console.log(`--- End Entry ${index + 1} Debug ---`);
        });
        console.log('Calculated Total Activity Final:', totalActivity);
        
        // Return both total activities and product interests
        return { totalActivity, productInterests: [...productInterests] };
    }

    // Render All Branch Snapshot (now uses PREDEFINED_BRANCHES and checks for participation)
    function renderAllBranchSnapshot() {
        reportDisplay.innerHTML = '<h2>All Branch Snapshot</h2>';
        
        const table = document.createElement('table');
        table.className = 'all-branch-snapshot-table';
        
        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        const headers = ['Branch Name', 'Employees with Activity', 'Total Visits', 'Total Calls', 'Total References', 'Total New Customer Leads'];
        headers.forEach(text => {
            const th = document.createElement('th');
            th.textContent = text;
            headerRow.appendChild(th);
        });

        const tbody = table.createTBody();

        PREDEFINED_BRANCHES.forEach(branch => {
            const branchActivityEntries = allCanvassingData.filter(entry => entry[HEADER_BRANCH_NAME] === branch);
            const { totalActivity } = calculateTotalActivity(branchActivityEntries); // Destructure to get totalActivity
            const employeeCodesInBranch = [...new Set(branchActivityEntries.map(entry => entry[HEADER_EMPLOYEE_CODE]))];
            const displayEmployeeCount = employeeCodesInBranch.length;

            const row = tbody.insertRow();
            // Assign data-label for mobile view
            row.insertCell().setAttribute('data-label', 'Branch Name');
            row.lastChild.textContent = branch;

            row.insertCell().setAttribute('data-label', 'Employees with Activity');
            row.lastChild.textContent = displayEmployeeCount;

            row.insertCell().setAttribute('data-label', 'Total Visits');
            row.lastChild.textContent = totalActivity['Visit'];

            row.insertCell().setAttribute('data-label', 'Total Calls');
            row.lastChild.textContent = totalActivity['Call'];

            row.insertCell().setAttribute('data-label', 'Total References');
            row.lastChild.textContent = totalActivity['Reference'];

            row.insertCell().setAttribute('data-label', 'Total New Customer Leads');
            row.lastChild.textContent = totalActivity['New Customer Leads'];
        });

        reportDisplay.appendChild(table);
    }

   
    // NEW: Render Non-Participating Branches Report (now Zero Visit Branches)
    function renderNonParticipatingBranches() {
        reportDisplay.innerHTML = '<h2>Zero Visit Branches</h2>'; // Changed title
        const zeroVisitBranches = [];

        PREDEFINED_BRANCHES.forEach(branch => {
            const branchActivityEntries = allCanvassingData.filter(entry => entry[HEADER_BRANCH_NAME] === branch);
            const { totalActivity } = calculateTotalActivity(branchActivityEntries); // Get total activities

            // Check if total visits for this branch is 0
            if (totalActivity['Visit'] === 0) {
                zeroVisitBranches.push(branch);
            }
        });

        if (zeroVisitBranches.length > 0) {
            const ul = document.createElement('ul');
            ul.className = 'non-participating-branch-list'; // Reusing existing class
            zeroVisitBranches.forEach(branch => {
                const li = document.createElement('li');
                li.textContent = branch;
                ul.appendChild(li);
            });
            reportDisplay.appendChild(ul);
        } else {
            reportDisplay.innerHTML += '<p class="no-participation-message">All predefined branches have recorded visits!</p>'; // Changed message
        }
    }

    // Render All Staff Overall Performance Report (for d1.PNG)
    function renderOverallStaffPerformanceReport() {
        reportDisplay.innerHTML = '<h2>Overall Staff Performance Report (This Month)</h2>';
        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container'; // For horizontal scrolling
        const table = document.createElement('table');
        table.className = 'performance-table';
        const thead = table.createTHead();
        let headerRow = thead.insertRow();
        // Main Headers
        headerRow.insertCell().textContent = 'Employee Name';
        headerRow.insertCell().textContent = 'Branch Name';
        headerRow.insertCell().textContent = 'Designation';

        // Define metrics for the performance table
        const metrics = ['Visit', 'Call', 'Reference', 'New Customer Leads'];
        metrics.forEach(metric => {
            const th = document.createElement('th');
            th.colSpan = 3; // 'Actual', 'Target', '%'
            th.textContent = metric;
            headerRow.appendChild(th);
        });

        // Sub-headers
        headerRow = thead.insertRow(); // New row for sub-headers
        headerRow.insertCell(); // Empty cell for Employee Name
        headerRow.insertCell(); // Empty cell for Branch Name
        headerRow.insertCell(); // Empty cell for Designation

        metrics.forEach(() => {
            ['Act', 'Tgt', '%'].forEach(subHeader => {
                const th = document.createElement('th');
                th.textContent = subHeader;
                headerRow.appendChild(th);
            });
        });

        const tbody = table.createTBody();

        const currentMonth = new Date().getMonth();
        const currentYear = new Date().getFullYear();

        // Get unique employees who have made at least one entry this month
        const employeesWithActivityThisMonth = [...new Set(allCanvassingData
            .filter(entry => {
                const entryDate = new Date(entry[HEADER_TIMESTAMP]);
                return entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear;
            })
            .map(entry => entry[HEADER_EMPLOYEE_CODE]))].sort((codeA, codeB) => {
                const nameA = employeeCodeToNameMap[codeA] || codeA;
                const nameB = employeeCodeToNameMap[codeB] || codeB;
                return nameA.localeCompare(nameB);
            });

        if (employeesWithActivityThisMonth.length === 0) {
            reportDisplay.innerHTML += '<p>No employee activity found for the current month.</p>';
            return;
        }

        employeesWithActivityThisMonth.forEach(employeeCode => {
            const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
            const branchName = allCanvassingData.find(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode)?.[HEADER_BRANCH_NAME] || 'N/A';
            const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';

            const employeeActivities = allCanvassingData.filter(entry =>
                entry[HEADER_EMPLOYEE_CODE] === employeeCode &&
                new Date(entry[HEADER_TIMESTAMP]).getMonth() === currentMonth &&
                new Date(entry[HEADER_TIMESTAMP]).getFullYear() === currentYear
            );

            const { totalActivity } = calculateTotalActivity(employeeActivities);
            const targets = TARGETS[designation] || TARGETS['Default'];

            const performance = calculatePerformance(totalActivity, targets);

            const row = tbody.insertRow();
            row.insertCell().textContent = employeeName;
            row.insertCell().textContent = branchName;
            row.insertCell().textContent = designation;

            metrics.forEach(metric => {
                const actualValue = totalActivity[metric] || 0;
                const targetValue = targets[metric] || 0; // Ensure target is 0 if undefined

                let percentValue = performance[metric];
                let displayPercent;
                let progressBarClass;
                let progressWidth;

                if (isNaN(percentValue) || targetValue === 0) { // If target is 0, it's N/A
                    displayPercent = 'N/A';
                    progressWidth = 0;
                    progressBarClass = 'no-activity';
                } else {
                    displayPercent = `${Math.round(percentValue)}%`;
                    progressWidth = Math.min(100, Math.round(percentValue));
                    progressBarClass = getProgressBarClass(percentValue);
                }

                // Actual cell
                const actualCell = row.insertCell();
                actualCell.textContent = actualValue;

                // Target cell
                const targetCell = row.insertCell();
                targetCell.textContent = targetValue === 0 ? 'N/A' : targetValue; // Display N/A for 0 targets

                // Percentage cell with progress bar
                const percentCell = row.insertCell();
                percentCell.className = 'progress-cell'; // Add a class for styling
                percentCell.innerHTML = `
                    <div class="progress-bar-container">
                        <div class="progress-bar ${progressBarClass}" style="width: ${progressWidth}%;"></div>
                        <span class="progress-text">${displayPercent}</span>
                    </div>
                `;
            });
        });
        table.appendChild(tbody);
        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }

    // Helper to calculate performance percentage
    function calculatePerformance(actuals, targets) {
        const performance = {};
        for (const metric in targets) {
            const actual = actuals[metric] || 0;
            const target = targets[metric] || 0;
            performance[metric] = target > 0 ? (actual / target) * 100 : 0; // Avoid division by zero
        }
        return performance;
    }

    // Helper to determine progress bar color
    function getProgressBarClass(percentage) {
        if (percentage >= 100) {
            return 'progress-green';
        } else if (percentage >= 75) {
            return 'progress-yellow';
        } else {
            return 'progress-red';
        }
    }

    // Render Branch Performance Report (for d2.PNG and d3.PNG)
    function renderBranchPerformanceReport(branchName) {
        if (!branchName) {
            reportDisplay.innerHTML = '<p>Please select a branch from the dropdown to view its performance report.</p>';
            return;
        }

        reportDisplay.innerHTML = `<h2>${branchName} - Performance Report</h2>`;

        const branchActivities = allCanvassingData.filter(entry => entry[HEADER_BRANCH_NAME] === branchName);

        if (branchActivities.length === 0) {
            reportDisplay.innerHTML += `<p>No activity found for ${branchName}.</p>`;
            return;
        }

        const currentMonth = new Date().getMonth();
        const currentYear = new Date().getFullYear();

        const employeesInBranchThisMonth = [...new Set(branchActivities
            .filter(entry => {
                const entryDate = new Date(entry[HEADER_TIMESTAMP]);
                return entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear;
            })
            .map(entry => entry[HEADER_EMPLOYEE_CODE]))].sort((codeA, codeB) => {
                const nameA = employeeCodeToNameMap[codeA] || codeA;
                const nameB = employeeCodeToNameMap[codeB] || codeB;
                return nameA.localeCompare(nameB);
            });

        if (employeesInBranchThisMonth.length === 0) {
            reportDisplay.innerHTML += `<p>No employee activity found for ${branchName} this month.</p>`;
            return;
        }

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';
        const table = document.createElement('table');
        table.className = 'performance-table';

        const thead = table.createTHead();
        let headerRow = thead.insertRow();
        headerRow.insertCell().textContent = 'Employee Name';
        headerRow.insertCell().textContent = 'Designation';

        const metrics = ['Visit', 'Call', 'Reference', 'New Customer Leads'];
        metrics.forEach(metric => {
            const th = document.createElement('th');
            th.colSpan = 3; // 'Actual', 'Target', '%'
            th.textContent = metric;
            headerRow.appendChild(th);
        });

        // Sub-headers
        headerRow = thead.insertRow(); // New row for sub-headers
        headerRow.insertCell(); // Empty cell for Employee Name
        headerRow.insertCell(); // Empty cell for Designation

        metrics.forEach(() => {
            ['Act', 'Tgt', '%'].forEach(subHeader => {
                const th = document.createElement('th');
                th.textContent = subHeader;
                headerRow.appendChild(th);
            });
        });

        const tbody = table.createTBody();

        employeesInBranchThisMonth.forEach(employeeCode => {
            const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
            const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';

            const employeeActivities = branchActivities.filter(entry =>
                entry[HEADER_EMPLOYEE_CODE] === employeeCode &&
                new Date(entry[HEADER_TIMESTAMP]).getMonth() === currentMonth &&
                new Date(entry[HEADER_TIMESTAMP]).getFullYear() === currentYear
            );

            const { totalActivity } = calculateTotalActivity(employeeActivities);
            const targets = TARGETS[designation] || TARGETS['Default'];
            const performance = calculatePerformance(totalActivity, targets);

            const row = tbody.insertRow();
            row.insertCell().textContent = employeeName;
            row.insertCell().textContent = designation;

            metrics.forEach(metric => {
                const actualValue = totalActivity[metric] || 0;
                const targetValue = targets[metric] || 0;

                let percentValue = performance[metric];
                let displayPercent;
                let progressBarClass;
                let progressWidth;

                if (isNaN(percentValue) || targetValue === 0) {
                    displayPercent = 'N/A';
                    progressWidth = 0;
                    progressBarClass = 'no-activity';
                } else {
                    displayPercent = `${Math.round(percentValue)}%`;
                    progressWidth = Math.min(100, Math.round(percentValue));
                    progressBarClass = getProgressBarClass(percentValue);
                }

                // Actual cell
                const actualCell = row.insertCell();
                actualCell.textContent = actualValue;

                // Target cell
                const targetCell = row.insertCell();
                targetCell.textContent = targetValue === 0 ? 'N/A' : targetValue;

                // Percentage cell with progress bar
                const percentCell = row.insertCell();
                percentCell.className = 'progress-cell';
                percentCell.innerHTML = `
                    <div class="progress-bar-container">
                        <div class="progress-bar ${progressBarClass}" style="width: ${progressWidth}%;"></div>
                        <span class="progress-text">${displayPercent}</span>
                    </div>
                `;
            });
        });

        table.appendChild(tbody);
        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }

    // Render Employee Summary (d4.PNG)
    function renderEmployeeSummary(employeeActivities) {
        if (employeeActivities.length === 0) {
            reportDisplay.innerHTML = '<p>No activities found for the selected employee in this branch.</p>';
            return;
        }

        const employeeCode = employeeActivities[0][HEADER_EMPLOYEE_CODE];
        const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
        const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';
        const branchName = employeeActivities[0][HEADER_BRANCH_NAME] || 'N/A'; // Assuming all activities are for the same branch once filtered

        reportDisplay.innerHTML = `
            <h2>${employeeName} (${designation}) - Summary for ${branchName}</h2>
            <div class="summary-cards">
                <div class="summary-card" id="customerCard1"></div>
                <div class="summary-card" id="customerCard2"></div>
                <div class="summary-card" id="customerCard3"></div>
            </div>
            <div class="data-table-container">
                <table class="activity-table">
                    <thead>
                        <tr>
                            <th>Date</th>
                            <th>Activity Type</th>
                            <th>Type of Customer</th>
                            <th>Prospect Name</th>
                            <th>Remarks</th>
                            <th>Next Follow-up Date</th>
                            <th>Product Interested</th>
                        </tr>
                    </thead>
                    <tbody id="employeeActivitiesTableBody">
                    </tbody>
                </table>
            </div>
        `;

        const { totalActivity, productInterests } = calculateTotalActivity(employeeActivities);
        const targets = TARGETS[designation] || TARGETS['Default'];
        const performance = calculatePerformance(totalActivity, targets);

        const summaryCard1 = document.getElementById('customerCard1');
        const summaryCard2 = document.getElementById('customerCard2');
        const summaryCard3 = document.getElementById('customerCard3');

        summaryCard1.innerHTML = `
            <h3>Overall Activity</h3>
            <p><strong>Visits:</strong> ${totalActivity['Visit']} / ${targets['Visit'] || 'N/A'} (${Math.round(performance['Visit'] || 0)}%)</p>
            <p><strong>Calls:</strong> ${totalActivity['Call']} / ${targets['Call'] || 'N/A'} (${Math.round(performance['Call'] || 0)}%)</p>
            <p><strong>References:</strong> ${totalActivity['Reference']} / ${targets['Reference'] || 'N/A'} (${Math.round(performance['Reference'] || 0)}%)</p>
            <p><strong>New Leads:</strong> ${totalActivity['New Customer Leads']} / ${targets['New Customer Leads'] || 'N/A'} (${Math.round(performance['New Customer Leads'] || 0)}%)</p>
        `;

        summaryCard2.innerHTML = `
            <h3>Recent Activities</h3>
            <ul class="activity-list">
                ${employeeActivities.slice(0, 5).map(entry => `
                    <li>${formatDate(entry[HEADER_TIMESTAMP])}: ${entry[HEADER_ACTIVITY_TYPE]} - ${entry[HEADER_PROSPECT_NAME]}</li>
                `).join('')}
            </ul>
        `;

        summaryCard3.innerHTML = `
            <h3>Product Interests</h3>
            <ul class="product-interest-list">
                ${productInterests.length > 0 ? productInterests.map(interest => `<li>${interest}</li>`).join('') : '<li>No specific interests recorded.</li>'}
            </ul>
        `;

        const employeeActivitiesTableBody = document.getElementById('employeeActivitiesTableBody');
        employeeActivitiesTableBody.innerHTML = ''; // Clear previous entries

        employeeActivities.forEach(entry => {
            const row = employeeActivitiesTableBody.insertRow();
            row.insertCell().textContent = formatDate(entry[HEADER_DATE] || entry[HEADER_TIMESTAMP]);
            row.insertCell().textContent = entry[HEADER_ACTIVITY_TYPE] || '';
            row.insertCell().textContent = entry[HEADER_TYPE_OF_CUSTOMER] || '';
            row.insertCell().textContent = entry[HEADER_PROSPECT_NAME] || '';
            row.insertCell().textContent = entry[HEADER_REMARKS] || '';
            row.insertCell().textContent = formatDate(entry[HEADER_NEXT_FOLLOW_UP_DATE]) || '';
            row.insertCell().textContent = entry[HEADER_PRODUCT_INTERESTED] || '';
        });
    }

    // Render All Entries (d5.PNG)
    function renderAllEntries(entries) {
        if (entries.length === 0) {
            reportDisplay.innerHTML = '<p>No entries found for the selected filters.</p>';
            return;
        }

        reportDisplay.innerHTML = `
            <h2>All Entries</h2>
            <div class="data-table-container">
                <table class="activity-table">
                    <thead>
                        <tr>
                            <th>Timestamp</th>
                            <th>Date</th>
                            <th>Branch Name</th>
                            <th>Employee Name</th>
                            <th>Employee Code</th>
                            <th>Designation</th>
                            <th>Activity Type</th>
                            <th>Type of Customer</th>
                            <th>Lead Source</th>
                            <th>How Contacted</th>
                            <th>Prospect Name</th>
                            <th>Phone Number (Whatsapp)</th>
                            <th>Address</th>
                            <th>Profession</th>
                            <th>DOB/WD</th>
                            <th>Product Interested</th>
                            <th>Remarks</th>
                            <th>Next Follow-up Date</th>
                            <th>Relation With Staff</th>
                        </tr>
                    </thead>
                    <tbody id="allEntriesTableBody">
                    </tbody>
                </table>
            </div>
        `;

        const allEntriesTableBody = document.getElementById('allEntriesTableBody');
        allEntriesTableBody.innerHTML = ''; // Clear previous entries

        entries.forEach(entry => {
            const row = allEntriesTableBody.insertRow();
            row.insertCell().textContent = entry[HEADER_TIMESTAMP] || '';
            row.insertCell().textContent = formatDate(entry[HEADER_DATE]) || '';
            row.insertCell().textContent = entry[HEADER_BRANCH_NAME] || '';
            row.insertCell().textContent = employeeCodeToNameMap[entry[HEADER_EMPLOYEE_CODE]] || entry[HEADER_EMPLOYEE_NAME] || ''; // Use mapped name
            row.insertCell().textContent = entry[HEADER_EMPLOYEE_CODE] || '';
            row.insertCell().textContent = employeeCodeToDesignationMap[entry[HEADER_EMPLOYEE_CODE]] || entry[HEADER_DESIGNATION] || ''; // Use mapped designation
            row.insertCell().textContent = entry[HEADER_ACTIVITY_TYPE] || '';
            row.insertCell().textContent = entry[HEADER_TYPE_OF_CUSTOMER] || '';
            row.insertCell().textContent = entry[HEADER_R_LEAD_SOURCE] || '';
            row.insertCell().textContent = entry[HEADER_HOW_CONTACTED] || '';
            row.insertCell().textContent = entry[HEADER_PROSPECT_NAME] || '';
            row.insertCell().textContent = entry[HEADER_PHONE_NUMBER_WHATSAPP] || '';
            row.insertCell().textContent = entry[HEADER_ADDRESS] || '';
            row.insertCell().textContent = entry[HEADER_PROFESSION] || '';
            row.insertCell().textContent = entry[HEADER_DOB_WD] || '';
            row.insertCell().textContent = entry[HEADER_PRODUCT_INTERESTED] || '';
            row.insertCell().textContent = entry[HEADER_REMARKS] || '';
            row.insertCell().textContent = formatDate(entry[HEADER_NEXT_FOLLOW_UP_DATE]) || '';
            row.insertCell().textContent = entry[HEADER_RELATION_WITH_STAFF] || '';
        });
    }

    // Render Performance Report (Assuming this is a general "performance by activity type" for selected scope)
    function renderPerformanceReport(entries) {
        if (entries.length === 0) {
            reportDisplay.innerHTML = '<p>No data to generate performance report for the selected filters.</p>';
            return;
        }

        reportDisplay.innerHTML = '<h2>Performance Report by Activity Type</h2>';

        const { totalActivity } = calculateTotalActivity(entries);

        const table = document.createElement('table');
        table.className = 'summary-table';
        const thead = table.createTHead();
        const tbody = table.createTBody();

        const headerRow = thead.insertRow();
        ['Metric', 'Total Count'].forEach(text => {
            const th = document.createElement('th');
            th.textContent = text;
            headerRow.appendChild(th);
        });

        for (const [metric, count] of Object.entries(totalActivity)) {
            const row = tbody.insertRow();
            row.insertCell().textContent = metric;
            row.insertCell().textContent = count;
        }

        table.appendChild(thead);
        table.appendChild(tbody);
        reportDisplay.appendChild(table);
    }

    // Leaderboard functions (Visit and Call)
    function renderLeaderboard(activityType, title) {
        reportDisplay.innerHTML = `<h2>${title} Leaderboard (This Month)</h2>`;

        const currentMonth = new Date().getMonth();
        const currentYear = new Date().getFullYear();

        const employeeActivityCounts = {};

        allCanvassingData.forEach(entry => {
            const entryDate = new Date(entry[HEADER_TIMESTAMP]);
            if (entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear &&
                entry[HEADER_ACTIVITY_TYPE] && entry[HEADER_ACTIVITY_TYPE].trim().toLowerCase() === activityType.toLowerCase()) {
                const employeeCode = entry[HEADER_EMPLOYEE_CODE];
                if (employeeCode) {
                    employeeActivityCounts[employeeCode] = (employeeActivityCounts[employeeCode] || 0) + 1;
                }
            }
        });

        const sortedEmployees = Object.entries(employeeActivityCounts).sort(([, countA], [, countB]) => countB - countA);

        if (sortedEmployees.length === 0) {
            reportDisplay.innerHTML += `<p>No ${activityType} activities recorded this month.</p>`;
            return;
        }

        const table = document.createElement('table');
        table.className = 'leaderboard-table';
        const thead = table.createTHead();
        const tbody = table.createTBody();

        const headerRow = thead.insertRow();
        ['Rank', 'Employee Name', 'Branch', `${title} Count`].forEach(text => {
            const th = document.createElement('th');
            th.textContent = text;
            headerRow.appendChild(th);
        });

        sortedEmployees.forEach(( [employeeCode, count], index) => {
            const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
            const branchName = allCanvassingData.find(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode)?.[HEADER_BRANCH_NAME] || 'N/A';

            const row = tbody.insertRow();
            row.insertCell().textContent = index + 1;
            row.insertCell().textContent = employeeName;
            row.insertCell().textContent = branchName;
            row.insertCell().textContent = count;
        });

        table.appendChild(thead);
        table.appendChild(tbody);
        reportDisplay.appendChild(table);
    }

    // Render Staff Participation Report
    function renderStaffParticipationReport() {
        reportDisplay.innerHTML = '<h2>Staff Participation Report (This Month)</h2>';

        const currentMonth = new Date().getMonth();
        const currentYear = new Date().getFullYear();

        const participatingEmployeeCodes = new Set();
        allCanvassingData.forEach(entry => {
            const entryDate = new Date(entry[HEADER_TIMESTAMP]);
            if (entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear) {
                participatingEmployeeCodes.add(entry[HEADER_EMPLOYEE_CODE]);
            }
        });

        const table = document.createElement('table');
        table.className = 'participation-table';
        const thead = table.createTHead();
        const tbody = table.createTBody();

        const headerRow = thead.insertRow();
        ['Employee Name', 'Branch', 'Designation', 'Participated This Month'].forEach(text => {
            const th = document.createElement('th');
            th.textContent = text;
            headerRow.appendChild(th);
        });

        // Loop through all unique employees identified from canvassing data
        allUniqueEmployees.forEach(employeeCode => {
            const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
            const designation = employeeCodeToDesignationMap[employeeCode] || 'N/A';
            const branchName = allCanvassingData.find(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode)?.[HEADER_BRANCH_NAME] || 'N/A';
            const participated = participatingEmployeeCodes.has(employeeCode) ? 'Yes' : 'No';

            const row = tbody.insertRow();
            row.insertCell().textContent = employeeName;
            row.insertCell().textContent = branchName;
            row.insertCell().textContent = designation;
            row.insertCell().textContent = participated;
            if (participated === 'No') {
                row.classList.add('non-participant'); // Add a class for styling
            }
        });

        if (allUniqueEmployees.length === 0) {
            reportDisplay.innerHTML += '<p>No employee data available for participation report.</p>';
            return;
        }

        table.appendChild(thead);
        table.appendChild(tbody);
        reportDisplay.appendChild(table);
    }

    // New function to download overall staff performance as CSV
    function downloadOverallStaffPerformanceReportCSV() {
        console.log('Download button clicked - Overall Performance Report CSV');
        const currentMonth = new Date().getMonth();
        const currentYear = new Date().getFullYear();

        const employeesWithActivityThisMonth = [...new Set(allCanvassingData
            .filter(entry => {
                const entryDate = new Date(entry[HEADER_TIMESTAMP]);
                return entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear;
            })
            .map(entry => entry[HEADER_EMPLOYEE_CODE]))];

        if (employeesWithActivityThisMonth.length === 0) {
            alert("No employee activity found for the current month to download.");
            return;
        }

        let csvContent = "Employee Name,Branch Name,Designation,Visit Act,Visit Tgt,Visit %,Call Act,Call Tgt,Call %,Reference Act,Reference Tgt,Reference %,New Customer Leads Act,New Customer Leads Tgt,New Customer Leads %\n";

        const metrics = ['Visit', 'Call', 'Reference', 'New Customer Leads'];

        employeesWithActivityThisMonth.forEach(employeeCode => {
            const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
            const branchName = allCanvassingData.find(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode)?.[HEADER_BRANCH_NAME] || 'N/A';
            const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';

            const employeeActivities = allCanvassingData.filter(entry =>
                entry[HEADER_EMPLOYEE_CODE] === employeeCode &&
                new Date(entry[HEADER_TIMESTAMP]).getMonth() === currentMonth &&
                new Date(entry[HEADER_TIMESTAMP]).getFullYear() === currentYear
            );

            const { totalActivity } = calculateTotalActivity(employeeActivities);
            const targets = TARGETS[designation] || TARGETS['Default'];
            const performance = calculatePerformance(totalActivity, targets);

            let rowData = `"${employeeName}","${branchName}","${designation}"`;

            metrics.forEach(metric => {
                const actualValue = totalActivity[metric] || 0;
                const targetValue = targets[metric] || 0;
                let percentValue = performance[metric];

                let displayPercent;
                if (isNaN(percentValue) || targetValue === 0) {
                    displayPercent = 'N/A';
                } else {
                    displayPercent = `${Math.round(percentValue)}%`;
                }
                rowData += `,"${actualValue}","${targetValue === 0 ? 'N/A' : targetValue}","${displayPercent}"`;
            });
            csvContent += rowData + "\n";
        });

        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        if (link.download !== undefined) { // feature detection
            const url = URL.createObjectURL(blob);
            link.setAttribute('href', url);
            link.setAttribute('download', 'Overall_Staff_Performance_Report.csv');
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    }


    // Function to download detailed customer report as CSV
    function downloadDetailedCustomerReportCSV(entries) {
        if (!entries || entries.length === 0) {
            alert("No customer data available to download.");
            return;
        }

        // Define CSV headers in the desired order
        const csvHeaders = [
            HEADER_TIMESTAMP,
            HEADER_DATE,
            HEADER_BRANCH_NAME,
            HEADER_EMPLOYEE_NAME, // Will use mapped name
            HEADER_EMPLOYEE_CODE,
            HEADER_DESIGNATION,   // Will use mapped designation
            HEADER_ACTIVITY_TYPE,
            HEADER_TYPE_OF_CUSTOMER,
            HEADER_R_LEAD_SOURCE,
            HEADER_HOW_CONTACTED,
            HEADER_PROSPECT_NAME,
            HEADER_PHONE_NUMBER_WHATSAPP,
            HEADER_ADDRESS,
            HEADER_PROFESSION,
            HEADER_DOB_WD,
            HEADER_PRODUCT_INTERESTED,
            HEADER_REMARKS,
            HEADER_NEXT_FOLLOW_UP_DATE,
            HEADER_RELATION_WITH_STAFF,
            HEADER_FAMILY_DETAILS_1,
            HEADER_FAMILY_DETAILS_2,
            HEADER_FAMILY_DETAILS_3,
            HEADER_FAMILY_DETAILS_4,
            HEADER_PROFILE_OF_CUSTOMER
        ];

        let csvContent = csvHeaders.map(header => `"${header}"`).join(',') + '\n';

        entries.forEach(entry => {
            const rowValues = csvHeaders.map(header => {
                let value = entry[header] || ''; // Get value, default to empty string

                // Handle specific headers for mapped names/designations if different from raw
                if (header === HEADER_EMPLOYEE_NAME) {
                    value = employeeCodeToNameMap[entry[HEADER_EMPLOYEE_CODE]] || entry[HEADER_EMPLOYEE_NAME] || '';
                } else if (header === HEADER_DESIGNATION) {
                    value = employeeCodeToDesignationMap[entry[HEADER_EMPLOYEE_CODE]] || entry[HEADER_DESIGNATION] || '';
                } else if (header === HEADER_DATE || header === HEADER_NEXT_FOLLOW_UP_DATE) {
                     value = formatDate(value); // Format dates
                }

                // Enclose in double quotes and escape existing double quotes
                return `"${String(value).replace(/"/g, '""')}"`;
            });
            csvContent += rowValues.join(',') + '\n';
        });

        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        if (link.download !== undefined) {
            const url = URL.createObjectURL(blob);
            link.setAttribute('href', url);
            link.setAttribute('download', 'Detailed_Customer_Report.csv');
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    }


    // --- Google Apps Script (GAS) Web App Integration ---
    async function sendDataToGoogleAppsScript(action, data) {
        displayEmployeeManagementMessage(`Sending ${action} request...`, false);
        try {
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                mode: 'cors', // Crucial for cross-origin requests
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ action, data }),
            });

            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`Server responded with an error: ${response.status} - ${errorText}`);
            }

            const result = await response.json();
            if (result.status === 'SUCCESS') {
                displayEmployeeManagementMessage(`${action} successful: ${result.message}`, false);
                // After successful operation, re-process data to update front-end
                await processData();
                return true;
            } else {
                displayEmployeeManagementMessage(`${action} failed: ${result.message}`, true);
                return false;
            }
        } catch (error) {
            console.error(`Error during ${action} operation:`, error);
            displayEmployeeManagementMessage(`Error: ${error.message}`, true);
            return false;
        }
    }

    // --- Employee Management Event Listeners ---
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const newEmployeeName = newEmployeeNameInput.value.trim();
            const newEmployeeCode = newEmployeeCodeInput.value.trim();
            const newBranchName = newBranchNameInput.value.trim();
            const newDesignation = newDesignationInput.value.trim();

            if (!newEmployeeName || !newEmployeeCode || !newBranchName || !newDesignation) {
                displayEmployeeManagementMessage('All fields are required for adding an employee.', true);
                return;
            }

            const addData = {
                [MASTER_HEADER_EMPLOYEE_NAME]: newEmployeeName,
                [MASTER_HEADER_EMPLOYEE_CODE]: newEmployeeCode,
                [MASTER_HEADER_BRANCH]: newBranchName,
                [MASTER_HEADER_DESIGNATION]: newDesignation
            };
            const success = await sendDataToGoogleAppsScript('add_employee', addData);

            if (success) {
                addEmployeeForm.reset();
            }
        });
    }

    if (bulkAddEmployeeForm) {
        bulkAddEmployeeForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const bulkBranchName = bulkEmployeeBranchNameInput.value.trim();
            const bulkDetails = bulkEmployeeDetailsTextarea.value.trim();

            if (!bulkBranchName || !bulkDetails) {
                displayEmployeeManagementMessage('Branch Name and Employee Details are required for bulk entry.', true);
                return;
            }

            const employees = bulkDetails.split('\n').map(line => {
                const parts = line.split(',').map(p => p.trim());
                if (parts.length === 3) {
                    return {
                        [MASTER_HEADER_EMPLOYEE_NAME]: parts[0],
                        [MASTER_HEADER_EMPLOYEE_CODE]: parts[1],
                        [MASTER_HEADER_DESIGNATION]: parts[2],
                        [MASTER_HEADER_BRANCH]: bulkBranchName
                    };
                }
                return null;
            }).filter(employee => employee !== null);

            if (employees.length === 0) {
                displayEmployeeManagementMessage('No valid employee details found in the bulk entry. Format: Name,Code,Designation', true);
                return;
            }

            const success = await sendDataToGoogleAppsScript('bulk_add_employees', { employees: employees });

            if (success) {
                bulkAddEmployeeForm.reset();
            }
        });
    }

    if (deleteEmployeeForm) {
        deleteEmployeeForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const employeeCodeToDelete = deleteEmployeeCodeInput.value.trim();

            if (!employeeCodeToDelete) {
                displayEmployeeManagementMessage('Employee Code is required for deletion.', true);
                return;
            }

            const deleteData = { [HEADER_EMPLOYEE_CODE]: employeeCodeToDelete };
            const success = await sendDataToGoogleAppsScript('delete_employee', deleteData);

            if (success) {
                deleteEmployeeForm.reset();
            }
        });
    }


    // --- Tab Switching Logic (existing, refined) ---
    const allSections = [
        reportsSection,
        detailedCustomerViewSection,
        employeeManagementSection,
        performanceSummarySection // NEW: Include the new section
    ];

    const allTabButtons = [
        allBranchSnapshotTabBtn,
        allStaffOverallPerformanceTabBtn,
        nonParticipatingBranchesTabBtn,
        branchPerformanceTabBtn,
        detailedCustomerViewTabBtn,
        employeeManagementTabBtn,
        performanceSummaryTabBtn // NEW: Include the new button
    ];

    function showTab(activeButtonId) {
        allSections.forEach(section => {
            if (section) { // Check if section exists
                section.style.display = 'none';
            }
        });

        allTabButtons.forEach(button => {
            if (button) { // Check if button exists
                button.classList.remove('active');
            }
        });

        // Determine which section to show based on the active button
        let sectionToShow;
        if (activeButtonId === 'allBranchSnapshotTabBtn' ||
            activeButtonId === 'allStaffOverallPerformanceTabBtn' ||
            activeButtonId === 'nonParticipatingBranchesTabBtn' ||
            activeButtonId === 'branchPerformanceTabBtn') {
            sectionToShow = reportsSection;
        } else if (activeButtonId === 'detailedCustomerViewTabBtn') {
            sectionToShow = detailedCustomerViewSection;
        } else if (activeButtonId === 'employeeManagementTabBtn') {
            sectionToShow = employeeManagementSection;
        } else if (activeButtonId === 'performanceSummaryTabBtn') { // NEW: Handle new section
            sectionToShow = performanceSummarySection;
        }

        if (sectionToShow) { // Show the determined section
            sectionToShow.style.display = 'block';
        }

        const activeButton = document.getElementById(activeButtonId);
        if (activeButton) { // Add active class to the clicked button
            activeButton.classList.add('active');
        }
    }


    // --- Event Listeners for Tab Buttons ---
    if (allBranchSnapshotTabBtn) {
        allBranchSnapshotTabBtn.addEventListener('click', () => {
            showTab('allBranchSnapshotTabBtn');
            renderAllBranchSnapshot();
        });
    }

    if (nonParticipatingBranchesTabBtn) {
        nonParticipatingBranchesTabBtn.addEventListener('click', () => {
            showTab('nonParticipatingBranchesTabBtn');
            renderNonParticipatingBranches();
        });
    }

    // NEW: Event Listener for "Branch Performance Reports" tab button
    if (branchPerformanceTabBtn) {
        branchPerformanceTabBtn.addEventListener('click', () => {
            showTab('branchPerformanceTabBtn');
            // When this tab is clicked, render the report based on the currently selected branch
            // If no branch is selected, the function will display a prompt.
            renderBranchPerformanceReport(branchSelect.value);
        });
    }

    // --- NEW: Event Listener for "All Staff Performance (Overall)" tab button ---
    if (allStaffOverallPerformanceTabBtn) {
        allStaffOverallPerformanceTabBtn.addEventListener('click', () => {
            showTab('allStaffOverallPerformanceTabBtn');
            renderOverallStaffPerformanceReport();
        });
    }

    // Event listener for detailed customer view tab
    if (detailedCustomerViewTabBtn) {
        detailedCustomerViewTabBtn.addEventListener('click', () => {
            showTab('detailedCustomerViewTabBtn');
            // Initial population for customer view dropdowns
            populateDropdown(customerViewBranchSelect, allUniqueBranches);
            // Clear any previous customer details
            customerCanvassedList.innerHTML = '<p>Select a branch and/or employee to view customer details.</p>';
            customerDetailsContent.style.display = 'none'; // Hide details panel initially
        });
    }

    // Customer View Branch Selection
    if (customerViewBranchSelect) {
        customerViewBranchSelect.addEventListener('change', () => {
            const selectedBranch = customerViewBranchSelect.value;
            customerViewEmployeeSelect.innerHTML = '<option value="">-- Select Employee --</option>'; // Reset employee dropdown

            if (selectedBranch) {
                // Filter employees who have canvassed in this branch based on allCanvassingData
                const employeesInBranch = [...new Set(
                    allCanvassingData
                        .filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch)
                        .map(entry => entry[HEADER_EMPLOYEE_CODE])
                )].sort((codeA, codeB) => {
                    const nameA = employeeCodeToNameMap[codeA] || codeA;
                    const nameB = employeeCodeToNameMap[codeB] || codeB;
                    return nameA.localeCompare(nameB);
                });
                populateDropdown(customerViewEmployeeSelect, employeesInBranch, true);
            }
            renderCustomerCanvassedList(); // Re-render list based on current filters
        });
    }

    // Customer View Employee Selection
    if (customerViewEmployeeSelect) {
        customerViewEmployeeSelect.addEventListener('change', () => {
            renderCustomerCanvassedList(); // Re-render list based on current filters
        });
    }

    // Function to render the list of canvassed customers based on filters
    function renderCustomerCanvassedList() {
        const selectedBranch = customerViewBranchSelect.value;
        const selectedEmployeeCode = customerViewEmployeeSelect.value;

        let filteredCustomers = allCanvassingData;

        if (selectedBranch) {
            filteredCustomers = filteredCustomers.filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch);
        }
        if (selectedEmployeeCode) {
            filteredCustomers = filteredCustomers.filter(entry => entry[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode);
        }

        customerCanvassedList.innerHTML = ''; // Clear previous list
        customerDetailsContent.style.display = 'none'; // Hide details when list updates

        if (filteredCustomers.length === 0) {
            customerCanvassedList.innerHTML = '<p>No customers canvassed for the selected criteria.</p>';
            return;
        }

        const ul = document.createElement('ul');
        ul.className = 'customer-list';
        filteredCustomers.forEach((customer, index) => {
            const li = document.createElement('li');
            li.textContent = `${customer[HEADER_PROSPECT_NAME] || 'Unknown'} (${customer[HEADER_ACTIVITY_TYPE] || 'N/A'}) - ${formatDate(customer[HEADER_DATE] || customer[HEADER_TIMESTAMP])}`;
            li.addEventListener('click', () => displayCustomerDetails(customer));
            ul.appendChild(li);
        });
        customerCanvassedList.appendChild(ul);
    }

    // Function to display detailed customer information
    function displayCustomerDetails(customer) {
        customerDetailsContent.style.display = 'block';

        const familyDetails = [
            customer[HEADER_FAMILY_DETAILS_1],
            customer[HEADER_FAMILY_DETAILS_2],
            customer[HEADER_FAMILY_DETAILS_3],
            customer[HEADER_FAMILY_DETAILS_4]
        ].filter(Boolean).map(detail => `<li>${detail}</li>`).join('') || '<li>No family details provided.</li>';


        customerCard1.innerHTML = `
            <h3>General Details</h3>
            <p><strong>Name:</strong> ${customer[HEADER_PROSPECT_NAME] || 'N/A'}</p>
            <p><strong>Phone:</strong> ${customer[HEADER_PHONE_NUMBER_WHATSAPP] || 'N/A'}</p>
            <p><strong>Address:</strong> ${customer[HEADER_ADDRESS] || 'N/A'}</p>
            <p><strong>Profession:</strong> ${customer[HEADER_PROFESSION] || 'N/A'}</p>
            <p><strong>DOB/WD:</strong> ${customer[HEADER_DOB_WD] || 'N/A'}</p>
            <p><strong>Relation with Staff:</strong> ${customer[HEADER_RELATION_WITH_STAFF] || 'N/A'}</p>
            <p><strong>Profile of Customer:</strong> ${customer[HEADER_PROFILE_OF_CUSTOMER] || 'N/A'}</p>
        `;

        customerCard2.innerHTML = `
            <h3>Activity Details</h3>
            <p><strong>Date:</strong> ${formatDate(customer[HEADER_DATE] || customer[HEADER_TIMESTAMP])}</p>
            <p><strong>Branch:</strong> ${customer[HEADER_BRANCH_NAME] || 'N/A'}</p>
            <p><strong>Employee:</strong> ${employeeCodeToNameMap[customer[HEADER_EMPLOYEE_CODE]] || customer[HEADER_EMPLOYEE_NAME] || 'N/A'}</p>
            <p><strong>Activity Type:</strong> ${customer[HEADER_ACTIVITY_TYPE] || 'N/A'}</p>
            <p><strong>Customer Type:</strong> ${customer[HEADER_TYPE_OF_CUSTOMER] || 'N/A'}</p>
            <p><strong>Lead Source:</strong> ${customer[HEADER_R_LEAD_SOURCE] || 'N/A'}</p>
            <p><strong>How Contacted:</strong> ${customer[HEADER_HOW_CONTACTED] || 'N/A'}</p>
            <p><strong>Product Interested:</strong> ${customer[HEADER_PRODUCT_INTERESTED] || 'N/A'}</p>
            <p><strong>Remarks:</strong> ${customer[HEADER_REMARKS] || 'N/A'}</p>
            <p><strong>Next Follow-up:</strong> ${formatDate(customer[HEADER_NEXT_FOLLOW_UP_DATE]) || 'N/A'}</p>
        `;

        customerCard3.innerHTML = `
            <h3>Family Details</h3>
            <ul class="family-details-list">
                ${familyDetails}
            </ul>
        `;

        // Attach download button functionality for current customer's data
        if (downloadDetailedCustomerReportBtn) {
            downloadDetailedCustomerReportBtn.onclick = () => downloadDetailedCustomerReportCSV([customer]);
        }
    }


    // NEW: Render Performance Summary Dashboard
    function renderPerformanceSummaryDashboard() {
        reportDisplay.innerHTML = '<h2>Performance Summary Dashboard</h2>';

        const currentMonth = new Date().getMonth();
        const currentYear = new Date().getFullYear();

        // Filter data for the current month/year
        const currentMonthData = allCanvassingData.filter(entry => {
            const entryDate = new Date(entry[HEADER_TIMESTAMP]);
            return entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear;
        });

        if (currentMonthData.length === 0) {
            reportDisplay.innerHTML += '<p>No activity recorded for the current month.</p>';
            return;
        }

        const monthlySummary = calculateMonthlySummary(currentMonthData);

        // Create main dashboard layout
        const dashboardGrid = document.createElement('div');
        dashboardGrid.className = 'dashboard-grid';

        // 1. Overall Company Performance
        const overallCard = document.createElement('div');
        overallCard.className = 'summary-card overall-performance';
        overallCard.innerHTML = `
            <h3>Overall Company Performance (This Month)</h3>
            <p><strong>Total Visits:</strong> ${monthlySummary.totalVisits}</p>
            <p><strong>Total Calls:</strong> ${monthlySummary.totalCalls}</p>
            <p><strong>Total References:</strong> ${monthlySummary.totalReferences}</p>
            <p><strong>Total New Customer Leads:</strong> ${monthlySummary.totalNewCustomerLeads}</p>
        `;
        dashboardGrid.appendChild(overallCard);

        // 2. Branch-wise Summary (Top 5 for Visits)
        const branchVisitsCard = document.createElement('div');
        branchVisitsCard.className = 'summary-card branch-visits';
        branchVisitsCard.innerHTML = '<h3>Branch-wise Visits (Top 5)</h3>';
        const branchVisitList = document.createElement('ul');
        Object.entries(monthlySummary.branchVisits)
            .sort(([, a], [, b]) => b - a)
            .slice(0, 5)
            .forEach(([branch, count]) => {
                const li = document.createElement('li');
                li.textContent = `${branch}: ${count} Visits`;
                branchVisitList.appendChild(li);
            });
        branchVisitsCard.appendChild(branchVisitList);
        dashboardGrid.appendChild(branchVisitsCard);

        // 3. Employee-wise Top Performers (Top 5 for New Customer Leads)
        const employeeLeadsCard = document.createElement('div');
        employeeLeadsCard.className = 'summary-card employee-leads';
        employeeLeadsCard.innerHTML = '<h3>Employee Top Performers (New Leads)</h3>';
        const employeeLeadList = document.createElement('ul');
        Object.entries(monthlySummary.employeeNewCustomerLeads)
            .sort(([, a], [, b]) => b - a)
            .slice(0, 5)
            .forEach(([employeeCode, count]) => {
                const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
                const li = document.createElement('li');
                li.textContent = `${employeeName}: ${count} New Leads`;
                employeeLeadList.appendChild(li);
            });
        employeeLeadsCard.appendChild(employeeLeadList);
        dashboardGrid.appendChild(employeeLeadsCard);
        
        // 4. Activity Type Distribution (Pie Chart idea, but as list for now)
        const activityDistributionCard = document.createElement('div');
        activityDistributionCard.className = 'summary-card activity-distribution';
        activityDistributionCard.innerHTML = '<h3>Activity Type Distribution</h3>';
        const activityDistList = document.createElement('ul');
        for (const [type, count] of Object.entries(monthlySummary.activityTypeCounts)) {
            const li = document.createElement('li');
            li.textContent = `${type}: ${count}`;
            activityDistList.appendChild(li);
        }
        activityDistributionCard.appendChild(activityDistList);
        dashboardGrid.appendChild(activityDistributionCard);

        reportDisplay.appendChild(dashboardGrid);

        // Add dropdown for Branch-specific Summary at the bottom
        const branchSummaryContainer = document.createElement('div');
        branchSummaryContainer.className = 'branch-summary-controls';
        branchSummaryContainer.innerHTML = `
            <h3>Branch-specific Summary</h3>
            <label for="dashboardBranchSelect">Select Branch:</label>
            <select id="dashboardBranchSelect"></select>
            <div id="selectedBranchSummary"></div>
        `;
        reportDisplay.appendChild(branchSummaryContainer);

        // Re-populate the dashboardBranchSelect dropdown
        const dashBranchSelect = document.getElementById('dashboardBranchSelect');
        populateDropdown(dashBranchSelect, allUniqueBranches);

        // Event listener for branch selection in dashboard
        dashBranchSelect.addEventListener('change', () => {
            const selectedBranch = dashBranchSelect.value;
            const selectedBranchSummaryDiv = document.getElementById('selectedBranchSummary');
            selectedBranchSummaryDiv.innerHTML = ''; // Clear previous summary

            if (selectedBranch) {
                const branchData = currentMonthData.filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch);
                const { totalActivity } = calculateTotalActivity(branchData);
                
                const summaryDiv = document.createElement('div');
                summaryDiv.className = 'selected-branch-summary-details';
                summaryDiv.innerHTML = `
                    <h4>${selectedBranch} - This Month</h4>
                    <p><strong>Visits:</strong> ${totalActivity['Visit']}</p>
                    <p><strong>Calls:</strong> ${totalActivity['Call']}</p>
                    <p><strong>References:</strong> ${totalActivity['Reference']}</p>
                    <p><strong>New Customer Leads:</strong> ${totalActivity['New Customer Leads']}</p>
                `;
                selectedBranchSummaryDiv.appendChild(summaryDiv);
            }
        });

        // Add download button functionality
        if (downloadPerformanceSummaryReportBtn) {
            downloadPerformanceSummaryReportBtn.onclick = () => downloadPerformanceSummaryReportCSV(currentMonthData);
        }
    }

    // Helper function for Performance Summary Dashboard
    function calculateMonthlySummary(data) {
        const summary = {
            totalVisits: 0,
            totalCalls: 0,
            totalReferences: 0,
            totalNewCustomerLeads: 0,
            branchVisits: {},
            employeeNewCustomerLeads: {},
            activityTypeCounts: {} // To count occurrences of each activity type
        };

        data.forEach(entry => {
            const branchName = entry[HEADER_BRANCH_NAME];
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            const activityType = entry[HEADER_ACTIVITY_TYPE];
            const typeOfCustomer = entry[HEADER_TYPE_OF_CUSTOMER];

            // Overall totals (same logic as calculateTotalActivity)
            const trimmedActivityType = activityType ? activityType.trim().toLowerCase() : '';
            const trimmedTypeOfCustomer = typeOfCustomer ? typeOfCustomer.trim().toLowerCase() : '';

            if (trimmedActivityType === 'visit') {
                summary.totalVisits++;
            } else if (trimmedActivityType === 'calls') {
                summary.totalCalls++;
            } else if (trimmedActivityType === 'referance') { // Typo "referance" is maintained as per sheet
                summary.totalReferences++;
            }

            if (trimmedTypeOfCustomer === 'new') {
                summary.totalNewCustomerLeads++;
            }

            // Branch-wise Visits
            if (branchName && trimmedActivityType === 'visit') {
                summary.branchVisits[branchName] = (summary.branchVisits[branchName] || 0) + 1;
            }

            // Employee-wise New Customer Leads
            if (employeeCode && trimmedTypeOfCustomer === 'new') {
                summary.employeeNewCustomerLeads[employeeCode] = (summary.employeeNewCustomerLeads[employeeCode] || 0) + 1;
            }

            // Activity Type Distribution
            if (activityType) {
                summary.activityTypeCounts[activityType] = (summary.activityTypeCounts[activityType] || 0) + 1;
            }
        });

        return summary;
    }

    // NEW: Function to download Performance Summary Dashboard as CSV
    function downloadPerformanceSummaryReportCSV(data) {
        if (data.length === 0) {
            alert("No activity data for the current month to download for the Performance Summary.");
            return;
        }

        const monthlySummary = calculateMonthlySummary(data);

        let csvContent = "Performance Summary Report (This Month)\n";
        csvContent += "Metric,Count\n";
        csvContent += `"Total Visits",${monthlySummary.totalVisits}\n`;
        csvContent += `"Total Calls",${monthlySummary.totalCalls}\n`;
        csvContent += `"Total References",${monthlySummary.totalReferences}\n`;
        csvContent += `"Total New Customer Leads",${monthlySummary.totalNewCustomerLeads}\n`;
        csvContent += "\nBranch-wise Visits (All Branches)\n";
        csvContent += "Branch,Visits\n";
        Object.entries(monthlySummary.branchVisits)
            .sort(([, a], [, b]) => b - a)
            .forEach(([branch, count]) => {
                csvContent += `"${branch}",${count}\n`;
            });
        csvContent += "\nEmployee Top Performers (New Leads - All Employees)\n";
        csvContent += "Employee Name,New Customer Leads\n";
        Object.entries(monthlySummary.employeeNewCustomerLeads)
            .sort(([, a], [, b]) => b - a)
            .forEach(([employeeCode, count]) => {
                const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
                csvContent += `"${employeeName}",${count}\n`;
            });
        csvContent += "\nActivity Type Distribution\n";
        csvContent += "Activity Type,Count\n";
        for (const [type, count] of Object.entries(monthlySummary.activityTypeCounts)) {
            csvContent += `"${type}",${count}\n`;
        }

        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        if (link.download !== undefined) {
            const url = URL.createObjectURL(blob);
            link.setAttribute('href', url);
            link.setAttribute('download', 'Performance_Summary_Dashboard.csv');
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    }


 // NEW: Event Listener for "Download Overall Staff Performance CSV" button
if (downloadOverallStaffPerformanceReportBtn) { // This variable is correct
    // CORRECTED LINE: Ensure this matches the declaration at the top
    downloadOverallStaffPerformanceReportBtn.addEventListener('click', () => { 
        downloadOverallStaffPerformanceReportCSV();
    });

}
    // NEW: Event Listener for "Performance Summary Dashboard" tab button
    if (performanceSummaryTabBtn) {
        performanceSummaryTabBtn.addEventListener('click', () => {
            showTab('performanceSummaryTabBtn'); // This will show the new section
            renderPerformanceSummaryDashboard(); // This is the new function we'll create right below
        });
    }

    // Initial data fetch and tab display when the page loads
    processData();
    showTab('allBranchSnapshotTabBtn'); // Your current default starting tab
}); // This is the closing brace for DOMContentLoaded
