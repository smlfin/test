document.addEventListener('DOMContentLoaded', () => {
    // *** Configuration ***
    // This URL is for your Canvassing Data sheet. Ensure it's correct and published as CSV.
    // NOTE: If you are still getting 404, this URL is the problem.
    const DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=696550092&single=true&output=csv"; 

    // IMPORTANT: Replace this with YOUR DEPLOYED GOOGLE APPS SCRIPT WEB APP URL
    // NOTE: If you are getting errors sending data, this URL is the problem.
    const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzEYf0CKgwP0O4-z1lup1lDZImD1dQVEveLWsHwa_7T5ltndfIuRWXVZqFDj03_proD/exec"; // <-- PASTE YOUR NEWLY DEPLOYED WEB APP URL HERE

    // We will IGNORE MasterEmployees sheet for data fetching and report generation
    // Employee management functions in Apps Script still use the MASTER_SHEET_ID you've set up in code.gs
    // For front-end reporting, all employee and branch data will come from Canvassing Data and predefined list.
    const EMPLOYEE_MASTER_DATA_URL = "UNUSED"; // Marked as UNUSED for clarity, won't be fetched for reports

    const MONTHLY_WORKING_DAYS = 22; // Common approximation for a month's working days

    // TARGETS will now be fetched dynamically, keep a placeholder for initial structure or default
    let TARGETS = {}; // This will be populated by fetched data

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


    // *** DOM Elements ***
    const branchSelect = document.getElementById('branchSelect');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel');
    const employeeSelect = document.getElementById('employeeSelect');
    const viewOptions = document.getElementById('viewOptions');
    const viewBranchPerformanceReportBtn = document.getElementById('viewBranchPerformanceReportBtn');
    const viewEmployeeSummaryBtn = document.getElementById('viewEmployeeSummaryBtn');
    const viewAllEntriesBtn = document.getElementById('viewAllEntriesBtn');
    const viewPerformanceReportBtn = document.getElementById('viewPerformanceReportBtn');

    // Main Report Display Area
    const reportDisplay = document.getElementById('reportDisplay');
    // Dedicated message area element
    const statusMessageDiv = document.getElementById('statusMessage');


    // Tab buttons for main navigation
    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');
    const nonParticipatingBranchesTabBtn = document.getElementById('nonParticipatingBranchesTabBtn');
    const detailedCustomerViewTabBtn = document.getElementById('detailedCustomerViewTabBtn');
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');
    const targetManagementTabBtn = document.getElementById('targetManagementTabBtn'); // NEW

    // Main Content Sections to toggle
    const reportsSection = document.getElementById('reportsSection');
    const detailedCustomerViewSection = document.getElementById('detailedCustomerViewSection');
    const employeeManagementSection = document.getElementById('employeeManagementSection');
    const targetManagementSection = document.getElementById('targetManagementSection'); // NEW

    // NEW: Detailed Customer View Elements
    const customerViewBranchSelect = document.getElementById('customerViewBranchSelect');
    const customerViewEmployeeSelect = document.getElementById('customerViewEmployeeSelect');
    const customerCanvassedList = document.getElementById('customerCanvassedList');
    const customerDetailsContent = document.getElementById('customerDetailsContent');


    // Employee Management Form Elements
    const addEmployeeForm = document.getElementById('addEmployeeForm');
    const newEmployeeNameInput = document.getElementById('newEmployeeName');
    const newEmployeeCodeInput = document.getElementById('newEmployeeCode');
    const newBranchNameInput = document.getElementById('newBranchName');
    const newDesignationInput = document.getElementById('newDesignation');
    const employeeManagementMessage = document.getElementById('employeeManagementMessage');

    const bulkAddEmployeeForm = document.getElementById('bulkAddEmployeeForm');
    const bulkEmployeeBranchNameInput = document.getElementById('bulkEmployeeBranchName');
    const bulkEmployeeDetailsTextarea = document.getElementById('bulkEmployeeDetails');

    const deleteEmployeeForm = document.getElementById('deleteEmployeeForm');
    const deleteEmployeeCodeInput = document.getElementById('deleteEmployeeCode');

    // NEW: Target Management Elements
    const targetManagementMessage = document.getElementById('targetManagementMessage');
    const setTargetsForm = document.getElementById('setTargetsForm');
    const targetDesignationSelect = document.getElementById('targetDesignationSelect');
    const targetVisitInput = document.getElementById('targetVisit');
    const targetCallInput = document.getElementById('targetCall');
    const targetReferenceInput = document.getElementById('targetReference');
    const targetNewCustomerLeadsInput = document.getElementById('targetNewCustomerLeads');
    const currentTargetsDisplay = document.getElementById('currentTargetsDisplay');


    // Global variables to store fetched data
    let allCanvassingData = []; // Raw activity data from Form Responses 2
    let allUniqueBranches = []; // Will be populated from PREDEFINED_BRANCHES
    let allUniqueEmployees = []; // Employee codes from Canvassing Data
    let employeeCodeToNameMap = {}; // {code: name} from Canvassing Data
    let employeeCodeToDesignationMap = {}; // {code: designation} from Canvassing Data
    let selectedBranchEntries = []; // Activity entries filtered by branch (for main reports section)
    let selectedEmployeeCodeEntries = []; // Activity entries filtered by employee code (for main reports section)
    let currentTargets = {}; // NEW: Stores the fetched targets

    // Utility to format date to ISO-MM-DD
    const formatDate = (dateString) => {
        if (!dateString) return '';
        const date = new Date(dateString);
        if (isNaN(date.getTime())) return dateString;
        return date.toISOString().split('T')[0];
    };

    // Helper to clear and display messages in a specific div (now targets statusMessageDiv)
    function displayMessage(message, type = 'info') { 
        if (statusMessageDiv) {
            statusMessageDiv.innerHTML = `<div class="message ${type}">${message}</div>`;
            statusMessageDiv.style.display = 'block';
            setTimeout(() => {
                statusMessageDiv.innerHTML = ''; // Clear message
                statusMessageDiv.style.display = 'none';
            }, 5000); // Hide after 5 seconds
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

    // NEW: Specific message display for target management forms
    function displayTargetManagementMessage(message, isError = false) {
        if (targetManagementMessage) {
            targetManagementMessage.innerHTML = `<div class="message ${isError ? 'error' : 'success'}">${message}</div>`;
            targetManagementMessage.style.display = 'block';
            setTimeout(() => {
                targetManagementMessage.innerHTML = '';
                targetManagementMessage.style.display = 'none';
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

    // NEW: Function to fetch targets from Google Apps Script
    async function fetchTargets() {
        displayTargetManagementMessage("Fetching targets...", 'info');
        try {
            const response = await fetch(`${WEB_APP_URL}?action=get_targets`);
            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`Failed to fetch targets. Status: ${response.status}. Details: ${errorText}`);
            }
            const data = await response.json();
            if (data.status === 'success') {
                currentTargets = data.targets;
                TARGETS = currentTargets; // Update the global TARGETS object
                console.log('--- Fetched Targets: ---');
                console.log(currentTargets);
                displayTargetManagementMessage("Targets loaded successfully!", 'success');
                // Re-render targets if on the target management tab
                if (targetManagementSection.style.display === 'block') {
                    renderCurrentTargets();
                }
            } else {
                throw new Error(data.message || 'Unknown error fetching targets.');
            }
        } catch (error) {
            console.error('Error fetching targets:', error);
            displayTargetManagementMessage(`Failed to load targets: ${error.message}.`, true);
            // Fallback to default targets if fetch fails
            TARGETS = {
                'Branch Manager': {
                    'Visit': 10,
                    'Call': 3 * MONTHLY_WORKING_DAYS,
                    'Reference': 1 * MONTHLY_WORKING_DAYS,
                    'New Customer Leads': 20
                },
                'Investment Staff': {
                    'Visit': 30,
                    'Call': 5 * MONTHLY_WORKING_DAYS,
                    'Reference': 1 * MONTHLY_WORKING_DAYS,
                    'New Customer Leads': 20
                },
                'Seniors': {
                    'Visit': 30,
                    'Call': 5 * MONTHLY_WORKING_DAYS,
                    'Reference': 1 * MONTHLY_WORKING_DAYS,
                    'New Customer Leads': 20
                },
                'Default': {
                    'Visit': 5,
                    'Call': 3 * MONTHLY_WORKING_DAYS,
                    'Reference': 1 * MONTHLY_WORKING_DAYS,
                    'New Customer Leads': 20
                }
            };
            currentTargets = TARGETS; // Set currentTargets to default as well
            console.warn('Using default hardcoded targets as fallback due to fetch error.');
            if (targetManagementSection.style.display === 'block') {
                renderCurrentTargets();
            }
        }
    }

    // Generic function to send data to Google Apps Script
    async function sendDataToGoogleAppsScript(action, data) {
        displayMessage(`Sending data for ${action}...`, 'info');
        try {
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                mode: 'cors', // Crucial for cross-origin requests
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: new URLSearchParams({
                    action: action,
                    data: JSON.stringify(data)
                }).toString()
            });

            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`HTTP error! Status: ${response.status}. Details: ${errorText}`);
            }

            const result = await response.json();
            if (result.status === 'success') {
                displayMessage(`${action.replace(/_/g, ' ')} successful!`, 'success');
                return true;
            } else {
                displayMessage(`${action.replace(/_/g, ' ')} failed: ${result.message}`, 'error');
                return false;
            }
        } catch (error) {
            console.error(`Error in sendDataToGoogleAppsScript for action ${action}:`, error);
            displayMessage(`Error during ${action.replace(/_/g, ' ')}: ${error.message}.`, 'error');
            return false;
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
        await fetchTargets(); // NEW: Fetch targets after canvassing data

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

        // Populate target designation dropdown
        const uniqueDesignations = [...new Set(Object.keys(TARGETS).filter(d => d !== 'Default'))].sort();
        populateDropdown(targetDesignationSelect, uniqueDesignations);

        console.log('Final All Unique Branches (Predefined):', allUniqueBranches);
        console.log('Final Employee Code To Name Map (from Canvassing Data):', employeeCodeToNameMap);
        console.log('Final Employee Code To Designation Map (from Canvassing Data):', employeeCodeToDesignationMap);
        console.log('Final All Unique Employees (Codes from Canvassing Data):', allUniqueEmployees);
        console.log('Initial TARGETS (after fetch or fallback):', TARGETS);

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
                option.value = item; // item is branch name or designation
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

    // NEW: Render Non-Participating Branches Report
    function renderNonParticipatingBranches() {
        reportDisplay.innerHTML = '<h2>Non-Participating Branches</h2>';
        const nonParticipatingBranches = [];

        PREDEFINED_BRANCHES.forEach(branch => {
            const hasActivity = allCanvassingData.some(entry => entry[HEADER_BRANCH_NAME] === branch);
            if (!hasActivity) {
                nonParticipatingBranches.push(branch);
            }
        });

        if (nonParticipatingBranches.length > 0) {
            const ul = document.createElement('ul');
            ul.className = 'non-participating-branch-list';
            nonParticipatingBranches.forEach(branch => {
                const li = document.createElement('li');
                li.textContent = branch;
                ul.appendChild(li);
            });
            reportDisplay.appendChild(ul);
        } else {
            reportDisplay.innerHTML += '<p class="no-participation-message">All predefined branches have recorded activity!</p>';
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
        let headerRow = thead.insertRow(); // Main Headers
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
            const targets = TARGETS[designation] || TARGETS['Default']; // Use dynamically fetched TARGETS
            const performance = calculatePerformance(totalActivity, targets);

            const row = tbody.insertRow();
            row.insertCell().textContent = employeeName;
            row.insertCell().textContent = branchName;
            row.insertCell().textContent = designation;

            metrics.forEach(metric => {
                const actual = totalActivity[metric] || 0;
                const target = targets[metric] || 0;
                const percentage = target > 0 ? ((actual / target) * 100).toFixed(1) : (actual > 0 ? 100 : 0).toFixed(1);

                row.insertCell().textContent = actual;
                row.insertCell().textContent = target;
                
                const percentageCell = row.insertCell();
                const progressBarContainer = document.createElement('div');
                progressBarContainer.className = 'progress-bar-container-small';
                const progressBar = document.createElement('div');
                progressBar.className = 'progress-bar';
                progressBar.style.width = `${Math.min(100, parseFloat(percentage))}%`;
                
                // Color coding for progress bar
                if (parseFloat(percentage) >= 100) {
                    progressBar.classList.add('success');
                } else if (parseFloat(percentage) >= 75) {
                    progressBar.classList.add('warning-high');
                } else if (parseFloat(percentage) >= 50) {
                    progressBar.classList.add('warning-medium');
                } else if (parseFloat(percentage) > 0) {
                    progressBar.classList.add('warning-low');
                } else {
                    progressBar.classList.add('danger');
                }
                
                progressBar.textContent = `${percentage}%`;
                progressBarContainer.appendChild(progressBar);
                percentageCell.appendChild(progressBarContainer);
            });
        });

        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }


    // Function to calculate performance percentages
    function calculatePerformance(actuals, targets) {
        const performance = {};
        for (const activity in targets) {
            const target = targets[activity];
            const actual = actuals[activity] || 0;
            performance[activity] = target > 0 ? (actual / target) * 100 : (actual > 0 ? 100 : 0);
        }
        return performance;
    }

    // Render Employee Summary (for d4.PNG)
    function renderEmployeeSummary(employeeActivities) {
        if (!employeeActivities || employeeActivities.length === 0) {
            reportDisplay.innerHTML = '<p>No activity found for this employee this month.</p>';
            return;
        }

        const employeeCode = employeeActivities[0][HEADER_EMPLOYEE_CODE];
        const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
        const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';
        
        reportDisplay.innerHTML = `<h2>Summary for ${employeeName} (${designation}) - Current Month</h2>`;

        const currentMonth = new Date().getMonth();
        const currentYear = new Date().getFullYear();

        const monthlyActivities = employeeActivities.filter(entry => {
            const entryDate = new Date(entry[HEADER_TIMESTAMP]);
            return entryDate.getMonth() === currentMonth && entryDate.getFullYear() === currentYear;
        });

        const { totalActivity, productInterests } = calculateTotalActivity(monthlyActivities);
        const targets = TARGETS[designation] || TARGETS['Default']; // Use dynamically fetched TARGETS
        const performancePercentages = calculatePerformance(totalActivity, targets);

        const summaryCard = document.createElement('div');
        summaryCard.className = 'summary-breakdown-card';

        // Activities Summary Section
        const activitiesSection = document.createElement('div');
        activitiesSection.innerHTML = `<h4>Activities Summary</h4>`;
        const activitiesTable = document.createElement('table');
        activitiesTable.innerHTML = `
            <thead>
                <tr><th>Activity Type</th><th>Actual</th><th>Target</th><th>Progress</th></tr>
            </thead>
            <tbody>
                <tr>
                    <td>Visits</td>
                    <td>${totalActivity['Visit']}</td>
                    <td>${targets['Visit'] || 0}</td>
                    <td>
                        <div class="progress-bar-container">
                            <div class="progress-bar ${getProgressBarColor(performancePercentages['Visit'])}" style="width: ${Math.min(100, performancePercentages['Visit'] || 0)}%;">
                                ${performancePercentages['Visit'] ? performancePercentages['Visit'].toFixed(1) + '%' : '0%'}
                            </div>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td>Calls</td>
                    <td>${totalActivity['Call']}</td>
                    <td>${targets['Call'] || 0}</td>
                    <td>
                        <div class="progress-bar-container">
                            <div class="progress-bar ${getProgressBarColor(performancePercentages['Call'])}" style="width: ${Math.min(100, performancePercentages['Call'] || 0)}%;">
                                ${performancePercentages['Call'] ? performancePercentages['Call'].toFixed(1) + '%' : '0%'}
                            </div>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td>References</td>
                    <td>${totalActivity['Reference']}</td>
                    <td>${targets['Reference'] || 0}</td>
                    <td>
                        <div class="progress-bar-container">
                            <div class="progress-bar ${getProgressBarColor(performancePercentages['Reference'])}" style="width: ${Math.min(100, performancePercentages['Reference'] || 0)}%;">
                                ${performancePercentages['Reference'] ? performancePercentages['Reference'].toFixed(1) + '%' : '0%'}
                            </div>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td>New Customer Leads</td>
                    <td>${totalActivity['New Customer Leads']}</td>
                    <td>${targets['New Customer Leads'] || 0}</td>
                    <td>
                        <div class="progress-bar-container">
                            <div class="progress-bar ${getProgressBarColor(performancePercentages['New Customer Leads'])}" style="width: ${Math.min(100, performancePercentages['New Customer Leads'] || 0)}%;">
                                ${performancePercentages['New Customer Leads'] ? performancePercentages['New Customer Leads'].toFixed(1) + '%' : '0%'}
                            </div>
                        </div>
                    </td>
                </tr>
            </tbody>
        `;
        activitiesSection.appendChild(activitiesTable);
        summaryCard.appendChild(activitiesSection);

        // Products Interested Section
        const productsSection = document.createElement('div');
        productsSection.innerHTML = `<h4>Products Interested By Customers</h4>`;
        if (productInterests.length > 0) {
            const ul = document.createElement('ul');
            productInterests.forEach(product => {
                const li = document.createElement('li');
                li.textContent = product;
                ul.appendChild(li);
            });
            productsSection.appendChild(ul);
        } else {
            productsSection.innerHTML += '<p>No specific product interests recorded.</p>';
        }
        summaryCard.appendChild(productsSection);

        // Overall Performance Section (simple text for now, could be a donut chart)
        const overallSection = document.createElement('div');
        overallSection.innerHTML = `<h4>Overall Monthly Performance (Rough Estimate)</h4>`;
        const overallScore = calculateOverallPerformanceScore(performancePercentages);
        overallSection.innerHTML += `<p style="font-size: 1.8em; font-weight: bold; color: ${getOverallPerformanceColor(overallScore)}; text-align: center;">${overallScore.toFixed(1)}%</p>`;
        overallSection.innerHTML += `<p style="text-align: center; font-size: 0.9em; color: #666;">(Average of activity completion percentages)</p>`;
        summaryCard.appendChild(overallSection);

        reportDisplay.appendChild(summaryCard);
    }

    // Helper to get progress bar color class
    function getProgressBarColor(percentage) {
        if (percentage >= 100) {
            return 'success';
        } else if (percentage >= 75) {
            return 'warning-high';
        } else if (percentage >= 50) {
            return 'warning-medium';
        } else if (percentage > 0) {
            return 'warning-low';
        } else {
            return 'danger';
        }
    }

    // Helper to calculate a simple overall performance score
    function calculateOverallPerformanceScore(performancePercentages) {
        let totalPercentage = 0;
        let count = 0;
        for (const activityType in performancePercentages) {
            totalPercentage += performancePercentages[activityType];
            count++;
        }
        return count > 0 ? totalPercentage / count : 0;
    }

    // View All Entries Button Handler
    viewAllEntriesBtn.addEventListener('click', () => {
        document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        viewAllEntriesBtn.classList.add('active');
        renderAllEntriesForEmployee(selectedEmployeeCodeEntries);
    });

    // Render All Entries (Employee)
    function renderAllEntriesForEmployee(entries) {
        if (!entries || entries.length === 0) {
            reportDisplay.innerHTML = '<p>No entries found for this employee.</p>';
            return;
        }

        const employeeName = employeeCodeToNameMap[entries[0][HEADER_EMPLOYEE_CODE]] || entries[0][HEADER_EMPLOYEE_CODE];
        reportDisplay.innerHTML = `<h2>All Entries for ${employeeName}</h2>`;

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';

        const table = document.createElement('table');
        table.className = 'all-entries-table'; // You might want to define styles for this

        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        const headers = [
            HEADER_DATE, HEADER_ACTIVITY_TYPE, HEADER_TYPE_OF_CUSTOMER,
            HEADER_PROSPECT_NAME, HEADER_PHONE_NUMBER_WHATSAPP, HEADER_PRODUCT_INTERESTED,
            HEADER_REMARKS, HEADER_NEXT_FOLLOW_UP_DATE
        ]; // Display a subset of relevant headers
        
        headers.forEach(headerText => {
            const th = document.createElement('th');
            th.textContent = headerText;
            headerRow.appendChild(th);
        });

        const tbody = table.createTBody();
        entries.forEach(entry => {
            const row = tbody.insertRow();
            headers.forEach(header => {
                const cell = row.insertCell();
                cell.textContent = entry[header] || ''; // Display content, or empty string if null/undefined
            });
        });

        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }

    // View Employee Performance Report (for d2.PNG - detailed report)
    viewPerformanceReportBtn.addEventListener('click', () => {
        document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        viewPerformanceReportBtn.classList.add('active');
        renderEmployeePerformanceReport(selectedEmployeeCodeEntries);
    });

    // Render Employee Performance Report (for d2.PNG)
    function renderEmployeePerformanceReport(employeeActivities) {
        if (!employeeActivities || employeeActivities.length === 0) {
            reportDisplay.innerHTML = '<p>No activities found for this employee for a performance report.</p>';
            return;
        }

        const employeeCode = employeeActivities[0][HEADER_EMPLOYEE_CODE];
        const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
        const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';
        const targets = TARGETS[designation] || TARGETS['Default']; // Use dynamically fetched TARGETS

        reportDisplay.innerHTML = `<h2>Performance Report for ${employeeName} (${designation})</h2>`;

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';

        const table = document.createElement('table');
        table.className = 'performance-table'; // Reuse performance-table class

        const thead = table.createTHead();
        let headerRow = thead.insertRow();
        headerRow.insertCell().textContent = 'Date';
        headerRow.insertCell().textContent = 'Activity Type';
        headerRow.insertCell().textContent = 'Customer Name';
        headerRow.insertCell().textContent = 'Target';
        headerRow.insertCell().textContent = 'Achieved';
        headerRow.insertCell().textContent = '%'; // Percentage

        const tbody = table.createTBody();

        // Group activities by date and type for daily totals
        const dailyActivitySummary = {}; // { 'YYYY-MM-DD': { 'Visit': {actual:X, target:Y}, 'Call': {...} } }

        employeeActivities.forEach(entry => {
            const date = formatDate(entry[HEADER_TIMESTAMP] || entry[HEADER_DATE]); // Use timestamp or date
            const activityType = entry[HEADER_ACTIVITY_TYPE];
            const typeOfCustomer = entry[HEADER_TYPE_OF_CUSTOMER];
            const prospectName = entry[HEADER_PROSPECT_NAME] || 'N/A';

            if (!dailyActivitySummary[date]) {
                dailyActivitySummary[date] = {};
            }
            // Normalize activity type to match TARGETS keys
            let normalizedActivityType = '';
            if (activityType.trim().toLowerCase() === 'visit') {
                normalizedActivityType = 'Visit';
            } else if (activityType.trim().toLowerCase() === 'calls') {
                normalizedActivityType = 'Call';
            } else if (activityType.trim().toLowerCase() === 'referance') {
                normalizedActivityType = 'Reference';
            }

            if (normalizedActivityType) {
                if (!dailyActivitySummary[date][normalizedActivityType]) {
                    dailyActivitySummary[date][normalizedActivityType] = { actual: 0, target: targets[normalizedActivityType] || 0 };
                }
                dailyActivitySummary[date][normalizedActivityType].actual++;
            }

            // Also count New Customer Leads separately if type of customer is 'new'
            if (typeOfCustomer && typeOfCustomer.trim().toLowerCase() === 'new') {
                if (!dailyActivitySummary[date]['New Customer Leads']) {
                    dailyActivitySummary[date]['New Customer Leads'] = { actual: 0, target: targets['New Customer Leads'] || 0 };
                }
                dailyActivitySummary[date]['New Customer Leads'].actual++;
            }

            // For individual entries, show them as well
            const row = tbody.insertRow();
            row.insertCell().textContent = date;
            row.insertCell().textContent = activityType;
            row.insertCell().textContent = prospectName;

            // Determine target for this specific entry (can be tricky for 'New Customer Leads' as it's not a direct activity)
            let individualTarget = '';
            let individualAchieved = '1'; // Assuming each entry means 1 achieved
            let individualPercentage = '';

            if (normalizedActivityType) {
                individualTarget = targets[normalizedActivityType] ? (targets[normalizedActivityType] / MONTHLY_WORKING_DAYS).toFixed(1) : 'N/A'; // Daily target approximation
                individualPercentage = individualTarget !== 'N/A' && parseFloat(individualTarget) > 0 ? (1 / parseFloat(individualTarget) * 100).toFixed(1) + '%' : (parseFloat(individualTarget) === 0 ? '100%' : 'N/A');
            } else if (typeOfCustomer && typeOfCustomer.trim().toLowerCase() === 'new') {
                 individualTarget = targets['New Customer Leads'] ? (targets['New Customer Leads'] / MONTHLY_WORKING_DAYS).toFixed(1) : 'N/A';
                 individualPercentage = individualTarget !== 'N/A' && parseFloat(individualTarget) > 0 ? (1 / parseFloat(individualTarget) * 100).toFixed(1) + '%' : (parseFloat(individualTarget) === 0 ? '100%' : 'N/A');
            } else {
                individualAchieved = 'N/A';
            }


            row.insertCell().textContent = individualTarget;
            row.insertCell().textContent = individualAchieved;
            row.insertCell().textContent = individualPercentage;

        });

        // Add summary rows for each day
        for (const date in dailyActivitySummary) {
            const dailyData = dailyActivitySummary[date];
            for (const activityType in dailyData) {
                const data = dailyData[activityType];
                const percentage = data.target > 0 ? ((data.actual / data.target) * 100).toFixed(1) : (data.actual > 0 ? 100 : 0).toFixed(1);

                const summaryRow = tbody.insertRow();
                summaryRow.style.fontWeight = 'bold';
                summaryRow.style.backgroundColor = '#e0e0e0';
                summaryRow.insertCell().textContent = date;
                summaryRow.insertCell().textContent = `${activityType} (Daily Total)`;
                summaryRow.insertCell(); // Empty for customer name
                summaryRow.insertCell().textContent = data.target;
                summaryRow.insertCell().textContent = data.actual;
                summaryRow.insertCell().textContent = `${percentage}%`;
            }
        }

        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }

    // View Branch Performance Report (for d3.PNG)
    viewBranchPerformanceReportBtn.addEventListener('click', () => {
        document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        viewBranchPerformanceReportBtn.classList.add('active');
        renderBranchPerformanceReport();
    });

    function renderBranchPerformanceReport() {
        const selectedBranch = branchSelect.value;
        if (!selectedBranch) {
            reportDisplay.innerHTML = '<p>Please select a branch to view its performance report.</p>';
            return;
        }

        reportDisplay.innerHTML = `<h2>Performance Report for ${selectedBranch}</h2>`;

        const branchActivityEntries = allCanvassingData.filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch);
        if (branchActivityEntries.length === 0) {
            reportDisplay.innerHTML += '<p>No activity found for this branch.</p>';
            return;
        }

        const employeesInBranch = [...new Set(branchActivityEntries.map(entry => entry[HEADER_EMPLOYEE_CODE]))].sort((codeA, codeB) => {
            const nameA = employeeCodeToNameMap[codeA] || codeA;
            const nameB = employeeCodeToNameMap[codeB] || codeB;
            return nameA.localeCompare(nameB);
        });

        const performanceGrid = document.createElement('div');
        performanceGrid.className = 'branch-performance-grid';

        employeesInBranch.forEach(employeeCode => {
            const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
            const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';
            const targets = TARGETS[designation] || TARGETS['Default']; // Use dynamically fetched TARGETS

            const employeeActivities = branchActivityEntries.filter(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode);
            const { totalActivity } = calculateTotalActivity(employeeActivities);
            const performancePercentages = calculatePerformance(totalActivity, targets);

            const employeeCard = document.createElement('div');
            employeeCard.className = 'employee-performance-card';
            employeeCard.innerHTML = `<h4>${employeeName} (${designation})</h4>`;

            const cardTable = document.createElement('table');
            cardTable.innerHTML = `
                <thead>
                    <tr><th>Activity</th><th>Actual</th><th>Target</th><th>%</th></tr>
                </thead>
                <tbody>
                    <tr><td>Visits</td><td>${totalActivity['Visit']}</td><td>${targets['Visit'] || 0}</td><td>${performancePercentages['Visit'] ? performancePercentages['Visit'].toFixed(1) + '%' : '0%'}</td></tr>
                    <tr><td>Calls</td><td>${totalActivity['Call']}</td><td>${targets['Call'] || 0}</td><td>${performancePercentages['Call'] ? performancePercentages['Call'].toFixed(1) + '%' : '0%'}</td></tr>
                    <tr><td>References</td><td>${totalActivity['Reference']}</td><td>${targets['Reference'] || 0}</td><td>${performancePercentages['Reference'] ? performancePercentages['Reference'].toFixed(1) + '%' : '0%'}</td></tr>
                    <tr><td>New Customer Leads</td><td>${totalActivity['New Customer Leads']}</td><td>${targets['New Customer Leads'] || 0}</td><td>${performancePercentages['New Customer Leads'] ? performancePercentages['New Customer Leads'].toFixed(1) + '%' : '0%'}</td></tr>
                </tbody>
            `;
            employeeCard.appendChild(cardTable);
            performanceGrid.appendChild(employeeCard);
        });

        reportDisplay.appendChild(performanceGrid);
    }

    // NEW: Detailed Customer View Logic
    customerViewBranchSelect.addEventListener('change', () => {
        const selectedBranch = customerViewBranchSelect.value;
        customerCanvassedList.innerHTML = '<p>Select an employee to view their canvassed customers.</p>';
        customerDetailsContent.style.display = 'none';

        if (selectedBranch) {
            const employeeCodesInBranchFromCanvassing = allCanvassingData
                .filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch)
                .map(entry => entry[HEADER_EMPLOYEE_CODE]);
            const uniqueEmployeeCodes = [...new Set(employeeCodesInBranchFromCanvassing)].sort((codeA, codeB) => {
                const nameA = employeeCodeToNameMap[codeA] || codeA;
                const nameB = employeeCodeToNameMap[codeB] || codeB;
                return nameA.localeCompare(nameB);
            });
            populateDropdown(customerViewEmployeeSelect, uniqueEmployeeCodes, true);
        } else {
            populateDropdown(customerViewEmployeeSelect, []); // Clear employee dropdown
        }
    });

    customerViewEmployeeSelect.addEventListener('change', () => {
        renderCustomerCanvassedList();
    });

    function renderCustomerCanvassedList() {
        const selectedBranch = customerViewBranchSelect.value;
        const selectedEmployeeCode = customerViewEmployeeSelect.value;

        if (!selectedBranch || !selectedEmployeeCode) {
            customerCanvassedList.innerHTML = '<p>Please select both a branch and an employee.</p>';
            customerDetailsContent.style.display = 'none';
            return;
        }

        const employeeCustomers = allCanvassingData.filter(entry =>
            entry[HEADER_BRANCH_NAME] === selectedBranch &&
            entry[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode &&
            entry[HEADER_PROSPECT_NAME] && entry[HEADER_PROSPECT_NAME].trim() !== ''
        );

        customerDetailsContent.style.display = 'none'; // Hide details when list is refreshed

        if (employeeCustomers.length === 0) {
            customerCanvassedList.innerHTML = `<p>No canvassed customers found for ${employeeCodeToNameMap[selectedEmployeeCode] || selectedEmployeeCode} in ${selectedBranch}.</p>`;
            return;
        }

        customerCanvassedList.innerHTML = `<h3>Canvassed Customers for ${employeeCodeToNameMap[selectedEmployeeCode] || selectedEmployeeCode}</h3>`;

        const ul = document.createElement('ul');
        ul.className = 'customer-list'; // Add a class for styling

        employeeCustomers.forEach((customer, index) => {
            const li = document.createElement('li');
            li.textContent = customer[HEADER_PROSPECT_NAME];
            li.dataset.customerIndex = index; // Store index for easy lookup
            li.addEventListener('click', () => displayCustomerDetails(customer));
            ul.appendChild(li);
        });
        customerCanvassedList.appendChild(ul);
    }

    function displayCustomerDetails(customer) {
        customerDetailsContent.style.display = 'block';
        customerDetailsContent.innerHTML = `<h3>Details for ${customer[HEADER_PROSPECT_NAME]}</h3>`;

        const detailsHtml = `
            <div class="detail-row"><span class="detail-label">Prospect Name:</span> <span class="detail-value">${customer[HEADER_PROSPECT_NAME] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Phone Number (Whatsapp):</span> <span class="detail-value">${customer[HEADER_PHONE_NUMBER_WHATSAPP] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Address:</span> <span class="detail-value">${customer[HEADER_ADDRESS] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Profession:</span> <span class="detail-value">${customer[HEADER_PROFESSION] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">DOB/WD:</span> <span class="detail-value">${customer[HEADER_DOB_WD] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Product Interested:</span> <span class="detail-value">${customer[HEADER_PRODUCT_INTERESTED] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Remarks:</span> <span class="detail-value">${customer[HEADER_REMARKS] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Next Follow-up Date:</span> <span class="detail-value">${customer[HEADER_NEXT_FOLLOW_UP_DATE] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Relation With Staff:</span> <span class="detail-value">${customer[HEADER_RELATION_WITH_STAFF] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Family Details (Wife/Husband Name):</span> <span class="detail-value">${customer[HEADER_FAMILY_DETAILS_1] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Family Details (Wife/Husband Job):</span> <span class="detail-value">${customer[HEADER_FAMILY_DETAILS_2] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Family Details (Children Names):</span> <span class="detail-value">${customer[HEADER_FAMILY_DETAILS_3] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Family Details (Children Details):</span> <span class="detail-value">${customer[HEADER_FAMILY_DETAILS_4] || 'N/A'}</span></div>
            <div class="detail-row"><span class="detail-label">Profile of Customer:</span> <span class="detail-value">${customer[HEADER_PROFILE_OF_CUSTOMER] || 'N/A'}</span></div>
        `;
        customerDetailsContent.innerHTML += `<div class="customer-details-card">${detailsHtml}</div>`; // Add a class for card styling
    }


    // Function to show/hide sections based on tab clicked
    function showTab(tabId) {
        // Remove 'active' class from all tab buttons
        document.querySelectorAll('.tab-button').forEach(button => {
            button.classList.remove('active');
        });

        // Hide all main content sections
        reportsSection.style.display = 'none';
        detailedCustomerViewSection.style.display = 'none';
        employeeManagementSection.style.display = 'none';
        targetManagementSection.style.display = 'none'; // NEW

        // Add 'active' class to the clicked button and show the corresponding section
        switch (tabId) {
            case 'allBranchSnapshotTabBtn':
                allBranchSnapshotTabBtn.classList.add('active');
                reportsSection.style.display = 'block';
                renderAllBranchSnapshot(); // Re-render content when tab is clicked
                break;
            case 'allStaffOverallPerformanceTabBtn':
                allStaffOverallPerformanceTabBtn.classList.add('active');
                reportsSection.style.display = 'block';
                renderOverallStaffPerformanceReport(); // Re-render content
                break;
            case 'nonParticipatingBranchesTabBtn':
                nonParticipatingBranchesTabBtn.classList.add('active');
                reportsSection.style.display = 'block';
                renderNonParticipatingBranches(); // Re-render content
                break;
            case 'detailedCustomerViewTabBtn':
                detailedCustomerViewTabBtn.classList.add('active');
                detailedCustomerViewSection.style.display = 'block';
                // Reset/clear customer view when tab is selected
                populateDropdown(customerViewBranchSelect, allUniqueBranches);
                populateDropdown(customerViewEmployeeSelect, []);
                customerCanvassedList.innerHTML = '<p>Select a branch and an employee to view their canvassed customers.</p>';
                customerDetailsContent.style.display = 'none';
                break;
            case 'employeeManagementTabBtn':
                employeeManagementTabBtn.classList.add('active');
                employeeManagementSection.style.display = 'block';
                displayEmployeeManagementMessage(''); // Clear any previous messages
                break;
            case 'targetManagementTabBtn': // NEW
                targetManagementTabBtn.classList.add('active');
                targetManagementSection.style.display = 'block';
                displayTargetManagementMessage(''); // Clear any previous messages
                renderCurrentTargets(); // Display current targets when tab is opened
                break;
            default:
                allBranchSnapshotTabBtn.classList.add('active');
                reportsSection.style.display = 'block';
                renderAllBranchSnapshot();
                break;
        }
    }

    // Event Listeners for Tab Buttons
    allBranchSnapshotTabBtn.addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    allStaffOverallPerformanceTabBtn.addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    nonParticipatingBranchesTabBtn.addEventListener('click', () => showTab('nonParticipatingBranchesTabBtn'));
    detailedCustomerViewTabBtn.addEventListener('click', () => showTab('detailedCustomerViewTabBtn'));
    employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));
    targetManagementTabBtn.addEventListener('click', () => showTab('targetManagementTabBtn')); // NEW


    // Employee Management Form Submissions
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const employeeName = newEmployeeNameInput.value.trim();
            const employeeCode = newEmployeeCodeInput.value.trim();
            const branchName = newBranchNameInput.value.trim();
            const designation = newDesignationInput.value.trim();

            if (!employeeName || !employeeCode || !branchName || !designation) {
                displayEmployeeManagementMessage('All fields are required for adding an employee.', true);
                return;
            }

            const employeeData = {
                [HEADER_EMPLOYEE_NAME]: employeeName,
                [HEADER_EMPLOYEE_CODE]: employeeCode,
                [HEADER_BRANCH_NAME]: branchName,
                [HEADER_DESIGNATION]: designation
            };

            const success = await sendDataToGoogleAppsScript('add_employee', employeeData);

            if (success) {
                addEmployeeForm.reset();
                // Re-process data to update dropdowns with new employee
                await processData(); 
            }
        });
    }

    if (bulkAddEmployeeForm) {
        bulkAddEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const branchName = bulkEmployeeBranchNameInput.value.trim();
            const employeeDetails = bulkEmployeeDetailsTextarea.value.trim();

            if (!branchName || !employeeDetails) {
                displayEmployeeManagementMessage('Branch Name and Employee Details are required for bulk addition.', true);
                return;
            }

            const lines = employeeDetails.split('\n').filter(line => line.trim() !== '');
            const employeesToAdd = [];

            for (const line of lines) {
                const parts = line.split(',').map(part => part.trim());
                if (parts.length >= 2) { // Name, Code, (Designation is optional here)
                    const employeeData = {
                        [HEADER_EMPLOYEE_NAME]: parts[0],
                        [HEADER_EMPLOYEE_CODE]: parts[1],
                        [HEADER_BRANCH_NAME]: branchName,
                        [HEADER_DESIGNATION]: parts[2] || '' // Designation can be empty if not provided
                    };
                    employeesToAdd.push(employeeData);
                }
            }

            if (employeesToAdd.length > 0) {
                const success = await sendDataToGoogleAppsScript('add_bulk_employees', employeesToAdd);
                if (success) {
                    bulkAddEmployeeForm.reset();
                    // Re-process data to update dropdowns with new employees
                    await processData();
                }
            } else {
                displayEmployeeManagementMessage('No valid employee entries found in the bulk details.', true);
            }
        });
    }

    // Event Listener for Delete Employee Form
    if (deleteEmployeeForm) {
        deleteEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const employeeCodeToDelete = deleteEmployeeCodeInput.value.trim();

            if (!employeeCodeToDelete) {
                displayEmployeeManagementMessage('Employee Code is required for deletion.', true);
                return;
            }

            const deleteData = { [HEADER_EMPLOYEE_CODE]: employeeCodeToDelete };
            const success = await sendDataToGoogleAppsScript('delete_employee', deleteData);

            if (success) {
                deleteEmployeeForm.reset();
                // Re-process data to update dropdowns
                await processData();
            }
        });
    }

    // NEW: Target Management Form Submission
    if (setTargetsForm) {
        setTargetsForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const designation = targetDesignationSelect.value;
            const visitTarget = parseInt(targetVisitInput.value);
            const callTarget = parseInt(targetCallInput.value);
            const referenceTarget = parseInt(targetReferenceInput.value);
            const newCustomerLeadsTarget = parseInt(targetNewCustomerLeadsInput.value);

            if (!designation) {
                displayTargetManagementMessage('Please select a designation.', true);
                return;
            }
            if (isNaN(visitTarget) || isNaN(callTarget) || isNaN(referenceTarget) || isNaN(newCustomerLeadsTarget)) {
                displayTargetManagementMessage('All target fields must be valid numbers.', true);
                return;
            }

            const targetData = {
                designation: designation,
                Visit: visitTarget,
                Call: callTarget,
                Reference: referenceTarget,
                'New Customer Leads': newCustomerLeadsTarget
            };

            // Send target data to Google Apps Script
            const success = await sendDataToGoogleAppsScript('set_target', targetData);

            if (success) {
                // Re-fetch targets to update the display and global TARGETS object
                await fetchTargets();
                renderCurrentTargets(); // Re-render immediately
            }
        });
    }

    // NEW: Populate target form with existing targets when designation changes
    targetDesignationSelect.addEventListener('change', () => {
        const selectedDesignation = targetDesignationSelect.value;
        if (selectedDesignation && currentTargets[selectedDesignation]) {
            const targets = currentTargets[selectedDesignation];
            targetVisitInput.value = targets['Visit'] || 0;
            targetCallInput.value = targets['Call'] || 0;
            targetReferenceInput.value = targets['Reference'] || 0;
            targetNewCustomerLeadsInput.value = targets['New Customer Leads'] || 0;
        } else {
            // Clear inputs if no designation selected or no targets found
            targetVisitInput.value = 0;
            targetCallInput.value = 0;
            targetReferenceInput.value = 0;
            targetNewCustomerLeadsInput.value = 0;
        }
        renderCurrentTargets(); // Re-render current targets display to highlight selected
    });

    // NEW: Function to render current targets overview
    function renderCurrentTargets() {
        currentTargetsDisplay.innerHTML = '<h3>Designation Targets</h3>';
        if (Object.keys(currentTargets).length === 0) {
            currentTargetsDisplay.innerHTML += '<p>No targets set. Please set targets using the form above.</p>';
            return;
        }

        const table = document.createElement('table');
        table.className = 'targets-table'; // Add a new class for styling

        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        ['Designation', 'Visits', 'Calls', 'References', 'New Customer Leads'].forEach(text => {
            const th = document.createElement('th');
            th.textContent = text;
            headerRow.appendChild(th);
        });

        const tbody = table.createTBody();
        // Sort designations, putting 'Default' at the end
        const sortedDesignations = Object.keys(currentTargets).sort((a, b) => {
            if (a === 'Default') return 1;
            if (b === 'Default') return -1;
            return a.localeCompare(b);
        });

        sortedDesignations.forEach(designation => {
            const targets = currentTargets[designation];
            const row = tbody.insertRow();
            if (designation === targetDesignationSelect.value) {
                row.classList.add('highlight-row'); // Highlight selected designation
            }
            row.insertCell().textContent = designation;
            row.insertCell().textContent = targets['Visit'] || 0;
            row.insertCell().textContent = targets['Call'] || 0;
            row.insertCell().textContent = targets['Reference'] || 0;
            row.insertCell().textContent = targets['New Customer Leads'] || 0;
        });
        currentTargetsDisplay.appendChild(table);
    }


    // Initial data fetch and tab display when the page loads
    processData();
    showTab('allBranchSnapshotTabBtn');
});
