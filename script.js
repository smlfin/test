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

    const TARGETS = {
        'Branch Manager': {
            'Visit': 10,
            'Call': 3 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 20
        },
        'Investment Staff': { // Added Investment Staff with custom Visit target
            'Visit': 15,
            'Call': 5 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 20
        },
        'Seniors': { // Added Investment Staff with custom Visit target
            'Visit': 15,
            'Call': 5 * MONTHLY_WORKING_DAYS,
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
    const HEADER_TYPE_OF_CUSTOMER = 'Type of Customer';
    const HEADER_R_LEAD_SOURCE = 'rLead Source';
    const HEADER_HOW_CONTACTED = 'How Contacted';
    const HEADER_PROSPECT_NAME = 'Prospect Name';
    const HEADER_PHONE_NUMBER_WHATSAPP = 'Phone Numebr(Whatsapp)';
    const HEADER_ADDRESS = 'Address';
    const HEADER_PROFESSION = 'Profession';
    const HEADER_DOB_WD = 'DOB/WD';
    const HEADER_PRODUCT_INTERESTED = 'Prodcut Interested';
    const HEADER_REMARKS = 'Remarks';
    const HEADER_NEXT_FOLLOW_UP_DATE = 'Next Follow-up Date';
    const HEADER_RELATION_WITH_STAFF = 'Relation With Staff';


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
    // NEW: Dedicated message area element
    const statusMessageDiv = document.getElementById('statusMessage');


    // Tab buttons for main navigation
    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');
    const nonParticipatingBranchesTabBtn = document.getElementById('nonParticipatingBranchesTabBtn');
    const dailyFollowUpReportTabBtn = document.getElementById('dailyFollowUpReportTabBtn'); // NEW DOM ELEMENT
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');

    // Main Content Sections to toggle
    const reportsSection = document.getElementById('reportsSection');
    const employeeManagementSection = document.getElementById('employeeManagementSection');

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


    // Global variables to store fetched data
    let allCanvassingData = []; // Raw activity data from Form Responses 2
    let allUniqueBranches = []; // Will be populated from PREDEFINED_BRANCHES
    let allUniqueEmployees = []; // Employee codes from Canvassing Data
    let employeeCodeToNameMap = {}; // {code: name} from Canvassing Data
    let employeeCodeToDesignationMap = {}; // {code: designation} from Canvassing Data
    let selectedBranchEntries = []; // Activity entries filtered by branch
    let selectedEmployeeCodeEntries = []; // Activity entries filtered by employee code

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
            const nameA = employeeCodeToNameMap[codeA] || codeA;
            const nameB = employeeCodeToNameMap[codeB] || codeB;
            return nameA.localeCompare(nameB);
        });

        populateDropdown(branchSelect, allUniqueBranches); // Populate branch dropdown with predefined branches
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
            console.log('--- End Entry ${index + 1} Debug ---');
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
        tableContainer.style.overflowX = 'auto'; // Make table scrollable on small screens
        const table = document.createElement('table');
        table.className = 'performance-table';

        // Create Header (d1.PNG style with merged cells)
        const thead = table.createTHead();
        const headerRow1 = thead.insertRow();
        const headerRow2 = thead.insertRow();

        // First row headers with rowspan
        const mainHeaders = ['Employee Name', 'Branch', 'Designation'];
        mainHeaders.forEach(text => {
            const th = document.createElement('th');
            th.textContent = text;
            th.rowSpan = 2; // Span across two rows
            headerRow1.appendChild(th);
        });

        // Metrics headers with colspan for the first row
        const metrics = ['Visit', 'Call', 'Reference', 'New Customer Leads'];
        metrics.forEach(metric => {
            const th = document.createElement('th');
            th.textContent = metric;
            th.colSpan = 3; // Span across three columns (Act, Tgt, %)
            headerRow1.appendChild(th);
        });

        // Second row headers (Act, Tgt, %)
        metrics.forEach(() => {
            ['Act', 'Tgt', '%'].forEach(subHeader => {
                const th = document.createElement('th');
                th.textContent = subHeader;
                headerRow2.appendChild(th);
            });
        });

        const tbody = table.createTBody();

        if (allUniqueEmployees.length === 0) {
            const row = tbody.insertRow();
            const cell = row.insertCell();
            cell.colSpan = 3 + (metrics.length * 3); // Span across all columns
            cell.textContent = 'No employee activity data found to generate overall staff performance report.';
            cell.style.textAlign = 'center';
            tbody.appendChild(row);
        } else {
            allUniqueEmployees.forEach(employeeCode => {
                const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
                const employeeDesignation = employeeCodeToDesignationMap[employeeCode] || 'Default';

                const employeeEntries = allCanvassingData.filter(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode);
                const { totalActivity } = calculateTotalActivity(employeeEntries);

                // Determine the branch of the employee based on their latest entry or a consistent branch if available
                const branchForEmployee = employeeEntries.length > 0 ? employeeEntries[0][HEADER_BRANCH_NAME] : 'N/A';

                const row = tbody.insertRow();
                row.insertCell().textContent = employeeName;
                row.insertCell().textContent = branchForEmployee;
                row.insertCell().textContent = employeeDesignation;

                metrics.forEach(metric => {
                    const actual = totalActivity[metric];
                    const target = TARGETS[employeeDesignation] ? TARGETS[employeeDesignation][metric] : TARGETS['Default'][metric];
                    const performancePercentage = target > 0 ? (actual / target) * 100 : 0;

                    row.insertCell().textContent = actual;
                    row.insertCell().textContent = target;
                    
                    const percentCell = row.insertCell();
                    const progressBarContainer = document.createElement('div');
                    progressBarContainer.className = 'performance-progress-bar-container';
                    const progressBar = document.createElement('div');
                    progressBar.className = `performance-progress-bar ${getProgressBarClass(performancePercentage)}`;
                    progressBar.style.width = `${Math.min(100, performancePercentage)}%`;
                    progressBar.textContent = `${Math.round(performancePercentage)}%`;

                    progressBarContainer.appendChild(progressBar);
                    percentCell.appendChild(progressBarContainer);
                });
            });
        }
        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }

    // Helper to get progress bar class based on percentage
    function getProgressBarClass(percentage) {
        if (percentage >= 100) return 'progress-green';
        if (percentage >= 75) return 'progress-yellow';
        if (percentage >= 50) return 'progress-orange';
        return 'progress-red';
    }


    // Render Employee Detailed Entries (d5.PNG) - View All Entries for a Selected Employee
    function renderEmployeeDetailedEntries(entries) {
        reportDisplay.innerHTML = '<h2>Detailed Employee Entries</h2>';
        if (entries.length === 0) {
            reportDisplay.innerHTML += '<p class="no-participation-message">No detailed entries found for this employee.</p>';
            return;
        }

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';

        const table = document.createElement('table');
        table.className = 'data-table employee-detailed-entries-table';

        const thead = table.createTHead();
        const headerRow = thead.insertRow();

        // Dynamically get headers from the first entry, or a predefined list
        const headers = [
            HEADER_TIMESTAMP, HEADER_DATE, HEADER_BRANCH_NAME, HEADER_EMPLOYEE_NAME,
            HEADER_EMPLOYEE_CODE, HEADER_DESIGNATION, HEADER_ACTIVITY_TYPE,
            HEADER_TYPE_OF_CUSTOMER, HEADER_R_LEAD_SOURCE, HEADER_HOW_CONTACTED,
            HEADER_PROSPECT_NAME, HEADER_PHONE_NUMBER_WHATSAPP, HEADER_ADDRESS,
            HEADER_PROFESSION, HEADER_DOB_WD, HEADER_PRODUCT_INTERESTED,
            HEADER_REMARKS, HEADER_NEXT_FOLLOW_UP_DATE, HEADER_RELATION_WITH_STAFF
        ];

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
                let value = entry[header] || ''; // Handle undefined values
                
                // Format dates specifically
                if (header === HEADER_DATE || header === HEADER_NEXT_FOLLOW_UP_DATE || header === HEADER_TIMESTAMP) {
                    value = formatDate(value); // Use the utility to format dates
                }
                
                cell.textContent = value;
                cell.setAttribute('data-label', header); // For mobile responsiveness
            });
        });

        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }

    // Render Employee Summary (d4.PNG) - default for selected employee
    function renderEmployeeSummary(entries) {
        reportDisplay.innerHTML = '<h2>Employee Activity Summary (Current Month)</h2>';

        if (entries.length === 0) {
            reportDisplay.innerHTML += '<p class="no-participation-message">No activity found for this employee this month.</p>';
            return;
        }

        const { totalActivity, productInterests } = calculateTotalActivity(entries);
        const employeeCode = entries[0][HEADER_EMPLOYEE_CODE];
        const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
        const employeeDesignation = employeeCodeToDesignationMap[employeeCode] || 'Default';
        const branchName = entries[0][HEADER_BRANCH_NAME];

        // Overall Summary Card (d4.PNG style)
        const summaryCard = document.createElement('div');
        summaryCard.className = 'summary-breakdown-card'; // Reusing this class for general breakdown display

        // Employee Info Section
        const employeeInfoDiv = document.createElement('div');
        employeeInfoDiv.innerHTML = '<h4>Employee Information</h4>';
        employeeInfoDiv.innerHTML += `<ul class="summary-list">
            <li><strong>Name:</strong> ${employeeName}</li>
            <li><strong>Code:</strong> ${employeeCode}</li>
            <li><strong>Branch:</strong> ${branchName}</li>
            <li><strong>Designation:</strong> ${employeeDesignation}</li>
        </ul>`;
        summaryCard.appendChild(employeeInfoDiv);

        // Activity Summary Section
        const activitySummaryDiv = document.createElement('div');
        activitySummaryDiv.innerHTML = '<h4>Activity Breakdown</h4>';
        activitySummaryDiv.innerHTML += `<ul class="summary-list">
            <li><strong>Total Visits:</strong> ${totalActivity['Visit']}</li>
            <li><strong>Total Calls:</strong> ${totalActivity['Call']}</li>
            <li><strong>Total References:</strong> ${totalActivity['Reference']}</li>
            <li><strong>New Customer Leads:</strong> ${totalActivity['New Customer Leads']}</li>
        </ul>`;
        summaryCard.appendChild(activitySummaryDiv);

        // Product Interest Section
        if (productInterests.length > 0) {
            const productInterestDiv = document.createElement('div');
            productInterestDiv.innerHTML = '<h4>Products Interested</h4>';
            productInterestDiv.innerHTML += `<ul class="summary-list">
                ${productInterests.map(p => `<li>${p}</li>`).join('')}
            </ul>`;
            summaryCard.appendChild(productInterestDiv);
        }

        reportDisplay.appendChild(summaryCard);
    }


    // Render Employee Performance Report (d3.PNG) - for a selected employee
    function renderPerformanceReport(entries) {
        reportDisplay.innerHTML = '<h2>Employee Performance Report (This Month)</h2>';

        if (entries.length === 0) {
            reportDisplay.innerHTML += '<p class="no-participation-message">No activity data found for this employee to generate a performance report.</p>';
            return;
        }

        const employeeCode = entries[0][HEADER_EMPLOYEE_CODE];
        const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
        const employeeDesignation = employeeCodeToDesignationMap[employeeCode] || 'Default';
        const branchName = entries[0][HEADER_BRANCH_NAME];

        const { totalActivity } = calculateTotalActivity(entries);

        // Employee Info Card
        const employeeInfoCard = document.createElement('div');
        employeeInfoCard.className = 'employee-summary-card'; // Reusing the summary card style
        employeeInfoCard.innerHTML = `<h4>${employeeName} (${employeeCode}) - ${branchName}</h4>
                                      <p>Designation: <strong>${employeeDesignation}</strong></p>`;
        reportDisplay.appendChild(employeeInfoCard);


        // Performance Grid (d3.PNG style)
        const performanceGrid = document.createElement('div');
        performanceGrid.className = 'branch-performance-grid'; // Reusing grid for single employee metrics

        const metrics = ['Visit', 'Call', 'Reference', 'New Customer Leads'];

        metrics.forEach(metric => {
            const actual = totalActivity[metric];
            const target = TARGETS[employeeDesignation] ? TARGETS[employeeDesignation][metric] : TARGETS['Default'][metric];
            const performancePercentage = target > 0 ? (actual / target) * 100 : 0;

            const metricCard = document.createElement('div');
            metricCard.className = 'employee-performance-card';
            metricCard.innerHTML = `<h4>${metric} Performance</h4>
                                    <p>Actual: <strong>${actual}</strong></p>
                                    <p>Target: <strong>${target}</strong></p>
                                    <p>Achieved: <strong>${Math.round(performancePercentage)}%</strong></p>
                                    <div class="performance-progress-bar-container">
                                        <div class="performance-progress-bar ${getProgressBarClass(performancePercentage)}" style="width: ${Math.min(100, performancePercentage)}%;"></div>
                                    </div>`;
            performanceGrid.appendChild(metricCard);
        });

        reportDisplay.appendChild(performanceGrid);
    }


    // Function to send data to Google Apps Script (e.g., for employee management)
    async function sendDataToGoogleAppsScript(action, data) {
        displayEmployeeManagementMessage('Processing request...', false);
        try {
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                mode: 'cors', // Crucial for cross-origin requests
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ action, data })
            });

            if (!response.ok) {
                const errorData = await response.json();
                console.error('Apps Script Error:', errorData.error);
                displayEmployeeManagementMessage(`Error: ${errorData.error}`, true);
                return false;
            }

            const result = await response.json();
            if (result.status === 'SUCCESS') {
                displayEmployeeManagementMessage(result.message, false);
                await processData(); // Re-fetch and re-process data to update reports
                return true;
            } else {
                displayEmployeeManagementMessage(`Failed: ${result.message}`, true);
                return false;
            }
        } catch (error) {
            console.error('Network or Fetch Error:', error);
            displayEmployeeManagementMessage(`Network error or script not deployed correctly: ${error.message}`, true);
            return false;
        }
    }


    // --- Event Listeners for main report buttons ---
    // viewBranchSummaryBtn.addEventListener('click', () => { // Removed as per request
    //     document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
    //     viewBranchSummaryBtn.classList.add('active');
    //     renderBranchSummary(selectedBranchEntries);
    // });
    
    viewBranchPerformanceReportBtn.addEventListener('click', () => {
        if (!branchSelect.value) {
            displayMessage('Please select a branch first.', 'error');
            return;
        }
        document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        viewBranchPerformanceReportBtn.classList.add('active');
        // Filter allCanvassingData for the selected branch
        selectedBranchEntries = allCanvassingData.filter(entry => entry[HEADER_BRANCH_NAME] === branchSelect.value);
        renderPerformanceReportByBranch(selectedBranchEntries, branchSelect.value);
    });

    viewEmployeeSummaryBtn.addEventListener('click', () => {
        if (!employeeSelect.value) {
            displayMessage('Please select an employee first.', 'error');
            return;
        }
        document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        viewEmployeeSummaryBtn.classList.add('active');
        // selectedEmployeeCodeEntries is already filtered by the employeeSelect change listener
        renderEmployeeSummary(selectedEmployeeCodeEntries);
    });

    viewAllEntriesBtn.addEventListener('click', () => {
        if (!employeeSelect.value) {
            displayMessage('Please select an employee first.', 'error');
            return;
        }
        document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        viewAllEntriesBtn.classList.add('active');
        // selectedEmployeeCodeEntries is already filtered by the employeeSelect change listener
        renderEmployeeDetailedEntries(selectedEmployeeCodeEntries);
    });

    viewPerformanceReportBtn.addEventListener('click', () => {
        if (!employeeSelect.value) {
            displayMessage('Please select an employee first.', 'error');
            return;
        }
        document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        viewPerformanceReportBtn.classList.add('active');
        // selectedEmployeeCodeEntries is already filtered by the employeeSelect change listener
        renderPerformanceReport(selectedEmployeeCodeEntries);
    });

    // Render Branch Performance Report (d2.PNG) - for a selected branch
    function renderPerformanceReportByBranch(entries, branchName) {
        reportDisplay.innerHTML = `<h2>Branch Performance Report - ${branchName} (This Month)</h2>`;

        if (entries.length === 0) {
            reportDisplay.innerHTML += '<p class="no-participation-message">No activity data found for this branch this month to generate a performance report.</p>';
            return;
        }

        const employeeActivityInBranch = {}; // {employeeCode: [entries]}

        entries.forEach(entry => {
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            if (!employeeActivityInBranch[employeeCode]) {
                employeeActivityInBranch[employeeCode] = [];
            }
            employeeActivityInBranch[employeeCode].push(entry);
        });

        const performanceGrid = document.createElement('div');
        performanceGrid.className = 'branch-performance-grid';

        for (const employeeCode in employeeActivityInBranch) {
            const employeeEntries = employeeActivityInBranch[employeeCode];
            const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
            const employeeDesignation = employeeCodeToDesignationMap[employeeCode] || 'Default';

            const { totalActivity } = calculateTotalActivity(employeeEntries);

            const metricCard = document.createElement('div');
            metricCard.className = 'employee-performance-card'; // Reusing this class for individual employee cards

            let cardContent = `<h4>${employeeName} (${employeeCode})</h4>
                               <p>Designation: <strong>${employeeDesignation}</strong></p>`;

            const metrics = ['Visit', 'Call', 'Reference', 'New Customer Leads'];

            metrics.forEach(metric => {
                const actual = totalActivity[metric];
                const target = TARGETS[employeeDesignation] ? TARGETS[employeeDesignation][metric] : TARGETS['Default'][metric];
                const performancePercentage = target > 0 ? (actual / target) * 100 : 0;

                cardContent += `<p>${metric}: <strong>${actual}</strong> / ${target} (${Math.round(performancePercentage)}%)</p>
                                <div class="performance-progress-bar-container">
                                    <div class="performance-progress-bar ${getProgressBarClass(performancePercentage)}" style="width: ${Math.min(100, performancePercentage)}%;"></div>
                                </div>`;
            });
            metricCard.innerHTML = cardContent;
            performanceGrid.appendChild(metricCard);
        }
        reportDisplay.appendChild(performanceGrid);
    }


    // NEW: Function to render the Daily Follow-up Report
    function renderDailyFollowUpReport() {
        reportDisplay.innerHTML = '<h2>Daily Follow-up Report</h2>';

        const today = new Date();
        // Set hours, minutes, seconds, milliseconds to 0 for accurate date comparison
        today.setHours(0, 0, 0, 0); 
        const todayIso = today.toISOString().split('T')[0]; // Get today's date in YYYY-MM-DD format

        const dueFollowUps = allCanvassingData.filter(entry => {
            const nextFollowUpDateString = entry[HEADER_NEXT_FOLLOW_UP_DATE];
            if (nextFollowUpDateString) {
                const followUpDate = new Date(nextFollowUpDateString);
                // Set hours, minutes, seconds, milliseconds to 0 for comparison
                followUpDate.setHours(0, 0, 0, 0);
                return followUpDate.toISOString().split('T')[0] === todayIso;
            }
            return false;
        });

        if (dueFollowUps.length === 0) {
            reportDisplay.innerHTML += '<p class="no-participation-message">No follow-ups due today!</p>';
            return;
        }

        reportDisplay.innerHTML += `<p><strong>Total Follow-ups Due Today:</strong> ${dueFollowUps.length}</p>`;

        // Optional: Breakdown by Employee and Branch
        const employeeBreakdown = {};
        const branchBreakdown = {};

        dueFollowUps.forEach(entry => {
            const employeeName = entry[HEADER_EMPLOYEE_NAME] || 'N/A';
            const branchName = entry[HEADER_BRANCH_NAME] || 'N/A';

            employeeBreakdown[employeeName] = (employeeBreakdown[employeeName] || 0) + 1;
            branchBreakdown[branchName] = (branchBreakdown[branchName] || 0) + 1;
        });

        const breakdownDiv = document.createElement('div');
        breakdownDiv.className = 'follow-up-breakdown';

        const employeeListDiv = document.createElement('div');
        employeeListDiv.innerHTML = '<h3>Due by Employee:</h3><ul class="summary-list">';
        for (const emp in employeeBreakdown) {
            employeeListDiv.innerHTML += `<li><strong>${emp}:</strong> ${employeeBreakdown[emp]}</li>`;
        }
        employeeListDiv.innerHTML += '</ul>';
        breakdownDiv.appendChild(employeeListDiv);

        const branchListDiv = document.createElement('div');
        branchListDiv.innerHTML = '<h3>Due by Branch:</h3><ul class="summary-list">';
        for (const branch in branchBreakdown) {
            branchListDiv.innerHTML += `<li><strong>${branch}:</strong> ${branchBreakdown[branch]}</li>`;
        }
        branchListDiv.innerHTML += '</ul>';
        breakdownDiv.appendChild(branchListDiv);

        reportDisplay.appendChild(breakdownDiv); // Add breakdown before the table

        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container';

        const table = document.createElement('table');
        table.className = 'data-table';
        const thead = table.createTHead();
        const headerRow = thead.insertRow();

        const headersToShow = [
            HEADER_NEXT_FOLLOW_UP_DATE,
            HEADER_PRODUCT_INTERESTED,
            HEADER_PROSPECT_NAME,
            HEADER_PHONE_NUMBER_WHATSAPP,
            HEADER_EMPLOYEE_NAME,
            HEADER_BRANCH_NAME
        ];

        headersToShow.forEach(headerText => {
            const th = document.createElement('th');
            th.textContent = headerText;
            headerRow.appendChild(th);
        });

        const tbody = table.createTBody();
        dueFollowUps.forEach(entry => {
            const row = tbody.insertRow();
            headersToShow.forEach(header => {
                const cell = row.insertCell();
                let value = entry[header] || '';

                if (header === HEADER_NEXT_FOLLOW_UP_DATE) {
                    value = formatDate(value); // Format the date
                }
                cell.textContent = value;
                cell.setAttribute('data-label', header);
            });
        });

        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }


    // Function to manage tab visibility and content rendering
    function showTab(tabButtonId) {
        // Deactivate all tab buttons
        document.querySelectorAll('.tab-button').forEach(button => {
            button.classList.remove('active');
        });
        // Activate the clicked tab button
        document.getElementById(tabButtonId).classList.add('active');

        // Hide all main content sections
        reportsSection.style.display = 'none';
        employeeManagementSection.style.display = 'none';

        // Hide view options and employee filter panel by default for new tab selection
        document.querySelector('.controls-panel').style.display = 'none';
        employeeFilterPanel.style.display = 'none';
        viewOptions.style.display = 'none';
        branchSelect.value = ''; // Clear branch selection
        employeeSelect.value = ''; // Clear employee selection
        reportDisplay.innerHTML = ''; // Clear report display

        // Show relevant section and render report based on tab
        if (tabButtonId === 'allBranchSnapshotTabBtn') {
            reportsSection.style.display = 'block';
            document.querySelector('.controls-panel').style.display = 'flex'; // Show controls
            renderAllBranchSnapshot();
        } else if (tabButtonId === 'allStaffOverallPerformanceTabBtn') {
            reportsSection.style.display = 'block';
            document.querySelector('.controls-panel').style.display = 'flex'; // Show controls
            renderOverallStaffPerformanceReport();
        } else if (tabButtonId === 'nonParticipatingBranchesTabBtn') {
            reportsSection.style.display = 'block';
            renderNonParticipatingBranches();
        } else if (tabButtonId === 'dailyFollowUpReportTabBtn') { // Handle the new tab
            reportsSection.style.display = 'block';
            renderDailyFollowUpReport();
        } else if (tabButtonId === 'employeeManagementTabBtn') {
            employeeManagementSection.style.display = 'block';
            displayEmployeeManagementMessage('', false); // Clear any old messages
        }
    }


    // Event listeners for main tab buttons
    if (allBranchSnapshotTabBtn) {
        allBranchSnapshotTabBtn.addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    }
    if (allStaffOverallPerformanceTabBtn) {
        allStaffOverallPerformanceTabBtn.addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    }
    if (nonParticipatingBranchesTabBtn) {
        nonParticipatingBranchesTabBtn.addEventListener('click', () => showTab('nonParticipatingBranchesTabBtn'));
    }
    if (dailyFollowUpReportTabBtn) { // Add event listener for the new button
        dailyFollowUpReportTabBtn.addEventListener('click', () => showTab('dailyFollowUpReportTabBtn'));
    }
    if (employeeManagementTabBtn) {
        employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));
    }


    // --- Employee Management Form Event Listeners ---
    // Event Listener for Add Employee Form
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
            }
        });
    }

    // Event Listener for Bulk Add Employee Form
    if (bulkAddEmployeeForm) {
        bulkAddEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const branchName = bulkEmployeeBranchNameInput.value.trim();
            const bulkDetails = bulkEmployeeDetailsTextarea.value.trim();

            if (!branchName || !bulkDetails) {
                displayEmployeeManagementMessage('Branch Name and Employee Details are required for bulk addition.', true);
                return;
            }

            const lines = bulkDetails.split('\n').map(line => line.trim()).filter(line => line.length > 0);
            const employeesToAdd = [];

            for (const line of lines) {
                const parts = line.split(',').map(p => p.trim());
                if (parts.length >= 2) { // Minimum Name, Code
                    const employeeData = {
                        [HEADER_EMPLOYEE_NAME]: parts[0],
                        [HEADER_EMPLOYEE_CODE]: parts[1],
                        [HEADER_BRANCH_NAME]: branchName,
                        [HEADER_DESIGNATION]: parts[2] || ''
                    };
                    employeesToAdd.push(employeeData);
                }
            }

            if (employeesToAdd.length > 0) {
                const success = await sendDataToGoogleAppsScript('add_bulk_employees', employeesToAdd);
                if (success) {
                    bulkAddEmployeeForm.reset();
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
            }
        });
    }

    // Initial data fetch and tab display when the page loads
    processData();
    showTab('allBranchSnapshotTabBtn'); // Keep default tab or change to your preference
});
