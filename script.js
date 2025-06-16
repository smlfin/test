
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
            'Visit': 30,
            'Call': 5 * MONTHLY_WORKING_DAYS,
            'Reference': 1 * MONTHLY_WORKING_DAYS,
            'New Customer Leads': 20
        },
        'Seniors': { // Added Investment Staff with custom Visit target
            'Visit': 30,
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
    const detailedCustomerViewTabBtn = document.getElementById('detailedCustomerViewTabBtn'); // NEW
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');

    // Main Content Sections to toggle
    const reportsSection = document.getElementById('reportsSection');
    const detailedCustomerViewSection = document.getElementById('detailedCustomerViewSection'); // NEW
    const employeeManagementSection = document.getElementById('employeeManagementSection');

    // NEW: Detailed Customer View Elements
    const customerViewBranchSelect = document.getElementById('customerViewBranchSelect');
    const customerViewEmployeeSelect = document.getElementById('customerViewEmployeeSelect');
    const customerCanvassedList = document.getElementById('customerCanvassedList');
    const customerDetailsContent = document.getElementById('customerDetailsContent'); // Parent of the cards

    // Get the specific card elements for direct population
    const customerCard1 = document.getElementById('customerCard1');
    const customerCard2 = document.getElementById('customerCard2');
    const customerCard3 = document.getElementById('customerCard3');


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
            // Use the name from the map if available, otherwise use the code for sorting and display
            const nameA = employeeCodeToNameMap[codeA] || codeA;
            const nameB = employeeCodeToNameMap[codeB] || codeB;
            return nameA.localeCompare(nameB);
        });

        populateDropdown(branchSelect, allUniqueBranches); // Populate branch dropdown with predefined branches
        populateDropdown(customerViewBranchSelect, allUniqueBranches); // Populate branch dropdown for detailed customer view
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


    // *** Report Rendering Functions ***

    function renderAllBranchSnapshot() {
        reportDisplay.innerHTML = '<h2>All Branch Snapshot</h2>';

        // Calculate total activities per branch
        const branchActivity = {};
        allCanvassingData.forEach(entry => {
            const branchName = entry[HEADER_BRANCH_NAME];
            const activityType = entry[HEADER_ACTIVITY_TYPE];
            if (branchName) {
                if (!branchActivity[branchName]) {
                    branchActivity[branchName] = { 'Visit': 0, 'Call': 0, 'Other': 0, 'Total': 0 };
                }
                const type = activityType ? activityType.trim() : 'Other';
                branchActivity[branchName][type]++;
                branchActivity[branchName]['Total']++;
            }
        });

        const table = document.createElement('table');
        table.className = 'all-branch-snapshot-table'; // Apply class for styling
        table.innerHTML = `
            <thead>
                <tr>
                    <th>Branch Name</th>
                    <th>Total Visits</th>
                    <th>Total Calls</th>
                    <th>Total Other Activities</th>
                    <th>Overall Total</th>
                </tr>
            </thead>
            <tbody>
            </tbody>
        `;
        const tbody = table.querySelector('tbody');

        PREDEFINED_BRANCHES.forEach(branch => { // Use PREDEFINED_BRANCHES for consistent display
            const data = branchActivity[branch] || { 'Visit': 0, 'Call': 0, 'Other': 0, 'Total': 0 };
            const row = tbody.insertRow();
            row.innerHTML = `
                <td data-label="Branch Name">${branch}</td>
                <td data-label="Total Visits">${data['Visit']}</td>
                <td data-label="Total Calls">${data['Call']}</td>
                <td data-label="Total Other Activities">${data['Other']}</td>
                <td data-label="Overall Total">${data['Total']}</td>
            `;
        });

        reportDisplay.appendChild(table);
    }


    function renderAllStaffOverallPerformance() {
        reportDisplay.innerHTML = '<h2>All Staff Overall Performance</h2>';

        // Group activities by employee code
        const employeeActivity = {};
        allCanvassingData.forEach(entry => {
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            const activityType = entry[HEADER_ACTIVITY_TYPE];
            const remarks = entry[HEADER_REMARKS];
            const customerName = entry[HEADER_CUSTOMER_NAME];
            const branchName = entry[HEADER_BRANCH_NAME];
            const date = entry[HEADER_DATE];

            if (employeeCode) {
                if (!employeeActivity[employeeCode]) {
                    employeeActivity[employeeCode] = {
                        'Visit': 0, 'Call': 0, 'Other': 0, 'Total': 0,
                        'name': entry[HEADER_EMPLOYEE_NAME] || employeeCode,
                        'designation': entry[HEADER_DESIGNATION] || 'N/A',
                        'branch': branchName || 'N/A',
                        'latestActivity': {}
                    };
                }
                const type = activityType ? activityType.trim() : 'Other';
                employeeActivity[employeeCode][type]++;
                employeeActivity[employeeCode]['Total']++;

                // Track latest activity for each employee
                const currentActivityDate = new Date(date);
                const latestDate = new Date(employeeActivity[employeeCode].latestActivity.date || 0);

                if (currentActivityDate > latestDate) {
                    employeeActivity[employeeCode].latestActivity = {
                        date: date,
                        type: type,
                        customer: customerName,
                        remarks: remarks
                    };
                }
            }
        });

        const table = document.createElement('table');
        table.className = 'all-staff-performance-table'; // Apply class for styling
        table.innerHTML = `
            <thead>
                <tr>
                    <th>Employee Name (Code)</th>
                    <th>Branch</th>
                    <th>Designation</th>
                    <th>Total Visits</th>
                    <th>Total Calls</th>
                    <th>Total Other Activities</th>
                    <th>Overall Total</th>
                    <th>Latest Activity Date</th>
                    <th>Latest Activity Type</th>
                    <th>Latest Customer</th>
                    <th>Latest Remarks</th>
                </tr>
            </thead>
            <tbody>
            </tbody>
        `;
        const tbody = table.querySelector('tbody');

        // Sort employees by overall total activity in descending order
        const sortedEmployees = Object.keys(employeeActivity).sort((a, b) => {
            return employeeActivity[b]['Total'] - employeeActivity[a]['Total'];
        });

        sortedEmployees.forEach(employeeCode => {
            const data = employeeActivity[employeeCode];
            const row = tbody.insertRow();
            row.innerHTML = `
                <td data-label="Employee Name (Code)">${data.name} (${employeeCode})</td>
                <td data-label="Branch">${data.branch}</td>
                <td data-label="Designation">${data.designation}</td>
                <td data-label="Total Visits">${data['Visit']}</td>
                <td data-label="Total Calls">${data['Call']}</td>
                <td data-label="Total Other Activities">${data['Other']}</td>
                <td data-label="Overall Total">${data['Total']}</td>
                <td data-label="Latest Activity Date">${data.latestActivity.date || 'N/A'}</td>
                <td data-label="Latest Activity Type">${data.latestActivity.type || 'N/A'}</td>
                <td data-label="Latest Customer">${data.latestActivity.customer || 'N/A'}</td>
                <td data-label="Latest Remarks">${data.latestActivity.remarks || 'N/A'}</td>
            `;
        });
        reportDisplay.appendChild(table);
    }

    function renderNonParticipatingBranches() {
        reportDisplay.innerHTML = '<h2>Non-Participating Branches (No Activity)</h2>';

        const activeBranches = new Set();
        allCanvassingData.forEach(entry => {
            if (entry[HEADER_BRANCH_NAME]) {
                activeBranches.add(entry[HEADER_BRANCH_NAME]);
            }
        });

        const nonParticipating = PREDEFINED_BRANCHES.filter(branch => !activeBranches.has(branch));

        if (nonParticipating.length > 0) {
            const list = document.createElement('ul');
            list.className = 'non-participating-list';
            nonParticipating.forEach(branch => {
                const li = document.createElement('li');
                li.textContent = branch;
                list.appendChild(li);
            });
            reportDisplay.appendChild(list);
        } else {
            reportDisplay.innerHTML += '<p class="success-message">All predefined branches have recorded activity!</p>';
        }
    }

    function renderCustomerCanvassedList(branchFilter = '', employeeFilter = '') {
        customerCanvassedList.innerHTML = ''; // Clear previous list

        let filteredData = allCanvassingData;

        if (branchFilter) {
            filteredData = filteredData.filter(row => row[HEADER_BRANCH_NAME] === branchFilter);
        }

        if (employeeFilter) {
            filteredData = filteredData.filter(row => row[HEADER_EMPLOYEE_CODE] === employeeFilter);
        }

        if (filteredData.length === 0) {
            customerCanvassedList.innerHTML = '<p class="info-message">No customer canvassed data available for the selected filters.</p>';
            return;
        }

        const ul = document.createElement('ul');
        ul.className = 'customer-list';

        // Sort by Date (most recent first)
        filteredData.sort((a, b) => new Date(b[HEADER_DATE]) - new Date(a[HEADER_DATE]));

        filteredData.forEach(customer => {
            const li = document.createElement('li');
            li.className = 'customer-list-item';
            li.innerHTML = `
                <div class="customer-name">${customer[HEADER_CUSTOMER_NAME] || 'N/A'}</div>
                <div class="customer-meta">
                    <span>${customer[HEADER_BRANCH_NAME] || 'N/A'}</span> -
                    <span>${customer[HEADER_EMPLOYEE_NAME] || 'N/A'}</span> -
                    <span>${customer[HEADER_DATE] || 'N/A'}</span>
                </div>
            `;
            li.addEventListener('click', () => displayCustomerDetails(customer));
            ul.appendChild(li);
        });
        customerCanvassedList.appendChild(ul);
    }

    function displayCustomerDetails(customer) {
        customerDetailsContent.innerHTML = `
            <h3>Customer Details: ${customer[HEADER_CUSTOMER_NAME] || 'N/A'}</h3>
            <div class="detail-grid">
                <div class="detail-row"><span class="detail-label">Customer Contact:</span> <span class="detail-value">${customer[HEADER_CUSTOMER_CONTACT] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Address:</span> <span class="detail-value">${customer[HEADER_ADDRESS] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Branch:</span> <span class="detail-value">${customer[HEADER_BRANCH_NAME] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Employee:</span> <span class="detail-value">${customer[HEADER_EMPLOYEE_NAME] || 'N/A'} (${customer[HEADER_EMPLOYEE_CODE] || 'N/A'})</span></div>
                <div class="detail-row"><span class="detail-label">Activity Type:</span> <span class="detail-value">${customer[HEADER_ACTIVITY_TYPE] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Date:</span> <span class="detail-value">${customer[HEADER_DATE] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Time:</span> <span class="detail-value">${customer[HEADER_TIMESTAMP] ? new Date(customer[HEADER_TIMESTAMP]).toLocaleTimeString() : 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Remarks:</span> <span class="detail-value">${customer[HEADER_REMARKS] || 'N/A'}</span></div>
                <div class="detail-row"><span class="detail-label">Status:</span> <span class="detail-value">${customer[HEADER_STATUS] || 'N/A'}</span></div>
            </div>
        `;
    }

    // Helper to calculate activity counts per branch
    function getBranchActivitySummary() {
        const branchSummary = {};

        allCanvassingData.forEach(entry => {
            const branchName = entry[HEADER_BRANCH_NAME];
            const activityType = entry[HEADER_ACTIVITY_TYPE];

            if (!branchName) return; // Skip entries without a branch name

            if (!branchSummary[branchName]) {
                branchSummary[branchName] = { 'Visit': 0, 'Call': 0 };
            }

            const trimmedActivityType = activityType ? activityType.trim().toLowerCase() : '';

            if (trimmedActivityType === 'visit') {
                branchSummary[branchName]['Visit']++;
            } else if (trimmedActivityType === 'calls') {
                branchSummary[branchName]['Call']++;
            }
        });

        return branchSummary;
    }

    // Render Branch Performance Reports (Max/Min Visits/Calls)
    function renderBranchPerformanceReports() {
        reportDisplay.innerHTML = '<h2>Branch Performance Overview</h2>';
        const branchActivity = getBranchActivitySummary();

        let maxVisitsBranch = { name: 'N/A', count: -1 };
        let minVisitsBranch = { name: 'N/A', count: Infinity };
        let maxCallsBranch = { name: 'N/A', count: -1 };
        let minCallsBranch = { name: 'N/A', count: Infinity };
        let hasData = false;

        for (const branchName in branchActivity) {
            hasData = true;
            const visits = branchActivity[branchName]['Visit'];
            const calls = branchActivity[branchName]['Call'];

            if (visits > maxVisitsBranch.count) {
                maxVisitsBranch = { name: branchName, count: visits };
            }
            if (visits < minVisitsBranch.count) {
                minVisitsBranch = { name: branchName, count: visits };
            }
            if (calls > maxCallsBranch.count) {
                maxCallsBranch = { name: branchName, count: calls };
            }
            if (calls < minCallsBranch.count) {
                minCallsBranch = { name: branchName, count: calls };
            }
        }

        if (!hasData) {
            reportDisplay.innerHTML += '<p class="info-message">No activity data available to generate branch performance reports.</p>';
            return;
        }

        const reportContent = `
            <div class="branch-performance-grid">
                <div class="performance-card">
                    <h3>Maximum Visits</h3>
                    <p><strong>${maxVisitsBranch.name}</strong> with ${maxVisitsBranch.count} visits</p>
                </div>
                <div class="performance-card">
                    <h3>Minimum Visits</h3>
                    <p><strong>${minVisitsBranch.name}</strong> with ${minVisitsBranch.count} visits</p>
                </div>
                <div class="performance-card">
                    <h3>Maximum Calls</h3>
                    <p><strong>${maxCallsBranch.name}</strong> with ${maxCallsBranch.count} calls</p>
                </div>
                <div class="performance-card">
                    <h3>Minimum Calls</h3>
                    <p><strong>${minCallsBranch.name}</strong> with ${minCallsBranch.count} calls</p>
                </div>
            </div>
            <h3 style="margin-top: 30px;">Branch-wise Activity Breakdown</h3>
            <table class="branch-summary-table">
                <thead>
                    <tr>
                        <th>Branch Name</th>
                        <th>Total Visits</th>
                        <th>Total Calls</th>
                    </tr>
                </thead>
                <tbody>
                    ${PREDEFINED_BRANCHES.map(branch => `
                        <tr>
                            <td>${branch}</td>
                            <td>${branchActivity[branch] ? branchActivity[branch]['Visit'] : 0}</td>
                            <td>${branchActivity[branch] ? branchActivity[branch]['Call'] : 0}</td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        `;
        reportDisplay.innerHTML += reportContent;
    }


    // *** Event Listeners ***

    // Tab switching logic
    function showTab(tabId) {
        // Deactivate all buttons
        document.querySelectorAll('.tab-button').forEach(button => {
            button.classList.remove('active');
        });

        // Hide all sections
        reportsSection.style.display = 'none';
        detailedCustomerViewSection.style.display = 'none';
        employeeManagementSection.style.display = 'none';

        // Activate selected button and show relevant section
        document.getElementById(tabId).classList.add('active');

        if (tabId === 'allBranchSnapshotTabBtn') {
            reportsSection.style.display = 'block';
            renderAllBranchSnapshot();
        } else if (tabId === 'allStaffOverallPerformanceTabBtn') {
            reportsSection.style.display = 'block';
            renderAllStaffOverallPerformance();
        } else if (tabId === 'nonParticipatingBranchesTabBtn') {
            reportsSection.style.display = 'block';
            renderNonParticipatingBranches();
        } else if (tabId === 'detailedCustomerViewTabBtn') {
            detailedCustomerViewSection.style.display = 'block';
            renderCustomerCanvassedList();
            customerDetailsContent.innerHTML = '<p class="info-message">Select a customer from the list to view details.</p>';
        } else if (tabId === 'employeeManagementTabBtn') {
            employeeManagementSection.style.display = 'block';
            displayEmployeeManagementMessage('', false);
        } else if (tabId === 'branchPerformanceTabBtn') { // NEW TAB CASE
            reportsSection.style.display = 'block';
            renderBranchPerformanceReports(); // Call the new function
        }
    }

    if (allBranchSnapshotTabBtn) {
        allBranchSnapshotTabBtn.addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    }
    if (allStaffOverallPerformanceTabBtn) {
        allStaffOverallPerformanceTabBtn.addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    }
    if (nonParticipatingBranchesTabBtn) {
        nonParticipatingBranchesTabBtn.addEventListener('click', () => showTab('nonParticipatingBranchesTabBtn'));
    }
    if (detailedCustomerViewTabBtn) {
        detailedCustomerViewTabBtn.addEventListener('click', () => showTab('detailedCustomerViewTabBtn'));
    }
    if (employeeManagementTabBtn) {
        employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));
    }
    // NEW Event Listener for Branch Performance Reports Tab
    if (branchPerformanceTabBtn) {
        branchPerformanceTabBtn.addEventListener('click', () => showTab('branchPerformanceTabBtn'));
    }


    if (branchSelect) {
        branchSelect.addEventListener('change', () => {
            const selectedBranch = branchSelect.value;
            // Only populate employee dropdown if detailed view is active
            if (detailedCustomerViewSection.style.display === 'block') {
                populateEmployeeDropdown(selectedBranch);
            }
            // Always re-render customer list based on current filters
            const selectedEmployee = employeeSelect ? employeeSelect.value : '';
            renderCustomerCanvassedList(selectedBranch, selectedEmployee);
        });
    }

    if (employeeSelect) {
        employeeSelect.addEventListener('change', () => {
            const selectedBranch = branchSelect ? branchSelect.value : '';
            const selectedEmployee = employeeSelect.value;
            renderCustomerCanvassedList(selectedBranch, selectedEmployee);
        });
    }

    // Event Listener for Add Employee Form
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const employeeName = employeeNameInput.value.trim();
            const employeeCode = employeeCodeInput.value.trim();
            const designation = designationInput.value.trim();
            const branchName = branchNameInput.value.trim();

            if (!employeeName || !employeeCode || !branchName) {
                displayEmployeeManagementMessage('Employee Name, Code, and Branch Name are required.', true);
                return;
            }

            const employeeData = {
                [HEADER_EMPLOYEE_NAME]: employeeName,
                [HEADER_EMPLOYEE_CODE]: employeeCode,
                [HEADER_DESIGNATION]: designation,
                [HEADER_BRANCH_NAME]: branchName
            };

            const success = await sendDataToGoogleAppsScript('add_employee', employeeData);

            if (success) {
                addEmployeeForm.reset();
            }
        });
    }

    // Event Listener for Update Employee Form
    if (updateEmployeeForm) {
        updateEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const employeeCode = updateEmployeeCodeInput.value.trim();
            const newEmployeeName = updateEmployeeNameInput.value.trim();
            const newDesignation = updateDesignationInput.value.trim();
            const newBranchName = updateBranchNameInput.value.trim();

            if (!employeeCode) {
                displayEmployeeManagementMessage('Employee Code is required to update.', true);
                return;
            }

            const updateData = {
                [HEADER_EMPLOYEE_CODE]: employeeCode
            };
            if (newEmployeeName) updateData[HEADER_EMPLOYEE_NAME] = newEmployeeName;
            if (newDesignation) updateData[HEADER_DESIGNATION] = newDesignation;
            if (newBranchName) updateData[HEADER_BRANCH_NAME] = newBranchName;


            if (Object.keys(updateData).length === 1) { // Only employeeCode, no other updates
                displayEmployeeManagementMessage('Please provide at least one field (Name, Designation, or Branch) to update.', true);
                return;
            }

            const success = await sendDataToGoogleAppsScript('update_employee', updateData);

            if (success) {
                updateEmployeeForm.reset();
            }
        });
    }

    // Event Listener for Bulk Add Employee Form
    if (bulkAddEmployeeForm) {
        bulkAddEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const branchName = bulkEmployeeBranchNameInput.value.trim();
            const employeeDetails = bulkEmployeeDetailsTextarea.value.trim();

            if (!branchName || !employeeDetails) {
                displayEmployeeManagementMessage('Branch Name and Employee Details are required for bulk addition.', true);
                return;
            }

            const lines = employeeDetails.split('\n');
            const employeesToAdd = [];

            for (const line of lines) {
                const parts = line.split(',').map(p => p.trim());
                if (parts.length >= 2 && parts[0] && parts[1]) { // Name and Code are mandatory
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
    showTab('allBranchSnapshotTabBtn');
});
