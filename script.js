document.addEventListener('DOMContentLoaded', () => {

    // --- START: TWO-TIERED FRONT-END PASSWORD PROTECTION ---
    const ACCESS_PASSWORD_FULL = "sml4576"; // Full access password
    const ACCESS_PASSWORD_LIMITED = "123";  // Limited access password
    let currentAccessLevel = null; // To store 'full' or 'limited'
    const accessDeniedOverlay = document.getElementById('accessDeniedOverlay');
    const dashboardContent = document.getElementById('dashboardContent');
    const secretPasswordInputContainer = document.getElementById('secretPasswordInputContainer');
    const secretPasswordInput = document.getElementById('secretPasswordInput');
    const submitSecretPasswordBtn = document.getElementById('submitSecretPassword');
    const passwordErrorMessage = document.getElementById('passwordErrorMessage');
   // Get references to buttons/tabs that need conditional access
    const downloadOverallStaffPerformanceReportBtn = document.getElementById('downloadOverallStaffPerformanceReportBtn');
    const detailedCustomerViewTabBtn = document.getElementById('detailedCustomerViewTabBtn');
    const viewAllEntriesButton = document.getElementById('viewAllEntriesBtn'); // <--- ADD THIS LINE (OR UPDATE THE EXISTING PLACEHOLDER)
       if (secretPasswordInput) {
        secretPasswordInput.focus();
    }
    submitSecretPasswordBtn.addEventListener('click', () => {
        checkAndSetAccess();
    });
    secretPasswordInput.addEventListener('keypress', (event) => {
        if (event.key === 'Enter') {
            checkAndSetAccess();
        }
    });

    function checkAndSetAccess() {
        const enteredPassword = secretPasswordInput.value;

        if (enteredPassword === ACCESS_PASSWORD_FULL) {
            currentAccessLevel = 'full';
            grantAccess();
        } else if (enteredPassword === ACCESS_PASSWORD_LIMITED) {
            currentAccessLevel = 'limited';
            grantAccess();
        } else {
            passwordErrorMessage.textContent = "Incorrect password. Try again.";
            passwordErrorMessage.style.display = 'block';
            secretPasswordInput.value = '';
            secretPasswordInput.focus();
        }
    }

   function grantAccess() {
    console.log("grantAccess function called!"); // Add this line
    accessDeniedOverlay.style.display = 'none'; // Hide overlay
    console.log("Access overlay hidden."); // Add this line

    console.log("Value of dashboardContent:", dashboardContent); // Add this line
    if (dashboardContent) { // Add this if statement
        dashboardContent.style.display = 'block';   // Show dashboard content
        console.log("Dashboard content display set to block."); // Add this line
    } else {
        console.error("Error: dashboardContent element not found!"); // Add this line
    }

    // Apply access restrictions based on currentAccessLevel
    applyAccessRestrictions();
    // Now, call your main processing function that starts everything
    processData(); // This fetches your data
    // Set initial tab based on access level
    if (currentAccessLevel === 'limited') {
        showTab('allBranchSnapshotTabBtn');
    } else {
        // Full access users get the default initial tab
        showTab('allBranchSnapshotTabBtn');
    }
}

    function applyAccessRestrictions() {
        if (currentAccessLevel === 'limited') {
            // Hide "Download Overall Performance" button
            if (downloadOverallStaffPerformanceReportBtn) {
                downloadOverallStaffPerformanceReportBtn.style.display = 'none';
            }
            // Hide "Detailed Customer View" tab
            if (detailedCustomerViewTabBtn) {
                detailedCustomerViewTabBtn.style.display = 'none';
            }
            // Hide "View all entries" button
            if (viewAllEntriesButton) { // <--- ADD THIS BLOCK
                viewAllEntriesButton.style.display = 'none';
            }
        } else if (currentAccessLevel === 'full') {
            // Ensure all are visible for full access
            if (downloadOverallStaffPerformanceReportBtn) {
                downloadOverallStaffPerformanceReportBtn.style.display = 'inline-block'; // Or 'block', depending on your CSS
            }
            if (detailedCustomerViewTabBtn) {
                detailedCustomerViewTabBtn.style.display = 'inline-block'; // Or 'block'
            }
            if (viewAllEntriesButton) { // <--- ADD THIS BLOCK
                viewAllEntriesButton.style.display = 'inline-block'; // Or 'block'
            }
        }
    }
   
// This URL is for your Canvassing Data sheet. Ensure it's correct and published as CSV.
const DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=696550092&single=true&output=csv"; 
// IMPORTANT: Replace this with YOUR DEPLOYED GOOGLE APPS SCRIPT WEB APP URL
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxe_hZyRXZdY1CbfchvH_pzIa596dxmEDnPVc4YGXWerxRmuJz30CpEbND279mR0lWf/exec"; // <-- PASTE YOUR NEWLY DEPLOYED WEB APP URL HERE
// For front-end reporting, all employee and branch data will come from Canvassing Data and predefined list.
   // const EMPLOYEE_MASTER_DATA_URL = "UNUSED"; // Marked as UNUSED for clarity, won't be fetched for reports

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
        "Muvattupuzha", "Thiruvalla", "Pathanamthitta", "Kunnamkulam", "HO KKM" // Corrected "Pathanamthitta" typo if it existed previously
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
    const HEADER_PHONE_NUMBER_WHATSAPP = 'Phone Number(Whatsapp)'; // CORRECTED TYPO HERE
    const HEADER_ADDRESS = 'Address';
    const HEADER_PROFESSION = 'Profession';
    const HEADER_DOB_WD = 'DOB/WD';
    const HEADER_PRODUCT_INTERESTED = 'Product Interested'; // CORRECTED TYPO HERE
    const HEADER_REMARKS = 'Remarks';
    const HEADER_NEXT_FOLLOW_UP_DATE = 'Next Follow-up Date';
    const HEADER_RELATION_WITH_STAFF = 'Relation With Staff';
    // NEW: Customer Detail Headers as provided by user
    const HEADER_FAMILY_DETAILS_1 = 'Family Deatils -1 Name of wife/Husband';
    const HEADER_FAMILY_DETAILS_2 = 'Family Deatils -2 Job of wife/Husband';
    const HEADER_FAMILY_DETAILS_3 = 'Family Deatils -3 Names of Children';
    const HEADER_FAMILY_DETAILS_4 = 'Family Deatils -4 Deatils of Children';
    const HEADER_PROFILE_OF_CUSTOMER = 'Profile of Customer';


    // Core Display and Status Elements
    const reportDisplay = document.getElementById('reportDisplay');
    const statusMessage = document.getElementById('statusMessage');

    // Main Content Sections to toggle
    const reportsSection = document.getElementById('reportsSection');
    const detailedCustomerViewSection = document.getElementById('detailedCustomerViewSection');
    const employeeManagementSection = document.getElementById('employeeManagementSection');

    // Tab buttons for main navigation
    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');
    const nonParticipatingBranchesTabBtn = document.getElementById('nonParticipatingBranchesTabBtn');
    const branchPerformanceTabBtn = document.getElementById('branchPerformanceTabBtn');
    // Removed the problematic duplicate declaration: const = document.getElementById('detailedCustomerViewTabBtn');
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');

    // Dropdowns & Filter Panels
    const branchSelect = document.getElementById('branchSelect');
    const employeeSelect = document.getElementById('employeeSelect');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel'); // New from your list
    const monthSelect = document.getElementById('monthSelect'); // NEW: Month Select
    const customerViewBranchSelect = document.getElementById('customerViewBranchSelect');
    const customerViewEmployeeSelect = document.getElementById('customerViewEmployeeSelect');
    const customerViewMonthSelect = document.getElementById('customerViewMonthSelect'); // NEW: Month Select for Customer View

    // View Option Buttons (from your provided list)
    const viewOptions = document.getElementById('viewOptions');
    const viewBranchPerformanceReportBtn = document.getElementById('viewBranchPerformanceReportBtn');
    const viewEmployeeSummaryBtn = document.getElementById('viewEmployeeSummaryBtn');
    const viewAllEntriesBtn = document.getElementById('viewAllEntriesBtn');
    const viewPerformanceReportBtn = document.getElementById('viewPerformanceReportBtn');
    const viewBranchVisitLeaderboardBtn = document.getElementById('viewBranchVisitLeaderboardBtn');
    const viewBranchCallLeaderboardBtn = document.getElementById('viewBranchCallLeaderboardBtn');
    const viewStaffParticipationBtn = document.getElementById('viewStaffParticipationBtn');

    // Detailed Customer View Specific Elements
    const customerCanvassedList = document.getElementById('customerCanvassedList');
    const customerDetailsContent = document.getElementById('customerDetailsContent');
    const customerCard1 = document.getElementById('customerCard1');
    const customerCard2 = document.getElementById('customerCard2');
    const customerCard3 = document.getElementById('customerCard3');
    const detailedCustomerReportTableBody = document.getElementById('detailedCustomerReportTableBody');

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

    // Download Buttons
    const downloadDetailedCustomerReportBtn = document.getElementById('downloadDetailedCustomerReportBtn');
    // const downloadOverallStaffPerformanceReportBtn = document.getElementById('downloadOverallStaffPerformanceReportBtn'); 
    
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

    // Helper function to standardize various timestamp formats into a parsable format
    function standardizeTimestampFormat(timestampStr) {
        if (!timestampStr) return null;

        // Attempt to parse as 'M/D/YYYY H:MM:SS AM/PM' (e.g., "7/1/2025 9:00:00 AM")
        // This is a common format from Google Sheets exports.
        let date = new Date(timestampStr);

        if (!isNaN(date.getTime())) {
            // If successfully parsed, return a standardized string or the Date object directly
            // For Date constructor, 'YYYY-MM-DDTHH:mm:ss' is generally reliable.
            // Example: "2025-07-01T09:00:00"
            return date.toISOString().slice(0, 19);
        }

        // Add more parsing logic for other potential formats if needed in the future.
        // For now, if the default Date constructor fails, return null or original string
        console.warn(`Could not parse timestamp: "${timestampStr}". Returning original string.`);
        return timestampStr; // Return original if cannot parse, for debugging
    }

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
        
        // NEW: Populate month dropdowns
        populateMonthDropdowns();

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

    // NEW: Function to populate month dropdowns
    function populateMonthDropdowns() {
        const months = [
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        ];
        const currentMonthIndex = new Date().getMonth(); // 0-indexed
        const currentYear = new Date().getFullYear();

        [monthSelect, customerViewMonthSelect].forEach(selectElement => {
            selectElement.innerHTML = ''; // Clear existing options
            
            // Add a default "Select Month" option
            const defaultOption = document.createElement('option');
            defaultOption.value = '';
            defaultOption.textContent = '-- Select Month --';
            selectElement.appendChild(defaultOption);

            months.forEach((month, index) => {
                const option = document.createElement('option');
                // Store value as "MM-YYYY" or just month index for filtering
                option.value = `${index + 1}-${currentYear}`; // e.g., "7-2025" for July 2025
                option.textContent = `${month} ${currentYear}`;
                selectElement.appendChild(option);
            });
            // Select the current month
            selectElement.value = `${currentMonthIndex + 1}-${currentYear}`;
        });
    }

    // NEW: Function to filter data by selected month and year
    function filterDataByMonth(data, selectedMonthValue) {
        if (!selectedMonthValue) {
            return data; // If no month is selected, return all data
        }
        const [monthIndexStr, yearStr] = selectedMonthValue.split('-');
        const selectedMonth = parseInt(monthIndexStr) - 1; // Convert to 0-indexed month
        const selectedYear = parseInt(yearStr);

        return data.filter(entry => {
            const standardizedTimestamp = standardizeTimestampFormat(entry[HEADER_TIMESTAMP]);
            // Ensure standardizedTimestamp is not null before creating a Date object
            if (!standardizedTimestamp) {
                console.warn(`Skipping entry due to unparsable timestamp: ${entry[HEADER_TIMESTAMP]}`);
                return false; // Skip this entry if timestamp is invalid
            }
            const entryDate = new Date(standardizedTimestamp);
            // Check if entryDate is a valid Date object
            if (isNaN(entryDate.getTime())) {
                console.warn(`Invalid Date object created from standardized timestamp: ${standardizedTimestamp}. Original: ${entry[HEADER_TIMESTAMP]}`);
                return false; // Skip this entry if date is invalid after standardization
            }
            return entryDate.getMonth() === selectedMonth && entryDate.getFullYear() === selectedYear;
        });
    }

    // Filter employees based on selected branch
    branchSelect.addEventListener('change', () => {
        const selectedBranch = branchSelect.value;
        const selectedMonthValue = monthSelect.value; // Get selected month value

        // Filter allCanvassingData by selected month first
        let filteredByMonthData = filterDataByMonth(allCanvassingData, selectedMonthValue);

        if (selectedBranch) {
            employeeFilterPanel.style.display = 'block';

            // Get employee codes ONLY from filteredByMonthData for the selected branch
            const employeeCodesInBranchFromCanvassing = filteredByMonthData
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
        const selectedMonthValue = monthSelect.value; // Get selected month value

        // Filter allCanvassingData by selected month first
        let filteredByMonthData = filterDataByMonth(allCanvassingData, selectedMonthValue);

        if (selectedEmployeeCode) {
            // Filter activity data by employee code (from filteredByMonthData)
            selectedEmployeeCodeEntries = filteredByMonthData.filter(entry =>
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

    // NEW: Event listener for monthSelect dropdown (main reports section)
    monthSelect.addEventListener('change', () => {
        // Re-render the active report or default report based on month change
        const currentActiveTab = document.querySelector('.tab-button.active');
        if (currentActiveTab) {
            if (currentActiveTab.id === 'allBranchSnapshotTabBtn') {
                renderAllBranchSnapshot();
            } else if (currentActiveTab.id === 'allStaffOverallPerformanceTabBtn') {
                renderOverallStaffPerformanceReport();
            } else if (currentActiveTab.id === 'nonParticipatingBranchesTabBtn') {
                renderNonParticipatingBranches();
            } else if (currentActiveTab.id === 'branchPerformanceTabBtn') {
                const selectedBranch = branchSelect.value;
                if (selectedBranch) {
                    renderBranchPerformanceReport(selectedBranch);
                } else {
                    reportDisplay.innerHTML = '<p>Please select a branch and a month to view the Branch Performance Report.</p>';
                }
            } else if (currentActiveTab.id === 'performanceSummaryTabBtn') {
                const selectedBranch = branchSelect.value;
                const selectedEmployee = employeeSelect.value;
                if (selectedBranch && selectedEmployee) {
                    const filteredEntries = filterDataByMonth(allCanvassingData, monthSelect.value).filter(entry =>
                        entry[HEADER_EMPLOYEE_CODE] === selectedEmployee &&
                        entry[HEADER_BRANCH_NAME] === selectedBranch
                    );
                    renderEmployeeSummary(filteredEntries);
                } else {
                    reportDisplay.innerHTML = '<p>Please select a branch, an employee, and a month to view the Performance Summary.</p>';
                }
            }
        } else {
            // Default to All Branch Snapshot if no tab is active (shouldn't happen with initial load)
            renderAllBranchSnapshot();
        }
    });


    // NEW: Event listener for customerViewMonthSelect dropdown
    customerViewMonthSelect.addEventListener('change', () => {
        const selectedBranch = customerViewBranchSelect.value;
        const selectedEmployee = customerViewEmployeeSelect.value;
        if (selectedBranch && selectedEmployee) {
            loadDetailedCustomerReport(); // Reload the customer report with new month filter
        } else {
            detailedCustomerReportTableBody.innerHTML = '<tr><td colspan="5">Select a branch, employee, and month to load customer data.</td></tr>';
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

        // Get selected month value
        const selectedMonthValue = monthSelect.value;
        const filteredData = filterDataByMonth(allCanvassingData, selectedMonthValue);

        PREDEFINED_BRANCHES.forEach(branch => {
            const branchActivityEntries = filteredData.filter(entry => entry[HEADER_BRANCH_NAME] === branch);
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

        // Get selected month value
        const selectedMonthValue = monthSelect.value;
        const filteredData = filterDataByMonth(allCanvassingData, selectedMonthValue);

        PREDEFINED_BRANCHES.forEach(branch => {
            const branchActivityEntries = filteredData.filter(entry => entry[HEADER_BRANCH_NAME] === branch);
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
            reportDisplay.innerHTML += '<p class="no-participation-message">All predefined branches have recorded visits for the selected month!</p>'; // Changed message
        }
    }

    // Render All Staff Overall Performance Report (for d1.PNG)
    function renderOverallStaffPerformanceReport() {
        reportDisplay.innerHTML = '<h2>Overall Staff Performance Report</h2>';
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
        const selectedMonthValue = monthSelect.value; // Get selected month value
        const filteredData = filterDataByMonth(allCanvassingData, selectedMonthValue);

        // Get unique employees who have made at least one entry in the selected month
        const employeesWithActivityThisMonth = [...new Set(filteredData
            .map(entry => entry[HEADER_EMPLOYEE_CODE]))].sort((codeA, codeB) => {
                const nameA = employeeCodeToNameMap[codeA] || codeA;
                const nameB = employeeCodeToNameMap[codeB] || codeB;
                return nameA.localeCompare(nameB);
            });


        if (employeesWithActivityThisMonth.length === 0) {
            reportDisplay.innerHTML += '<p>No employee activity found for the selected month.</p>';
            return;
        }

        employeesWithActivityThisMonth.forEach(employeeCode => {
            const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
            const branchName = filteredData.find(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode)?.[HEADER_BRANCH_NAME] || 'N/A';
            const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';

            const employeeActivities = filteredData.filter(entry =>
                entry[HEADER_EMPLOYEE_CODE] === employeeCode
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

                if (isNaN(percentValue) || targetValue === 0) { // If target is 0,...
                    displayPercent = 'N/A';
                    progressWidth = 0;
                    progressBarClass = 'no-activity';
                } else {
                    displayPercent = `${Math.round(percentValue)}%`;
                    progressWidth = Math.min(100, Math.round(percentValue));
                    progressBarClass = getProgressBarClass(percentValue);
                }
                // Special handling for 0 actuals with positive targets to show 0% and danger color
                if (actualValue === 0 && targetValue > 0) {
                    displayPercent = '0%';
                    progressWidth = 0;
                    progressBarClass = 'danger';
                }
                row.insertCell().textContent = actualValue;
                row.insertCell().textContent = targetValue;
                const percentCell = row.insertCell();
                percentCell.innerHTML = `
                    <div class="progress-bar-container-small">
                        <div class="progress-bar ${progressBarClass}" style="width: ${progressWidth === 0 && displayPercent !== 'N/A' ? '30px' : progressWidth}%">
                            ${displayPercent}
                        </div>
                    </div>
                `;
            });
        });
        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }
    // Helper to calculate performance percentage
    function calculatePerformance(actuals, targets) {
        const performance = {};
        for (const metric in targets) {
            const actual = actuals[metric] || 0;
            const target = targets[metric];
            if (target > 0) {
                performance[metric] = (actual / target) * 100;
            } else {
                performance[metric] = NaN; // Or 0, depending on how you want to represent it
            }
        }
        return performance;
    }

    // Helper to get progress bar class based on percentage
    function getProgressBarClass(percentage) {
        if (percentage >= 100) {
            return 'excellent';
        } else if (percentage >= 75) {
            return 'good';
        } else if (percentage >= 50) {
            return 'average';
        } else {
            return 'danger';
        }
    }

    // Render Branch Performance Report (for d2.PNG)
    function renderBranchPerformanceReport(branchName) {
        reportDisplay.innerHTML = `<h2>Performance Report for ${branchName}</h2>`;
        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container'; // For horizontal scrolling

        const table = document.createElement('table');
        table.className = 'performance-table';
        
        const thead = table.createTHead();
        let headerRow = thead.insertRow();
        headerRow.insertCell().textContent = 'Employee Name';
        headerRow.insertCell().textContent = 'Designation';
        
        const metrics = ['Visit', 'Call', 'Reference', 'New Customer Leads'];
        metrics.forEach(metric => {
            const th = document.createElement('th');
            th.colSpan = 3; // Actual, Target, %
            th.textContent = metric;
            headerRow.appendChild(th);
        });

        headerRow = thead.insertRow(); // New row for sub-headers
        headerRow.insertCell(); // Empty for Employee Name
        headerRow.insertCell(); // Empty for Designation
        metrics.forEach(() => {
            ['Act', 'Tgt', '%'].forEach(subHeader => {
                const th = document.createElement('th');
                th.textContent = subHeader;
                headerRow.appendChild(th);
            });
        });

        const tbody = table.createTBody();
        const selectedMonthValue = monthSelect.value; // Get selected month value
        const filteredData = filterDataByMonth(allCanvassingData, selectedMonthValue);

        const employeesInBranch = [...new Set(filteredData
            .filter(entry => entry[HEADER_BRANCH_NAME] === branchName)
            .map(entry => entry[HEADER_EMPLOYEE_CODE]))].sort((codeA, codeB) => {
                const nameA = employeeCodeToNameMap[codeA] || codeA;
                const nameB = employeeCodeToNameMap[codeB] || codeB;
                return nameA.localeCompare(nameB);
            });


        if (employeesInBranch.length === 0) {
            reportDisplay.innerHTML += `<p>No employee activity found for ${branchName} in the selected month.</p>`;
            return;
        }

        employeesInBranch.forEach(employeeCode => {
            const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
            const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';

            const employeeActivities = filteredData.filter(entry =>
                entry[HEADER_EMPLOYEE_CODE] === employeeCode && entry[HEADER_BRANCH_NAME] === branchName
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
                 if (actualValue === 0 && targetValue > 0) {
                    displayPercent = '0%';
                    progressWidth = 0;
                    progressBarClass = 'danger';
                }
                row.insertCell().textContent = actualValue;
                row.insertCell().textContent = targetValue;
                const percentCell = row.insertCell();
                percentCell.innerHTML = `
                    <div class="progress-bar-container-small">
                        <div class="progress-bar ${progressBarClass}" style="width: ${progressWidth === 0 && displayPercent !== 'N/A' ? '30px' : progressWidth}%">
                            ${displayPercent}
                        </div>
                    </div>
                `;
            });
        });
        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }

    // Render Employee Summary (for d4.PNG)
    function renderEmployeeSummary(employeeEntries) {
        if (employeeEntries.length === 0) {
            reportDisplay.innerHTML = '<p>No activity found for this employee in the selected month.</p>';
            return;
        }

        const employeeCode = employeeEntries[0][HEADER_EMPLOYEE_CODE];
        const employeeName = employeeCodeToNameMap[employeeCode] || employeeCode;
        const branchName = employeeEntries[0][HEADER_BRANCH_NAME];
        const designation = employeeCodeToDesignationMap[employeeCode] || 'Default';

        const { totalActivity, productInterests } = calculateTotalActivity(employeeEntries);
        const targets = TARGETS[designation] || TARGETS['Default'];
        const performance = calculatePerformance(totalActivity, targets);

        let html = `
            <h2>Employee Performance Summary</h2>
            <div class="employee-summary-details">
                <p><strong>Employee Name:</strong> ${employeeName}</p>
                <p><strong>Employee Code:</strong> ${employeeCode}</p>
                <p><strong>Branch:</strong> ${branchName}</p>
                <p><strong>Designation:</strong> ${designation}</p>
            </div>
            <h3>Activity Overview</h3>
            <table class="summary-table">
                <thead>
                    <tr>
                        <th>Metric</th>
                        <th>Actual</th>
                        <th>Target</th>
                        <th>%</th>
                    </tr>
                </thead>
                <tbody>
        `;

        const metrics = ['Visit', 'Call', 'Reference', 'New Customer Leads'];
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
            if (actualValue === 0 && targetValue > 0) {
                displayPercent = '0%';
                progressWidth = 0;
                progressBarClass = 'danger';
            }

            html += `
                <tr>
                    <td>${metric}</td>
                    <td>${actualValue}</td>
                    <td>${targetValue}</td>
                    <td>
                        <div class="progress-bar-container">
                            <div class="progress-bar ${progressBarClass}" style="width: ${progressWidth}%">
                                ${displayPercent}
                            </div>
                        </div>
                    </td>
                </tr>
            `;
        });

        html += `
                </tbody>
            </table>
            <h3>Product Interests</h3>
        `;

        if (productInterests.length > 0) {
            html += `<p>${productInterests.join(', ')}</p>`;
        } else {
            html += `<p>No specific product interests recorded.</p>`;
        }

        reportDisplay.innerHTML = html;
    }

    // Render All Entries (for d5.PNG)
    function renderAllEntries(entries) {
        reportDisplay.innerHTML = '<h2>All Canvassing Entries</h2>';
        const tableContainer = document.createElement('div');
        tableContainer.className = 'data-table-container'; // For horizontal scrolling

        const table = document.createElement('table');
        table.className = 'all-entries-table';

        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        const headers = [
            HEADER_TIMESTAMP, HEADER_DATE, HEADER_BRANCH_NAME, HEADER_EMPLOYEE_NAME,
            HEADER_EMPLOYEE_CODE, HEADER_DESIGNATION, HEADER_ACTIVITY_TYPE, HEADER_TYPE_OF_CUSTOMER,
            HEADER_R_LEAD_SOURCE, HEADER_HOW_CONTACTED, HEADER_PROSPECT_NAME,
            HEADER_PHONE_NUMBER_WHATSAPP, HEADER_ADDRESS, HEADER_PROFESSION,
            HEADER_DOB_WD, HEADER_PRODUCT_INTERESTED, HEADER_REMARKS,
            HEADER_NEXT_FOLLOW_UP_DATE, HEADER_RELATION_WITH_STAFF,
            HEADER_FAMILY_DETAILS_1, HEADER_FAMILY_DETAILS_2, HEADER_FAMILY_DETAILS_3,
            HEADER_FAMILY_DETAILS_4, HEADER_PROFILE_OF_CUSTOMER
        ];

        headers.forEach(headerText => {
            const th = document.createElement('th');
            th.textContent = headerText;
            headerRow.appendChild(th);
        });

        const tbody = table.createTBody();
        const selectedMonthValue = monthSelect.value; // Get selected month value
        const filteredData = filterDataByMonth(allCanvassingData, selectedMonthValue);

        if (filteredData.length === 0) {
            reportDisplay.innerHTML += '<p>No entries found for the selected month.</p>';
            return;
        }

        filteredData.forEach(entry => {
            const row = tbody.insertRow();
            headers.forEach(header => {
                const cell = row.insertCell();
                cell.textContent = entry[header] || ''; // Display empty string if data is missing
            });
        });

        tableContainer.appendChild(table);
        reportDisplay.appendChild(tableContainer);
    }

    // Render Branch Visit Leaderboard (d6.PNG)
    function renderBranchVisitLeaderboard() {
        reportDisplay.innerHTML = '<h2>Branch Visit Leaderboard</h2>';
        const branchVisits = {};

        // Get selected month value
        const selectedMonthValue = monthSelect.value;
        const filteredData = filterDataByMonth(allCanvassingData, selectedMonthValue);

        filteredData.forEach(entry => {
            const branch = entry[HEADER_BRANCH_NAME];
            const activityType = entry[HEADER_ACTIVITY_TYPE] ? entry[HEADER_ACTIVITY_TYPE].trim().toLowerCase() : '';
            if (branch && activityType === 'visit') {
                branchVisits[branch] = (branchVisits[branch] || 0) + 1;
            }
        });

        const sortedBranches = Object.entries(branchVisits).sort(([, visitsA], [, visitsB]) => visitsB - visitsA);

        if (sortedBranches.length === 0) {
            reportDisplay.innerHTML += '<p>No visits recorded for any branch in the selected month.</p>';
            return;
        }

        const table = document.createElement('table');
        table.className = 'leaderboard-table';
        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        ['Rank', 'Branch Name', 'Total Visits'].forEach(text => {
            const th = document.createElement('th');
            th.textContent = text;
            headerRow.appendChild(th);
        });

        const tbody = table.createTBody();
        sortedBranches.forEach(([branch, visits], index) => {
            const row = tbody.insertRow();
            row.insertCell().textContent = index + 1;
            row.insertCell().textContent = branch;
            row.insertCell().textContent = visits;
        });

        reportDisplay.appendChild(table);
    }

    // Render Branch Call Leaderboard (d7.PNG)
    function renderBranchCallLeaderboard() {
        reportDisplay.innerHTML = '<h2>Branch Call Leaderboard</h2>';
        const branchCalls = {};

        // Get selected month value
        const selectedMonthValue = monthSelect.value;
        const filteredData = filterDataByMonth(allCanvassingData, selectedMonthValue);

        filteredData.forEach(entry => {
            const branch = entry[HEADER_BRANCH_NAME];
            const activityType = entry[HEADER_ACTIVITY_TYPE] ? entry[HEADER_ACTIVITY_TYPE].trim().toLowerCase() : '';
            if (branch && activityType === 'calls') { // Note: 'calls' lowercase for comparison
                branchCalls[branch] = (branchCalls[branch] || 0) + 1;
            }
        });

        const sortedBranches = Object.entries(branchCalls).sort(([, callsA], [, callsB]) => callsB - callsA);

        if (sortedBranches.length === 0) {
            reportDisplay.innerHTML += '<p>No calls recorded for any branch in the selected month.</p>';
            return;
        }

        const table = document.createElement('table');
        table.className = 'leaderboard-table';
        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        ['Rank', 'Branch Name', 'Total Calls'].forEach(text => {
            const th = document.createElement('th');
            th.textContent = text;
            headerRow.appendChild(th);
        });

        const tbody = table.createTBody();
        sortedBranches.forEach(([branch, calls], index) => {
            const row = tbody.insertRow();
            row.insertCell().textContent = index + 1;
            row.insertCell().textContent = branch;
            row.insertCell().textContent = calls;
        });

        reportDisplay.appendChild(table);
    }

    // Render Staff Participation (d8.PNG)
    function renderStaffParticipation() {
        reportDisplay.innerHTML = '<h2>Staff Participation</h2>';
        const staffParticipation = {};

        // Get selected month value
        const selectedMonthValue = monthSelect.value;
        const filteredData = filterDataByMonth(allCanvassingData, selectedMonthValue);

        // Track all unique employee codes present in the filtered data
        const allEmployeesWithActivity = new Set();
        filteredData.forEach(entry => {
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            if (employeeCode) {
                allEmployeesWithActivity.add(employeeCode);
            }
        });

        // Initialize participation status for all employees who ever appeared in the data
        allUniqueEmployees.forEach(employeeCode => {
            staffParticipation[employeeCode] = {
                name: employeeCodeToNameMap[employeeCode] || employeeCode,
                participated: allEmployeesWithActivity.has(employeeCode)
            };
        });

        // Sort by participation status (participated first), then by name
        const sortedStaff = Object.values(staffParticipation).sort((a, b) => {
            if (a.participated !== b.participated) {
                return a.participated ? -1 : 1; // true comes before false
            }
            return a.name.localeCompare(b.name);
        });

        if (sortedStaff.length === 0) {
            reportDisplay.innerHTML += '<p>No staff data available.</p>';
            return;
        }

        const table = document.createElement('table');
        table.className = 'staff-participation-table';
        const thead = table.createTHead();
        const headerRow = thead.insertRow();
        ['Employee Name', 'Participation Status'].forEach(text => {
            const th = document.createElement('th');
            th.textContent = text;
            headerRow.appendChild(th);
        });

        const tbody = table.createTBody();
        sortedStaff.forEach(staff => {
            const row = tbody.insertRow();
            row.insertCell().textContent = staff.name;
            row.insertCell().textContent = staff.participated ? 'Participated' : 'Not Participated';
            row.classList.add(staff.participated ? 'participated' : 'not-participated'); // Add class for styling
        });

        reportDisplay.appendChild(table);
    }

    // --- Tab and Section Management ---
    function showTab(tabId) {
        // Deactivate all tab buttons
        document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
        // Hide all content sections
        reportsSection.style.display = 'none';
        detailedCustomerViewSection.style.display = 'none';
        employeeManagementSection.style.display = 'none';

        // Activate the selected tab button
        document.getElementById(tabId).classList.add('active');

        // Show the relevant content section and render report
        if (tabId === 'allBranchSnapshotTabBtn') {
            reportsSection.style.display = 'block';
            document.getElementById('branchFilterPanel').style.display = 'none'; // Hide branch filter
            employeeFilterPanel.style.display = 'none'; // Hide employee filter
            viewOptions.style.display = 'none'; // Hide view options
            renderAllBranchSnapshot();
        } else if (tabId === 'nonParticipatingBranchesTabBtn') {
            reportsSection.style.display = 'block';
            document.getElementById('branchFilterPanel').style.display = 'none'; // Hide branch filter
            employeeFilterPanel.style.display = 'none'; // Hide employee filter
            viewOptions.style.display = 'none'; // Hide view options
            renderNonParticipatingBranches();
        } else if (tabId === 'allStaffOverallPerformanceTabBtn') {
            reportsSection.style.display = 'block';
            document.getElementById('branchFilterPanel').style.display = 'none'; // Hide branch filter
            employeeFilterPanel.style.display = 'none'; // Hide employee filter
            viewOptions.style.display = 'none'; // Hide view options
            renderOverallStaffPerformanceReport();
        } else if (tabId === 'branchPerformanceTabBtn') {
            reportsSection.style.display = 'block';
            document.getElementById('branchFilterPanel').style.display = 'flex'; // Show branch filter
            // employeeFilterPanel and viewOptions will be controlled by branchSelect change listener
            // if a branch is already selected, it will trigger the render
            if (branchSelect.value) {
                renderBranchPerformanceReport(branchSelect.value);
            } else {
                reportDisplay.innerHTML = '<p>Please select a branch to view its performance report.</p>';
            }
        } else if (tabId === 'detailedCustomerViewTabBtn') {
            detailedCustomerViewSection.style.display = 'block';
            // Populate customer view specific dropdowns
            populateDropdown(customerViewBranchSelect, allUniqueBranches);
            // Employee dropdown for customer view will be populated when branch is selected
            customerViewBranchSelect.value = ''; // Reset branch selection
            customerViewEmployeeSelect.innerHTML = '<option value="">-- Select Employee --</option>'; // Clear employee dropdown
            detailedCustomerReportTableBody.innerHTML = '<tr><td colspan="5">Select a branch, employee, and month to load customer data.</td></tr>'; // Clear table
            customerDetailsContent.style.display = 'none'; // Hide customer details
        } else if (tabId === 'employeeManagementTabBtn') {
            employeeManagementSection.style.display = 'block';
            populateDropdown(newBranchNameInput, PREDEFINED_BRANCHES);
            populateDropdown(bulkEmployeeBranchNameInput, PREDEFINED_BRANCHES);
        }
    }

    // --- Event Listeners for Tab Buttons ---
    allBranchSnapshotTabBtn.addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    nonParticipatingBranchesTabBtn.addEventListener('click', () => showTab('nonParticipatingBranchesTabBtn'));
    allStaffOverallPerformanceTabBtn.addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    branchPerformanceTabBtn.addEventListener('click', () => showTab('branchPerformanceTabBtn'));
    detailedCustomerViewTabBtn.addEventListener('click', () => showTab('detailedCustomerViewTabBtn'));
    employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));


    // --- Event Listeners for View Option Buttons (within main reports section) ---
    viewBranchPerformanceReportBtn.addEventListener('click', () => {
        document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        viewBranchPerformanceReportBtn.classList.add('active');
        const selectedBranch = branchSelect.value;
        if (selectedBranch) {
            renderBranchPerformanceReport(selectedBranch);
        } else {
            reportDisplay.innerHTML = '<p>Please select a branch from the dropdown.</p>';
        }
    });

    viewEmployeeSummaryBtn.addEventListener('click', () => {
        document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        viewEmployeeSummaryBtn.classList.add('active');
        if (selectedEmployeeCodeEntries.length > 0) {
            renderEmployeeSummary(selectedEmployeeCodeEntries);
        } else {
            reportDisplay.innerHTML = '<p>Please select an employee to view their summary.</p>';
        }
    });

    viewAllEntriesBtn.addEventListener('click', () => {
        document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        viewAllEntriesBtn.classList.add('active');
        // Filter by month first, then render all entries for that month
        const selectedMonthValue = monthSelect.value;
        const filteredByMonthData = filterDataByMonth(allCanvassingData, selectedMonthValue);
        renderAllEntries(filteredByMonthData);
    });

    viewBranchVisitLeaderboardBtn.addEventListener('click', () => {
        document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        viewBranchVisitLeaderboardBtn.classList.add('active');
        renderBranchVisitLeaderboard();
    });

    viewBranchCallLeaderboardBtn.addEventListener('click', () => {
        document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        viewBranchCallLeaderboardBtn.classList.add('active');
        renderBranchCallLeaderboard();
    });

    viewStaffParticipationBtn.addEventListener('click', () => {
        document.querySelectorAll('.view-options .btn').forEach(btn => btn.classList.remove('active'));
        viewStaffParticipationBtn.classList.add('active');
        renderStaffParticipation();
    });

    // --- Detailed Customer View Event Listeners ---
    customerViewBranchSelect.addEventListener('change', () => {
        const selectedBranch = customerViewBranchSelect.value;
        const selectedMonthValue = customerViewMonthSelect.value; // Get selected month value

        // Filter allCanvassingData by selected month first
        let filteredByMonthData = filterDataByMonth(allCanvassingData, selectedMonthValue);

        if (selectedBranch) {
            const employeeCodesInBranchFromCanvassing = filteredByMonthData
                .filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch)
                .map(entry => entry[HEADER_EMPLOYEE_CODE]);

            const combinedEmployeeCodes = new Set([...employeeCodesInBranchFromCanvassing]);
            const sortedEmployeeCodesInBranch = [...combinedEmployeeCodes].sort((codeA, codeB) => {
                const nameA = employeeCodeToNameMap[codeA] || codeA;
                const nameB = employeeCodeToNameMap[codeB] || codeB;
                return nameA.localeCompare(nameB);
            });
            populateDropdown(customerViewEmployeeSelect, sortedEmployeeCodesInBranch, true);
            
            customerViewEmployeeSelect.value = ''; // Reset employee selection
            detailedCustomerReportTableBody.innerHTML = '<tr><td colspan="5">Select an employee and month to load customer data.</td></tr>';
            customerDetailsContent.style.display = 'none'; // Hide customer details
        } else {
            customerViewEmployeeSelect.innerHTML = '<option value="">-- Select Employee --</option>';
            detailedCustomerReportTableBody.innerHTML = '<tr><td colspan="5">Select a branch, employee, and month to load customer data.</td></tr>';
            customerDetailsContent.style.display = 'none'; // Hide customer details
        }
    });

    customerViewEmployeeSelect.addEventListener('change', () => {
        loadDetailedCustomerReport();
    });

    downloadDetailedCustomerReportBtn.addEventListener('click', () => {
        downloadDetailedCustomerReportCSV();
    });

    // Function to download the detailed customer report as CSV
    function downloadDetailedCustomerReportCSV() {
        const selectedBranch = customerViewBranchSelect.value;
        const selectedEmployeeCode = customerViewEmployeeSelect.value;
        const selectedMonthValue = customerViewMonthSelect.value;

        if (!selectedBranch || !selectedEmployeeCode || !selectedMonthValue) {
            displayMessage("Please select a Branch, Employee, and Month before downloading.", 'info');
            return;
        }

        let filteredCustomerEntries = filterDataByMonth(allCanvassingData, selectedMonthValue);
        filteredCustomerEntries = filteredCustomerEntries.filter(entry =>
            entry[HEADER_BRANCH_NAME] === selectedBranch &&
            entry[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode
        );

        if (filteredCustomerEntries.length === 0) {
            displayMessage("No customer entries found for the selected criteria to download.", 'info');
            return;
        }

        const headers = [
            HEADER_TIMESTAMP, HEADER_DATE, HEADER_BRANCH_NAME, HEADER_EMPLOYEE_NAME,
            HEADER_EMPLOYEE_CODE, HEADER_DESIGNATION, HEADER_ACTIVITY_TYPE, HEADER_TYPE_OF_CUSTOMER,
            HEADER_R_LEAD_SOURCE, HEADER_HOW_CONTACTED, HEADER_PROSPECT_NAME,
            HEADER_PHONE_NUMBER_WHATSAPP, HEADER_ADDRESS, HEADER_PROFESSION,
            HEADER_DOB_WD, HEADER_PRODUCT_INTERESTED, HEADER_REMARKS,
            HEADER_NEXT_FOLLOW_UP_DATE, HEADER_RELATION_WITH_STAFF,
            HEADER_FAMILY_DETAILS_1, HEADER_FAMILY_DETAILS_2, HEADER_FAMILY_DETAILS_3,
            HEADER_FAMILY_DETAILS_4, HEADER_PROFILE_OF_CUSTOMER
        ];

        const csvRows = [];
        csvRows.push(headers.map(h => `"${h.replace(/"/g, '""')}"`).join(',')); // Quote headers

        filteredCustomerEntries.forEach(entry => {
            const rowData = headers.map(header => {
                const cellValue = entry[header] || '';
                // Ensure values are properly quoted if they contain commas or quotes
                const stringCell = String(cellValue);
                return `"${stringCell.replace(/"/g, '""')}"`;
            });
            csvRows.push(rowData.join(','));
        });

        const csvString = csvRows.join('\n');
        const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        if (link.download !== undefined) {
            const url = URL.createObjectURL(blob);
            link.setAttribute('href', url);
            link.setAttribute('download', `Customer_Report_${selectedEmployeeCode}_${selectedMonthValue}.csv`);
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            displayMessage("Detailed Customer Report downloaded successfully!", 'success');
        } else {
            displayMessage("Your browser does not support automatic downloads. Please copy the data manually.", 'error');
            console.log(csvString);
        }
    }


    // Function to display customer details in cards and table
    function displayCustomerDetails(customerEntry) {
        customerDetailsContent.style.display = 'block';

        // Clear previous content
        customerCard1.innerHTML = '';
        customerCard2.innerHTML = '';
        customerCard3.innerHTML = '';

        // Card 1: Basic Info
        customerCard1.innerHTML = `
            <h3>Customer Info</h3>
            <p><strong>Prospect Name:</strong> ${customerEntry[HEADER_PROSPECT_NAME] || 'N/A'}</p>
            <p><strong>Phone:</strong> ${customerEntry[HEADER_PHONE_NUMBER_WHATSAPP] || 'N/A'}</p>
            <p><strong>Address:</strong> ${customerEntry[HEADER_ADDRESS] || 'N/A'}</p>
            <p><strong>Profession:</strong> ${customerEntry[HEADER_PROFESSION] || 'N/A'}</p>
            <p><strong>DOB/WD:</strong> ${customerEntry[HEADER_DOB_WD] || 'N/A'}</p>
            <p><strong>Type of Customer:</strong> ${customerEntry[HEADER_TYPE_OF_CUSTOMER] || 'N/A'}</p>
            <p><strong>rLead Source:</strong> ${customerEntry[HEADER_R_LEAD_SOURCE] || 'N/A'}</p>
            <p><strong>How Contacted:</strong> ${customerEntry[HEADER_HOW_CONTACTED] || 'N/A'}</p>
        `;

        // Card 2: Financial/Product Interest
        customerCard2.innerHTML = `
            <h3>Product & Remarks</h3>
            <p><strong>Product Interested:</strong> ${customerEntry[HEADER_PRODUCT_INTERESTED] || 'N/A'}</p>
            <p><strong>Remarks:</strong> ${customerEntry[HEADER_REMARKS] || 'N/A'}</p>
            <p><strong>Next Follow-up Date:</strong> ${customerEntry[HEADER_NEXT_FOLLOW_UP_DATE] || 'N/A'}</p>
            <p><strong>Relation With Staff:</strong> ${customerEntry[HEADER_RELATION_WITH_STAFF] || 'N/A'}</p>
        `;

        // Card 3: Family & Profile Details
        customerCard3.innerHTML = `
            <h3>Family & Profile</h3>
            <p><strong>Spouse Name:</strong> ${customerEntry[HEADER_FAMILY_DETAILS_1] || 'N/A'}</p>
            <p><strong>Spouse Job:</strong> ${customerEntry[HEADER_FAMILY_DETAILS_2] || 'N/A'}</p>
            <p><strong>Children Names:</strong> ${customerEntry[HEADER_FAMILY_DETAILS_3] || 'N/A'}</p>
            <p><strong>Children Details:</strong> ${customerEntry[HEADER_FAMILY_DETAILS_4] || 'N/A'}</p>
            <p><strong>Customer Profile:</strong> ${customerEntry[HEADER_PROFILE_OF_CUSTOMER] || 'N/A'}</p>
        `;
    }

    // --- Employee Management Functions ---

    // Add Employee Form Submission
    addEmployeeForm.addEventListener('submit', async (event) => {
        event.preventDefault(); // Prevent default form submission

        const name = newEmployeeNameInput.value.trim();
        const code = newEmployeeCodeInput.value.trim();
        const branch = newBranchNameInput.value.trim();
        const designation = newDesignationInput.value.trim();

        if (!name || !code || !branch || !designation) {
            displayEmployeeManagementMessage('All fields are required for adding an employee.', true);
            return;
        }

        displayEmployeeManagementMessage('Adding employee...', false);
        try {
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: new URLSearchParams({
                    action: 'addEmployee',
                    name: name,
                    code: code,
                    branch: branch,
                    designation: designation
                })
            });

            const result = await response.json();

            if (result.status === 'SUCCESS') {
                displayEmployeeManagementMessage('Employee added successfully! Data will refresh shortly.', false);
                addEmployeeForm.reset(); // Clear form fields
                await processData(); // Re-fetch and process data to update dropdowns
            } else {
                displayEmployeeManagementMessage(`Error adding employee: ${result.message || 'Unknown error'}`, true);
            }
        } catch (error) {
            console.error('Error adding employee:', error);
            displayEmployeeManagementMessage(`Network error or API issue: ${error.message}`, true);
        }
    });

    // Bulk Add Employee Form Submission
    bulkAddEmployeeForm.addEventListener('submit', async (event) => {
        event.preventDefault();

        const branch = bulkEmployeeBranchNameInput.value.trim();
        const detailsText = bulkEmployeeDetailsTextarea.value.trim();

        if (!branch || !detailsText) {
            displayEmployeeManagementMessage('Branch and employee details are required for bulk addition.', true);
            return;
        }

        // Parse the textarea content: Name,Code,Designation (one per line)
        const lines = detailsText.split('\n').filter(line => line.trim() !== '');
        const employeesToAdd = [];
        for (const line of lines) {
            const parts = line.split(',').map(p => p.trim());
            if (parts.length === 3) {
                employeesToAdd.push({ name: parts[0], code: parts[1], designation: parts[2], branch: branch });
            } else {
                displayEmployeeManagementMessage(`Skipping malformed line: "${line}". Expected Name,Code,Designation.`, true);
                return; // Stop processing if any line is malformed, or adjust to continue
            }
        }

        if (employeesToAdd.length === 0) {
            displayEmployeeManagementMessage('No valid employee entries found for bulk addition.', true);
            return;
        }

        displayEmployeeManagementMessage(`Adding ${employeesToAdd.length} employees in bulk...`, false);

        try {
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: new URLSearchParams({
                    action: 'bulkAddEmployees',
                    employees: JSON.stringify(employeesToAdd) // Send as JSON string
                })
            });

            const result = await response.json();

            if (result.status === 'SUCCESS') {
                displayEmployeeManagementMessage(`${result.count || employeesToAdd.length} employees added successfully! Data will refresh shortly.`, false);
                bulkAddEmployeeForm.reset();
                await processData();
            } else {
                displayEmployeeManagementMessage(`Error bulk adding employees: ${result.message || 'Unknown error'}`, true);
            }
        } catch (error) {
            console.error('Error bulk adding employees:', error);
            displayEmployeeManagementMessage(`Network error or API issue: ${error.message}`, true);
        }
    });


    // Delete Employee Form Submission
    deleteEmployeeForm.addEventListener('submit', async (event) => {
        event.preventDefault();

        const codeToDelete = deleteEmployeeCodeInput.value.trim();

        if (!codeToDelete) {
            displayEmployeeManagementMessage('Employee Code is required for deletion.', true);
            return;
        }

        if (!confirm(`Are you sure you want to delete employee with code: ${codeToDelete}? This action cannot be undone.`)) {
            return; // User cancelled
        }

        displayEmployeeManagementMessage('Deleting employee...', false);
        try {
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: new URLSearchParams({
                    action: 'deleteEmployee',
                    code: codeToDelete
                })
            });

            const result = await response.json();

            if (result.status === 'SUCCESS') {
                displayEmployeeManagementMessage('Employee deleted successfully! Data will refresh shortly.', false);
                deleteEmployeeForm.reset();
                await processData(); // Re-fetch and process data to update dropdowns
            } else {
                displayEmployeeManagementMessage(`Error deleting employee: ${result.message || 'Unknown error'}`, true);
            }
        } catch (error) {
            console.error('Error deleting employee:', error);
            displayEmployeeManagementMessage(`Network error or API issue: ${error.message}`, true);
        }
    });


    // Initial data fetch and tab display when the page loads
    processData();
    showTab('allBranchSnapshotTabBtn');
}); // This is the closing brace for DOMContentLoaded
