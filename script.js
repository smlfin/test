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
    // const viewOptions = document.getElementById('viewOptions'); // No longer directly used for visibility
    const viewBranchPerformanceReportBtn = document.getElementById('viewBranchPerformanceReportBtn');
    const viewEmployeeSummaryBtn = document.getElementById('viewEmployeeSummaryBtn');
    const viewAllEntriesBtn = document.getElementById('viewAllEntriesBtn');
    const viewPerformanceReportBtn = document.getElementById('viewPerformanceReportBtn');

    // Main Report Display Area (now within tab contents)
    // No longer a single 'reportDisplay' but specific table bodies or divs
    const allBranchSnapshotTableBody = document.getElementById('branchSnapshotTableBody');
    const overallPerformanceTableBody = document.getElementById('overallPerformanceTableBody');
    const customerDetailView = document.getElementById('customerDetailView');
    const overallPerformanceListView = document.getElementById('overallPerformanceListView');
    const backToOverallListBtn = document.getElementById('backToOverallListBtn');

    // Specific detail elements for the customer report
    const detailDate = document.getElementById('detailDate');
    const detailBranchName = document.getElementById('detailBranchName');
    const detailEmployeeName = document.getElementById('detailEmployeeName');
    const detailDesignation = document.getElementById('detailDesignation');
    const detailCustomerName = document.getElementById('detailCustomerName');
    const detailAddress = document.getElementById('detailAddress');
    const detailPhoneNumber = document.getElementById('detailPhoneNumber');
    const detailEmailId = document.getElementById('detailEmailId');
    const detailRemarks = document.getElementById('detailRemarks');
    const detailFollowupDate = document.getElementById('detailFollowupDate');


    // NEW: Dedicated message area element
    const statusMessageDiv = document.getElementById('statusMessage');


    // Tab buttons for main navigation
    const allBranchSnapshotTabBtn = document.getElementById('allBranchSnapshotTabBtn');
    const allStaffOverallPerformanceTabBtn = document.getElementById('allStaffOverallPerformanceTabBtn');
    const nonParticipatingBranchesTabBtn = document.getElementById('nonParticipatingBranchesTabBtn');
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');

    // Main Content Sections to toggle
    const reportsSection = document.getElementById('reportsSection');
    const allBranchSnapshotTabContent = document.getElementById('allBranchSnapshotTabContent');
    const allStaffOverallPerformanceTabContent = document.getElementById('allStaffOverallPerformanceTabContent');
    const nonParticipatingBranchesTabContent = document.getElementById('nonParticipatingBranchesTabContent');
    const employeeManagementSection = document.getElementById('employeeManagementSection');
    const noParticipationMessage = document.getElementById('noParticipationMessage');
    const nonParticipatingBranchesList = document.getElementById('nonParticipatingBranchesList');


    // Employee Management Form Elements
    const addEmployeeForm = document.getElementById('addEmployeeForm');
    const employeeNameInput = document.getElementById('employeeName'); // Corrected from newEmployeeNameInput
    const employeeCodeInput = document.getElementById('employeeCode');   // Corrected from newEmployeeCodeInput
    const employeeBranchInput = document.getElementById('employeeBranch'); // Corrected from newBranchNameInput
    const employeeDesignationInput = document.getElementById('employeeDesignation'); // Corrected from newDesignationInput
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
        if (isNaN(date.getTime())) return dateString; // Return original if invalid date
        return date.toISOString().split('T')[0];
    };

    // Helper to clear and display messages in a specific div
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
        await fetchCanvassingData();

        allUniqueBranches = [...PREDEFINED_BRANCHES].sort();

        employeeCodeToNameMap = {};
        employeeCodeToDesignationMap = {};
        allCanvassingData.forEach(entry => {
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            const employeeName = entry[HEADER_EMPLOYEE_NAME];
            const designation = entry[HEADER_DESIGNATION];

            if (employeeCode) {
                employeeCodeToNameMap[employeeCode] = employeeName || employeeCode;
                employeeCodeToDesignationMap[employeeCode] = designation || 'Default';
            }
        });

        allUniqueEmployees = [...new Set(allCanvassingData.map(entry => entry[HEADER_EMPLOYEE_CODE]))].sort((codeA, codeB) => {
            const nameA = employeeCodeToNameMap[codeA] || codeA;
            const nameB = employeeCodeToNameMap[codeB] || codeB;
            return nameA.localeCompare(nameB);
        });

        populateDropdown(branchSelect, allUniqueBranches);
        console.log('Final All Unique Branches (Predefined):', allUniqueBranches);
        console.log('Final Employee Code To Name Map (from Canvassing Data):', employeeCodeToNameMap);
        console.log('Final Employee Code To Designation Map (from Canvassing Data):', employeeCodeToDesignationMap);
        console.log('Final All Unique Employees (Codes from Canvassing Data):', allUniqueEmployees);

        // After data is loaded and maps are populated, render the initial report
        renderAllBranchSnapshot();
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

    // Filter employees based on selected branch (functionality remains, but for specific tab)
    branchSelect.addEventListener('change', () => {
        const selectedBranch = branchSelect.value;
        if (selectedBranch) {
            employeeFilterPanel.style.display = 'block';

            const employeeCodesInBranchFromCanvassing = allCanvassingData
                .filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch)
                .map(entry => entry[HEADER_EMPLOYEE_CODE]);

            const combinedEmployeeCodes = new Set([
                ...employeeCodesInBranchFromCanvassing
            ]);

            const sortedEmployeeCodesInBranch = [...combinedEmployeeCodes].sort((codeA, codeB) => {
                const nameA = employeeCodeToNameMap[codeA] || codeA;
                const nameB = employeeCodeToNameMap[codeB] || codeB;
                return nameA.localeCompare(nameB);
            });

            populateDropdown(employeeSelect, sortedEmployeeCodesInBranch, true);
            // viewOptions.style.display = 'flex'; // This panel is no longer used for tab content
            employeeSelect.value = "";
            selectedEmployeeCodeEntries = [];

        } else {
            employeeFilterPanel.style.display = 'none';
            // viewOptions.style.display = 'none'; // This panel is no longer used for tab content
            selectedBranchEntries = [];
            selectedEmployeeCodeEntries = [];
        }
        // When branch changes, ensure the correct tab content is shown for active tab
        const activeTabButton = document.querySelector('.tab-button.active');
        if (activeTabButton.id === 'allBranchSnapshotTabBtn') {
            renderAllBranchSnapshot(); // Re-render for 'All Branch Snapshot' tab
        } else if (activeTabButton.id === 'allStaffOverallPerformanceTabBtn') {
            // No direct re-render for this tab on branch select change, as it lists all employees
        }
    });

    // Handle employee selection (now based on employee CODE) - mainly for individual reports
    employeeSelect.addEventListener('change', () => {
        const selectedEmployeeCode = employeeSelect.value;
        if (selectedEmployeeCode) {
            selectedEmployeeCodeEntries = allCanvassingData.filter(entry =>
                entry[HEADER_EMPLOYEE_CODE] === selectedEmployeeCode &&
                entry[HEADER_BRANCH_NAME] === branchSelect.value
            );
            // No automatic rendering of specific reports here, as per new tab structure
        } else {
            selectedEmployeeCodeEntries = [];
        }
    });

    // Helper to calculate total activity from a set of activity entries based on Activity Type
    function calculateTotalActivity(entries) {
        const totalActivity = { 'Visit': 0, 'Call': 0, 'Reference': 0, 'New Customer Leads': 0 }; // Initialize counters
        const productInterests = new Set(); // To collect unique product interests

        entries.forEach(entry => {
            let activityType = entry[HEADER_ACTIVITY_TYPE];
            let typeOfCustomer = entry[HEADER_TYPE_OF_CUSTOMER];
            let productInterested = entry[HEADER_PRODUCT_INTERESTED];

            const trimmedActivityType = activityType ? activityType.trim().toLowerCase() : '';
            const trimmedTypeOfCustomer = typeOfCustomer ? typeOfCustomer.trim().toLowerCase() : '';
            const trimmedProductInterested = productInterested ? productInterested.trim() : '';

            if (trimmedActivityType === 'visit') {
                totalActivity['Visit']++;
            } else if (trimmedActivityType === 'calls') {
                totalActivity['Call']++;
            } else if (trimmedActivityType === 'referance') {
                totalActivity['Reference']++;
            }

            if (trimmedTypeOfCustomer === 'new') {
                totalActivity['New Customer Leads']++;
            }

            if (trimmedProductInterested) {
                productInterests.add(trimmedProductInterested);
            }
        });
        return { totalActivity, productInterests: [...productInterests] };
    }

    // Render All Branch Snapshot
    function renderAllBranchSnapshot() {
        allBranchSnapshotTableBody.innerHTML = ''; // Clear existing content

        PREDEFINED_BRANCHES.forEach(branch => {
            const branchActivityEntries = allCanvassingData.filter(entry => entry[HEADER_BRANCH_NAME] === branch);
            const { totalActivity } = calculateTotalActivity(branchActivityEntries);
            const employeeCodesInBranch = [...new Set(branchActivityEntries.map(entry => entry[HEADER_EMPLOYEE_CODE]))];
            const displayEmployeeCount = employeeCodesInBranch.length;

            const row = allBranchSnapshotTableBody.insertRow();
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
    }

    // Render Non-Participating Branches Report
    function renderNonParticipatingBranches() {
        nonParticipatingBranchesList.innerHTML = '';
        noParticipationMessage.style.display = 'none';

        const nonParticipatingBranches = [];
        PREDEFINED_BRANCHES.forEach(branch => {
            const hasActivity = allCanvassingData.some(entry => entry[HEADER_BRANCH_NAME] === branch);
            if (!hasActivity) {
                nonParticipatingBranches.push(branch);
            }
        });

        if (nonParticipatingBranches.length > 0) {
            nonParticipatingBranches.forEach(branch => {
                const li = document.createElement('li');
                li.textContent = branch;
                nonParticipatingBranchesList.appendChild(li);
            });
        } else {
            noParticipationMessage.style.display = 'block';
        }
    }


    // NEW: Render "All Staff Performance (Overall)" - displays only Date and Customer Name (hyperlink)
    function renderOverallStaffPerformanceSummaryList() {
        overallPerformanceTableBody.innerHTML = ''; // Clear previous entries
        overallPerformanceListView.style.display = 'block';
        customerDetailView.style.display = 'none';

        if (allCanvassingData.length === 0) {
            const row = overallPerformanceTableBody.insertRow();
            const cell = row.insertCell();
            cell.colSpan = 2; // Span across two columns
            cell.textContent = 'No canvassing entries found.';
            cell.classList.add('no-participation-message-cell');
            return;
        }

        allCanvassingData.forEach((entry, index) => {
            const row = overallPerformanceTableBody.insertRow();
            const dateCell = row.insertCell();
            const customerNameCell = row.insertCell();

            dateCell.textContent = formatDate(entry[HEADER_DATE]);
            dateCell.setAttribute('data-label', 'Date');

            const customerNameLink = document.createElement('a');
            customerNameLink.href = '#'; // Make it a clickable link
            customerNameLink.textContent = entry[HEADER_PROSPECT_NAME] || 'N/A';
            customerNameLink.classList.add('customer-name-link');
            customerNameLink.dataset.entryIndex = index; // Store index to retrieve full entry later

            customerNameLink.addEventListener('click', (e) => {
                e.preventDefault();
                const clickedIndex = parseInt(e.target.dataset.entryIndex);
                if (!isNaN(clickedIndex) && allCanvassingData[clickedIndex]) {
                    renderCustomerDetailReport(allCanvassingData[clickedIndex]);
                }
            });

            customerNameCell.appendChild(customerNameLink);
            customerNameCell.setAttribute('data-label', 'Customer Name');
        });
    }

    // NEW: Render Detailed Customer Report
    function renderCustomerDetailReport(entry) {
        overallPerformanceListView.style.display = 'none'; // Hide the list view
        customerDetailView.style.display = 'block'; // Show the detail view

        detailDate.textContent = formatDate(entry[HEADER_DATE]);
        detailBranchName.textContent = entry[HEADER_BRANCH_NAME] || 'N/A';
        detailEmployeeName.textContent = entry[HEADER_EMPLOYEE_NAME] || 'N/A';
        detailDesignation.textContent = entry[HEADER_DESIGNATION] || 'N/A';
        detailCustomerName.textContent = entry[HEADER_PROSPECT_NAME] || 'N/A';
        detailAddress.textContent = entry[HEADER_ADDRESS] || 'N/A';
        detailPhoneNumber.textContent = entry[HEADER_PHONE_NUMBER_WHATSAPP] || 'N/A';
        // Note: Email ID might not be a standard header, check your sheet. Using HEADER_EMAIL_ID if it exists.
        // For now, assuming it's not a direct column, you might need to add it or fetch from remarks.
        // If you have an 'Email ID' column, uncomment/change below:
        // detailEmailId.textContent = entry['Email ID'] || 'N/A';
        detailEmailId.textContent = 'Not Available'; // Placeholder if no direct email column
        detailRemarks.textContent = entry[HEADER_REMARKS] || 'N/A';
        detailFollowupDate.textContent = formatDate(entry[HEADER_NEXT_FOLLOW_UP_DATE]);
    }

    // Event Listener for "Back to List" button
    if (backToOverallListBtn) {
        backToOverallListBtn.addEventListener('click', () => {
            renderOverallStaffPerformanceSummaryList(); // Go back to the list view
        });
    }

    // Function to calculate performance percentage
    function calculatePerformance(actual, target) {
        const performance = {};
        for (const key in actual) {
            if (target[key] !== undefined && target[key] > 0) {
                performance[key] = (actual[key] / target[key]) * 100;
            } else {
                performance[key] = NaN;
            }
        }
        return performance;
    }

    // Helper for progress bar styling
    function getProgressBarClass(percentage) {
        if (isNaN(percentage)) return 'no-activity';
        const p = parseFloat(percentage);
        if (p >= 100) return 'overachieved';
        if (p >= 75) return 'success';
        if (p >= 50) return 'warning';
        return 'danger';
    }

    // Function to send data to Google Apps Script
    async function sendDataToGoogleAppsScript(actionType, data = {}) {
        displayEmployeeManagementMessage('Processing request...', false);

        try {
            const formData = new URLSearchParams();
            formData.append('actionType', actionType);
            formData.append('data', JSON.stringify(data));

            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded'
                },
                body: formData
            });

            if (!response.ok) {
                const errorText = await response.text();
                console.error(`HTTP error from Apps Script! Status: ${response.status}. Details: ${errorText}`);
                throw new Error(`Failed to send data to Apps Script. Status: ${response.status}. Please check WEB_APP_URL and Apps Script deployment.`);
            }

            const result = await response.json();

            if (result.status === 'SUCCESS') {
                displayEmployeeManagementMessage(result.message, false);
                return true;
            } else {
                displayEmployeeManagementMessage(`Error: ${result.message}`, true);
                return false;
            }
        } catch (error) {
            console.error('Error sending data to Apps Script:', error);
            displayEmployeeManagementMessage(`Error sending data: ${error.message}. Please check WEB_APP_URL and Apps Script deployment.`, true);
            return false;
        } finally {
            // Re-fetch all data to ensure reports are up-to-date after any employee management action
            await processData(); // Re-fetch canvassing data and re-populate maps/dropdowns
            // Re-render the current report or provide a message
            const activeTabButton = document.querySelector('.tab-button.active');
            if (activeTabButton && reportsSection.style.display === 'block') { // Only re-render if we're on a reports tab
                if (activeTabButton.id === 'allBranchSnapshotTabBtn') {
                    renderAllBranchSnapshot();
                } else if (activeTabButton.id === 'allStaffOverallPerformanceTabBtn') {
                    renderOverallStaffPerformanceSummaryList(); // Re-render the list view
                } else if (activeTabButton.id === 'nonParticipatingBranchesTabBtn') {
                    renderNonParticipatingBranches();
                }
            }
        }
    }


    // Function to manage tab visibility
    function showTab(tabButtonId) {
        document.querySelectorAll('.tab-button').forEach(button => {
            button.classList.remove('active');
        });
        document.getElementById(tabButtonId).classList.add('active');

        // Hide all tab contents initially
        allBranchSnapshotTabContent.style.display = 'none';
        allStaffOverallPerformanceTabContent.style.display = 'none';
        nonParticipatingBranchesTabContent.style.display = 'none';
        employeeManagementSection.style.display = 'none';

        reportsSection.style.display = 'none'; // Default hide reports section

        // Hide filtering controls by default, show only if needed by tab
        document.querySelector('.controls-panel').style.display = 'none';
        employeeFilterPanel.style.display = 'none'; // Also hide employee specific filter

        // Reset dropdowns if we are switching away from a branch/employee dependent tab
        branchSelect.value = '';
        employeeSelect.value = '';

        // Show relevant content based on tab selection
        if (tabButtonId === 'allBranchSnapshotTabBtn') {
            reportsSection.style.display = 'block';
            allBranchSnapshotTabContent.style.display = 'block';
            document.querySelector('.controls-panel').style.display = 'flex'; // Show controls
            renderAllBranchSnapshot();
        } else if (tabButtonId === 'allStaffOverallPerformanceTabBtn') {
            reportsSection.style.display = 'block';
            allStaffOverallPerformanceTabContent.style.display = 'block';
            document.querySelector('.controls-panel').style.display = 'none'; // Hide controls for this tab as it shows all entries
            renderOverallStaffPerformanceSummaryList(); // Render the new customer list view
        } else if (tabButtonId === 'nonParticipatingBranchesTabBtn') {
            reportsSection.style.display = 'block';
            nonParticipatingBranchesTabContent.style.display = 'block';
            document.querySelector('.controls-panel').style.display = 'none'; // Hide controls for this tab
            renderNonParticipatingBranches();
        } else if (tabButtonId === 'employeeManagementTabBtn') {
            employeeManagementSection.style.display = 'block';
            displayEmployeeManagementMessage('', false); // Clear messages
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
    if (employeeManagementTabBtn) {
        employeeManagementTabBtn.addEventListener('click', () => showTab('employeeManagementTabBtn'));
    }

    // Event Listener for Add Employee Form
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();

            const employeeName = employeeNameInput.value.trim();
            const employeeCode = employeeCodeInput.value.trim();
            const branchName = employeeBranchInput.value.trim();
            const designation = employeeDesignationInput.value.trim();

            if (!employeeName || !employeeCode || !branchName) {
                displayEmployeeManagementMessage('Please fill in Employee Name, Code, and Branch Name.', true);
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
                displayEmployeeManagementMessage('Branch Name and Employee Details are required for bulk entry.', true);
                return;
            }

            const employeeLines = bulkDetails.split('\n').filter(line => line.trim() !== '');
            const employeesToAdd = [];

            for (const line of employeeLines) {
                const parts = line.split(',').map(part => part.trim());
                if (parts.length < 2) {
                    displayEmployeeManagementMessage(`Skipping invalid line: "${line}". Each line must have at least Employee Name and Employee Code.`, true);
                    continue;
                }

                const employeeData = {
                    [HEADER_EMPLOYEE_NAME]: parts[0],
                    [HEADER_EMPLOYEE_CODE]: parts[1],
                    [HEADER_BRANCH_NAME]: branchName,
                    [HEADER_DESIGNATION]: parts[2] || ''
                };
                employeesToAdd.push(employeeData);
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
    showTab('allBranchSnapshotTabBtn'); // Default tab to display on load
});
