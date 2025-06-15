document.addEventListener('DOMContentLoaded', () => {
    // *** Configuration ***
    // This URL is for your Canvassing Data sheet. Ensure it's correct and published as CSV.
    const DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=696550092&single=true&output=csv";

    // IMPORTANT: Replace this with YOUR DEPLOYED GOOGLE APPS SCRIPT WEB APP URL
    const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzEYf0CKgwP0O4-z1lup1lDZImD1dQVEveLWsHwa_7T5ltndfIuRWXVZqFDj03_proD/exec"; // <-- PASTE YOUR NEWLY DEPLOYED WEB APP URL HERE

    // NEW: URL for your published MasterEmployees sheet (as CSV).
    const EMPLOYEE_MASTER_DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=696550092&single=true&output=csv"; // This should be your *actual* Master Employees URL

    // --- Column Headers Mapping (IMPORTANT: Adjust these to match your actual CSV headers) ---
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
    const HEADER_FAMILY_DETAILS_1 = 'Family Deatils -1 Name of wife/Husband';
    const HEADER_FAMILY_DETAILS_2 = 'Family Deatils -2 Job of wife/Husband';
    const HEADER_FAMILY_DETAILS_3 = 'Family Deatils -3 Names of Children';
    const HEADER_FAMILY_DETAILS_4 = 'Family Deatils -4 Deatils of Children';
    const HEADER_PROFILE_OF_CUSTOMER = 'Profile of Customer';

    // --- Global Variables (will be populated after data fetch) ---
    let allCanvassingData = []; // To store all fetched canvassing entries
    let employeeMasterData = []; // To store employee master data
    let branchNames = new Set(); // To store unique branch names
    let allEmployees = new Set(); // To store all unique employee names from canvassing data
    let allEmployeeCodes = new Set(); // To store all unique employee codes from canvassing data

    // --- DOM Elements ---
    const statusMessageDiv = document.getElementById('statusMessage');
    const branchSelect = document.getElementById('branchSelect');
    const employeeSelect = document.getElementById('employeeSelect');
    const reportsSection = document.getElementById('reportsSection');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel');
    const branchSummaryTableBody = document.getElementById('branchSummaryTableBody');
    // Removed allStaffOverallPerformanceTableBody as it does not exist in HTML
    const nonParticipatingBranchesTableBody = document.getElementById('nonParticipatingBranchesTableBody');

    // Detailed Customer View Elements
    // detailedCustomerViewContent is not defined in HTML, so removed from here as well.
    const customerFilterBranchSelect = document.getElementById('customerFilterBranchSelect');
    const customerFilterEmployeeSelect = document.getElementById('customerFilterEmployeeSelect');
    const customerDetailsTableBody = document.getElementById('customerDetailsTableBody');
    const customerFilterApplyButton = document.getElementById('customerFilterApplyButton');

    // Modal Elements
    const customerDetailsModal = document.getElementById('customerDetailsModal');
    const closeModalButton = document.querySelector('.modal .close-button');
    const modalBodyContent = document.getElementById('modalBodyContent');

    // Employee Management Elements
    const employeeManagementTabBtn = document.getElementById('employeeManagementTabBtn');
    const employeeManagementSection = document.getElementById('employeeManagementSection');
    const addEmployeeForm = document.getElementById('addEmployeeForm');
    const bulkAddEmployeeForm = document.getElementById('bulkAddEmployeeForm');
    const deleteEmployeeForm = document.getElementById('deleteEmployeeForm');
    const employeeManagementMessage = document.getElementById('employeeManagementMessage');
    const employeeBranchNameInput = document.getElementById('employeeBranchName');
    const employeeNameInput = document.getElementById('employeeName');
    const employeeCodeInput = document.getElementById('employeeCode');
    const employeeDesignationInput = document.getElementById('employeeDesignation');
    const bulkEmployeeBranchNameInput = document.getElementById('bulkEmployeeBranchName');
    const bulkEmployeeDetailsTextarea = document.getElementById('bulkEmployeeDetails');
    const deleteEmployeeCodeInput = document.getElementById('deleteEmployeeCode');


    // --- Data Groupings for Prospect Details Pop-up ---
    const GENERAL_DETAILS_HEADERS = [
        HEADER_TIMESTAMP, HEADER_DATE, HEADER_BRANCH_NAME, HEADER_EMPLOYEE_NAME, HEADER_EMPLOYEE_CODE, HEADER_DESIGNATION
    ];
    const CONTACT_INFO_HEADERS = [
        HEADER_PROSPECT_NAME, HEADER_PHONE_NUMBER_WHATSAPP, HEADER_ADDRESS, HEADER_PROFESSION
    ];
    const ACTIVITY_DETAILS_HEADERS = [
        HEADER_ACTIVITY_TYPE, HEADER_TYPE_OF_CUSTOMER, HEADER_R_LEAD_SOURCE, HEADER_HOW_CONTACTED, HEADER_PRODUCT_INTERESTED, HEADER_REMARKS, HEADER_NEXT_FOLLOW_UP_DATE
    ];
    const PERSONAL_DETAILS_HEADERS = [
        HEADER_DOB_WD, HEADER_RELATION_WITH_STAFF, HEADER_PROFILE_OF_CUSTOMER
    ];
    const FAMILY_DETAILS_HEADERS = [
        HEADER_FAMILY_DETAILS_1, HEADER_FAMILY_DETAILS_2, HEADER_FAMILY_DETAILS_3, HEADER_FAMILY_DETAILS_4
    ];

    // --- Helper Functions ---

    function showMessage(message, isError = false) {
        statusMessageDiv.textContent = message;
        statusMessageDiv.className = isError ? 'message-container error' : 'message-container success';
        statusMessageDiv.style.display = 'block';
    }

    function clearMessage() {
        statusMessageDiv.textContent = '';
        statusMessageDiv.style.display = 'none';
    }

    // Parses CSV text into an array of objects
    function parseCSV(csvText) {
        const lines = csvText.trim().split('\n');
        if (lines.length === 0) return [];

        // Simple check to see if the first line is likely a header
        const firstLine = lines[0].trim();
        let headers;
        if (firstLine.includes(',')) {
            // Assume comma-separated if it contains commas
            headers = firstLine.split(',').map(header => header.trim().replace(/"/g, ''));
        } else {
            // Otherwise, assume tab-separated (TSV) or single column
            headers = firstLine.split('\t').map(header => header.trim().replace(/"/g, ''));
        }
        
        const data = [];
        for (let i = 1; i < lines.length; i++) {
            const line = lines[i].trim();
            if (!line) continue; // Skip empty lines

            let values;
            if (line.includes(',')) {
                // Split by comma for CSV-like lines
                values = line.split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/).map(value => value.trim().replace(/"/g, ''));
            } else {
                // Split by tab for TSV-like lines
                values = line.split('\t').map(value => value.trim().replace(/"/g, ''));
            }

            const entry = {};
            headers.forEach((header, index) => {
                entry[header] = values[index] || '';
            });
            data.push(entry);
        }
        return data;
    }


    // Helper to calculate total activity from a set of activity entries based on Activity Type
    function calculateTotalActivity(entries) {
        // Initialize counters for all metrics, including the calculated 'New Customer Leads'
        const totalActivity = { 'Visit': 0, 'Calls': 0, 'Referance': 0, 'New Customer Leads': 0 }; 
        
        entries.forEach(entry => {
            const activityType = entry[HEADER_ACTIVITY_TYPE];
            const customerType = entry[HEADER_TYPE_OF_CUSTOMER]; // Get Type of Customer

            switch (activityType) {
                case 'Visit':
                    totalActivity['Visit']++;
                    break;
                case 'Calls':
                    totalActivity['Calls']++;
                    break;
                case 'Referance': // Sticking to the exact spelling 'Referance'
                    totalActivity['Referance']++;
                    break;
                default:
                    // This is the warning you saw. It means HEADER_ACTIVITY_TYPE was undefined or unexpected.
                    console.warn(`Unknown or uncounted Activity Type encountered: ${activityType}`);
            }

            // Logic for 'New Customer Leads' based on your new definition
            if ((activityType === 'Visit' || activityType === 'Calls') && customerType === 'New') {
                totalActivity['New Customer Leads']++;
            }
        });
        return totalActivity;
    }

    // Function to fetch data from a given URL
    async function fetchData(url) {
        try {
            const response = await fetch(url);
            if (!response.ok) {
                throw new Error(`HTTP error! Status: ${response.status}`);
            }
            const text = await response.text();
            return parseCSV(text);
        } catch (error) {
            showMessage(`Failed to fetch data from ${url}. Please check the URL and network connection. Error: ${error.message}`, true);
            console.error('Fetch error:', error);
            return null;
        }
    }

    // Function to send data to Google Apps Script
    async function sendDataToGoogleAppsScript(action, data) {
        clearMessage();
        try {
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                mode: 'cors',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: new URLSearchParams({
                    action: action,
                    data: JSON.stringify(data)
                })
            });

            if (!response.ok) {
                throw new Error(`HTTP error! Status: ${response.status}`);
            }

            const result = await response.json();
            if (result.status === 'SUCCESS') {
                showMessage(result.message || `${action} successful!`);
                // Re-fetch data after successful operation
                await processData(); 
                return true;
            } else {
                showMessage(result.message || `${action} failed.`, true);
                return false;
            }
        } catch (error) {
            showMessage(`Error during ${action}: ${error.message}`, true);
            console.error(`Error during ${action}:`, error);
            return false;
        }
    }

    // Main function to fetch and process all data
    async function processData() {
        showMessage('Loading data...');
        allCanvassingData = await fetchData(DATA_URL);
        // employeeMasterData is now fetched from a URL, no longer hardcoded
        employeeMasterData = await fetchData(EMPLOYEE_MASTER_DATA_URL); 
        clearMessage();

        if (allCanvassingData) {
            // Populate branch and employee dropdowns based on all canvassing data
            branchNames = new Set(allCanvassingData.map(entry => entry[HEADER_BRANCH_NAME]).filter(name => name));
            allEmployees = new Set(allCanvassingData.map(entry => entry[HEADER_EMPLOYEE_NAME]).filter(name => name));
            allEmployeeCodes = new Set(allCanvassingData.map(entry => entry[HEADER_EMPLOYEE_CODE]).filter(code => code));

            populateDropdowns(branchSelect, Array.from(branchNames));
            populateDropdowns(employeeSelect, Array.from(allEmployees)); // For main filters
            populateDropdowns(customerFilterBranchSelect, Array.from(branchNames), true); // For customer filters
            populateDropdowns(customerFilterEmployeeSelect, Array.from(allEmployees), true); // For customer filters

            // Render all reports initially
            renderBranchSummary();
            renderAllStaffOverallPerformance();
            renderNonParticipatingBranches();
            filterAndRenderCustomerDetails(); // Initial render for detailed customer view
        }
    }

    // Populates a given select element with options from an array
    function populateDropdowns(selectElement, items, addAllOption = false) {
        selectElement.innerHTML = ''; // Clear existing options
        if (addAllOption) {
            const allOption = document.createElement('option');
            allOption.value = '';
            allOption.textContent = `-- ${selectElement.id.includes('Branch') ? 'All Branches' : 'All Employees'} --`;
            selectElement.appendChild(allOption);
        } else {
            const defaultOption = document.createElement('option');
            defaultOption.value = '';
            defaultOption.textContent = `-- Select a ${selectElement.id.includes('Branch') ? 'Branch' : 'Employee'} --`;
            selectElement.appendChild(defaultOption);
        }

        items.sort().forEach(item => {
            const option = document.createElement('option');
            option.value = item;
            option.textContent = item;
            selectElement.appendChild(option);
        });
    }


    // --- Report Rendering Functions ---

    function renderBranchSummary() {
        // Calculate overall totals for the summary cards
        const overallTotals = calculateTotalActivity(allCanvassingData);
        document.getElementById('overallTotalVisits').textContent = overallTotals['Visit'];
        document.getElementById('overallTotalCalls').textContent = overallTotals['Calls'];
        document.getElementById('overallTotalReferences').textContent = overallTotals['Referance'];
        document.getElementById('overallTotalNewLeads').textContent = overallTotals['New Customer Leads'];

        // Calculate today's totals
        const today = new Date().toISOString().split('T')[0]; //YYYY-MM-DD
        const todaysData = allCanvassingData.filter(entry => entry[HEADER_DATE] === today);
        const dailyTotals = calculateTotalActivity(todaysData);
        document.getElementById('dailyTotalVisits').textContent = dailyTotals['Visit'];
        document.getElementById('dailyTotalCalls').textContent = dailyTotals['Calls'];
        document.getElementById('dailyTotalReferences').textContent = dailyTotals['Referance'];
        document.getElementById('dailyTotalNewLeads').textContent = dailyTotals['New Customer Leads'];


        branchSummaryTableBody.innerHTML = '';
        const branchActivity = {}; // { 'BranchName': { 'Visit': 0, 'Calls': 0, ... } }

        allCanvassingData.forEach(entry => {
            const branch = entry[HEADER_BRANCH_NAME];
            if (!branchActivity[branch]) {
                branchActivity[branch] = { 'Visit': 0, 'Calls': 0, 'Referance': 0, 'New Customer Leads': 0 };
            }
            // Use calculateTotalActivity for individual branches as well for consistency
            const tempActivity = calculateTotalActivity([entry]); // Pass single entry as array
            for (const key in tempActivity) {
                branchActivity[branch][key] += tempActivity[key];
            }
        });

        // Convert to array for sorting and rendering
        const sortedBranches = Object.keys(branchActivity).sort();

        sortedBranches.forEach(branch => {
            const totals = branchActivity[branch];
            const row = branchSummaryTableBody.insertRow();
            row.insertCell().textContent = branch;
            row.insertCell().textContent = totals['Visit'];
            row.insertCell().textContent = totals['Calls'];
            row.insertCell().textContent = totals['Referance'];
            row.insertCell().textContent = totals['New Customer Leads'];
        });
    }

    function renderAllStaffOverallPerformance() {
        // Removed: allStaffOverallPerformanceTableBody.innerHTML = ''; // This caused the ReferenceError

        const employeePerformance = {}; // { 'EmployeeName': { 'Visit': 0, 'Calls': 0, 'Referance': 0, 'New Customer Leads': 0, 'Branch': '', 'Code': '' } }

        allCanvassingData.forEach(entry => {
            const employeeName = entry[HEADER_EMPLOYEE_NAME];
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            const branchName = entry[HEADER_BRANCH_NAME];

            if (!employeeName) return;

            if (!employeePerformance[employeeName]) {
                employeePerformance[employeeName] = { 
                    'Visit': 0, 'Calls': 0, 'Referance': 0, 'New Customer Leads': 0,
                    'Branch': branchName || 'N/A', // Assign branch, assuming consistent per employee
                    'Code': employeeCode || 'N/A'
                };
            }
            
            const tempActivity = calculateTotalActivity([entry]);
            for (const key in tempActivity) {
                employeePerformance[employeeName][key] += tempActivity[key];
            }
        });

        const staffPerformanceCardsContainer = document.getElementById('staffPerformanceCards');
        staffPerformanceCardsContainer.innerHTML = ''; // Clear existing cards (Correctly targets the div)

        const sortedEmployees = Object.keys(employeePerformance).sort();

        if (sortedEmployees.length === 0) {
            staffPerformanceCardsContainer.innerHTML = '<p class="no-data-message">No employee activity data available to display.</p>';
            return;
        }

        sortedEmployees.forEach(employeeName => {
            const data = employeePerformance[employeeName];
            const card = document.createElement('div');
            card.className = 'employee-performance-card';
            card.innerHTML = `
                <h3>${employeeName} (${data['Code']})</h3>
                <p><strong>Branch:</strong> ${data['Branch']}</p>
                <p>Visits: <span>${data['Visit']}</span></p>
                <p>Calls: <span>${data['Calls']}</span></p>
                <p>References: <span>${data['Referance']}</span></p>
                <p>New Customer Leads: <span>${data['New Customer Leads']}</span></p>
            `;
            staffPerformanceCardsContainer.appendChild(card);
        });
    }

    function renderNonParticipatingBranches() {
        nonParticipatingBranchesTableBody.innerHTML = '';

        const branchEmployeeCounts = {}; // { 'BranchName': { totalEmployees: 0, activeEmployees: Set<'EmployeeCode'> } }

        // Populate total employees per branch from master data
        employeeMasterData.forEach(employee => {
            const branch = employee[HEADER_BRANCH_NAME];
            if (branch) {
                if (!branchEmployeeCounts[branch]) {
                    branchEmployeeCounts[branch] = { totalEmployees: 0, activeEmployees: new Set() };
                }
                branchEmployeeCounts[branch].totalEmployees++;
            }
        });

        // Identify active employees from canvassing data
        allCanvassingData.forEach(entry => {
            const branch = entry[HEADER_BRANCH_NAME];
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            if (branch && employeeCode && branchEmployeeCounts[branch]) {
                branchEmployeeCounts[branch].activeEmployees.add(employeeCode);
            }
        });

        const sortedBranches = Object.keys(branchEmployeeCounts).sort();

        if (sortedBranches.length === 0) {
            nonParticipatingBranchesTableBody.innerHTML = '<tr><td colspan="4" class="no-data-message">No branch data available.</td></tr>';
            return;
        }

        sortedBranches.forEach(branch => {
            const totals = branchEmployeeCounts[branch];
            const nonParticipatingCount = totals.totalEmployees - totals.activeEmployees.size;

            const row = nonParticipatingBranchesTableBody.insertRow();
            row.insertCell().textContent = branch;
            row.insertCell().textContent = totals.totalEmployees;
            row.insertCell().textContent = totals.activeEmployees.size;
            const nonParticipatingCell = row.insertCell();
            nonParticipatingCell.textContent = nonParticipatingCount;
            if (nonParticipatingCount > 0) {
                nonParticipatingCell.classList.add('highlight-red'); // Add a class for styling non-zero counts
            }
        });
    }

    // Renders the detailed customer view table
    function renderCustomerDetailsTable(entries, showOnlyProspectName = false) {
        customerDetailsTableBody.innerHTML = '';
        const tableHeaderRow = customerDetailsTableBody.previousElementSibling.querySelector('thead tr');
        tableHeaderRow.innerHTML = ''; // Clear existing headers

        const headersToShow = showOnlyProspectName ? [HEADER_PROSPECT_NAME] : [
            HEADER_TIMESTAMP, HEADER_DATE, HEADER_BRANCH_NAME, HEADER_EMPLOYEE_NAME, HEADER_ACTIVITY_TYPE,
            HEADER_TYPE_OF_CUSTOMER, HEADER_PROSPECT_NAME, HEADER_PHONE_NUMBER_WHATSAPP, HEADER_PRODUCT_INTERESTED, HEADER_REMARKS
            // You can add more headers here if needed for the full view,
            // but for "show only prospect name", it's just prospect name.
        ];

        headersToShow.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header;
            tableHeaderRow.appendChild(th);
        });

        if (entries.length === 0) {
            const row = customerDetailsTableBody.insertRow();
            const cell = row.insertCell();
            cell.colSpan = headersToShow.length;
            cell.textContent = "No customer entries found for the selected filters.";
            cell.className = 'no-data-message';
            return;
        }

        entries.forEach(entry => {
            const row = customerDetailsTableBody.insertRow();
            headersToShow.forEach(header => {
                const cell = row.insertCell();
                cell.textContent = entry[header] || '';
                if (header === HEADER_PROSPECT_NAME) {
                    cell.classList.add('prospect-name-clickable');
                    cell.dataset.prospectData = JSON.stringify(entry); // Store full entry data
                }
            });
        });
    }

    // Filters and renders the customer details based on dropdown selections
    function filterAndRenderCustomerDetails() {
        const selectedBranch = customerFilterBranchSelect.value;
        const selectedEmployee = customerFilterEmployeeSelect.value;

        let filteredData = allCanvassingData;

        if (selectedBranch) {
            filteredData = filteredData.filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch);
        }

        let showOnlyProspectName = false;
        if (selectedEmployee) {
            filteredData = filteredData.filter(entry => entry[HEADER_EMPLOYEE_NAME] === selectedEmployee);
            showOnlyProspectName = true; // If an employee is selected, show only prospect names
        }

        renderCustomerDetailsTable(filteredData, showOnlyProspectName);
    }

    // Function to display the prospect details in a pop-up modal
    function showProspectDetailsPopup(prospectData) {
        modalBodyContent.innerHTML = ''; // Clear previous content

        // Title
        const title = document.createElement('h3');
        title.textContent = `Details for ${prospectData[HEADER_PROSPECT_NAME] || 'N/A'}`;
        modalBodyContent.appendChild(title);

        const createDetailSection = (sectionTitle, headers) => {
            const sectionDiv = document.createElement('div');
            sectionDiv.className = 'detail-section';

            const sectionHeading = document.createElement('h4');
            sectionHeading.textContent = sectionTitle;
            sectionDiv.appendChild(sectionHeading);

            headers.forEach(header => {
                const value = prospectData[header] || 'N/A';
                const rowDiv = document.createElement('div');
                rowDiv.className = 'detail-row';

                const labelSpan = document.createElement('span');
                labelSpan.className = 'detail-label';
                labelSpan.textContent = `${header}:`;
                rowDiv.appendChild(labelSpan);

                const valueSpan = document.createElement('span');
                valueSpan.className = 'detail-value';
                valueSpan.textContent = value;
                rowDiv.appendChild(valueSpan);

                sectionDiv.appendChild(rowDiv);
            });
            modalBodyContent.appendChild(sectionDiv);
        };

        createDetailSection('General Details', GENERAL_DETAILS_HEADERS);
        createDetailSection('Contact Information', CONTACT_INFO_HEADERS);
        createDetailSection('Activity Details', ACTIVITY_DETAILS_HEADERS);
        createDetailSection('Personal Details', PERSONAL_DETAILS_HEADERS);
        createDetailSection('Family Details', FAMILY_DETAILS_HEADERS);
        
        customerDetailsModal.style.display = 'block'; // Show the modal
    }


    // --- Tab Switching Logic ---
    function showTab(tabButtonId) {
        document.querySelectorAll('.report-section').forEach(section => {
            section.classList.remove('active-tab');
        });
        document.querySelectorAll('.tab-button').forEach(button => {
            button.classList.remove('active');
        });

        // Determine which section to show based on the clicked button
        let targetSection;
        switch (tabButtonId) {
            case 'allBranchSnapshotTabBtn':
                targetSection = document.getElementById('allBranchSnapshotReport');
                // Re-render summary to ensure up-to-date data if filters were applied elsewhere
                renderBranchSummary(); 
                break;
            case 'allStaffOverallPerformanceTabBtn':
                targetSection = document.getElementById('allStaffOverallPerformanceReport');
                renderAllStaffOverallPerformance();
                break;
            case 'nonParticipatingBranchesTabBtn':
                targetSection = document.getElementById('nonParticipatingBranchesReport');
                renderNonParticipatingBranches();
                break;
            case 'detailedCustomerViewTabBtn':
                targetSection = document.getElementById('detailedCustomerViewSection');
                // Re-render customer details based on current filters
                filterAndRenderCustomerDetails(); 
                break;
            case 'employeeManagementTabBtn':
                targetSection = document.getElementById('employeeManagementSection');
                break;
            default:
                console.warn('Unknown tab button ID:', tabButtonId);
                return;
        }

        if (targetSection) {
            targetSection.classList.add('active-tab');
            document.getElementById(tabButtonId).classList.add('active');
            
            // Adjust visibility of main filters based on tab
            const isMainReportTab = ['allBranchSnapshotTabBtn', 'allStaffOverallPerformanceTabBtn'].includes(tabButtonId);
            branchSelect.parentElement.style.display = isMainReportTab ? 'flex' : 'none';
            employeeFilterPanel.style.display = (isMainReportTab && branchSelect.value) ? 'flex' : 'none';
            document.getElementById('viewOptions').style.display = isMainReportTab ? 'flex' : 'none';
        }
    }

    // --- Event Listeners ---

    // Tab button click listeners
    document.getElementById('allBranchSnapshotTabBtn').addEventListener('click', () => showTab('allBranchSnapshotTabBtn'));
    document.getElementById('allStaffOverallPerformanceTabBtn').addEventListener('click', () => showTab('allStaffOverallPerformanceTabBtn'));
    document.getElementById('nonParticipatingBranchesTabBtn').addEventListener('click', () => showTab('nonParticipatingBranchesTabBtn'));
    document.getElementById('detailedCustomerViewTabBtn').addEventListener('click', () => showTab('detailedCustomerViewTabBtn'));
    document.getElementById('employeeManagementTabBtn').addEventListener('click', () => showTab('employeeManagementTabBtn'));


    // Main filter dropdowns (Branch and Employee)
    branchSelect.addEventListener('change', () => {
        const selectedBranch = branchSelect.value;
        if (selectedBranch) {
            employeeFilterPanel.style.display = 'flex';
            // Filter employees based on selected branch
            const employeesInBranch = allCanvassingData
                                    .filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch)
                                    .map(entry => entry[HEADER_EMPLOYEE_NAME])
                                    .filter(name => name); // Ensure no empty names
            populateDropdowns(employeeSelect, Array.from(new Set(employeesInBranch)));
        } else {
            employeeSelect.value = ''; // Reset employee selection
            employeeFilterPanel.style.display = 'none';
            populateDropdowns(employeeSelect, Array.from(allEmployees)); // Repopulate with all employees
        }
        // Re-render reports based on new filter
        renderBranchSummary();
        renderAllStaffOverallPerformance();
    });

    employeeSelect.addEventListener('change', () => {
        // Re-render reports based on new filter
        renderBranchSummary();
        renderAllStaffOverallPerformance();
    });

    // Customer filter dropdowns
    customerFilterBranchSelect.addEventListener('change', () => {
        const selectedBranch = customerFilterBranchSelect.value;
        // Filter employee dropdown for customer view based on selected branch
        if (selectedBranch) {
            const employeesInBranch = allCanvassingData
                                    .filter(entry => entry[HEADER_BRANCH_NAME] === selectedBranch)
                                    .map(entry => entry[HEADER_EMPLOYEE_NAME])
                                    .filter(name => name);
            populateDropdowns(customerFilterEmployeeSelect, Array.from(new Set(employeesInBranch)), true);
        } else {
            populateDropdowns(customerFilterEmployeeSelect, Array.from(allEmployees), true); // Show all employees if no branch selected
        }
    });

    // Apply Filters button for Detailed Customer View
    customerFilterApplyButton.addEventListener('click', filterAndRenderCustomerDetails);

    // Event delegation for prospect name clicks in Detailed Customer View
    customerDetailsTableBody.addEventListener('click', (event) => {
        const targetCell = event.target.closest('.prospect-name-clickable');
        if (targetCell && targetCell.dataset.prospectData) {
            const prospectData = JSON.parse(targetCell.dataset.prospectData);
            showProspectDetailsPopup(prospectData);
        }
    });

    // Modal Close Button
    closeModalButton.addEventListener('click', () => {
        customerDetailsModal.style.display = 'none';
    });

    // Close modal if user clicks outside of the modal content
    window.addEventListener('click', (event) => {
        if (event.target === customerDetailsModal) {
            customerDetailsModal.style.display = 'none';
        }
    });


    // Event Listener for Add Employee Form
    if (addEmployeeForm) {
        addEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const employeeData = {
                [HEADER_BRANCH_NAME]: employeeBranchNameInput.value.trim(),
                [HEADER_EMPLOYEE_NAME]: employeeNameInput.value.trim(),
                [HEADER_EMPLOYEE_CODE]: employeeCodeInput.value.trim(),
                [HEADER_DESIGNATION]: employeeDesignationInput.value.trim()
            };
            if (!employeeData[HEADER_BRANCH_NAME] || !employeeData[HEADER_EMPLOYEE_NAME] || !employeeData[HEADER_EMPLOYEE_CODE]) {
                displayEmployeeManagementMessage('Branch Name, Employee Name, and Employee Code are required.', true);
                return;
            }
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
            const detailsText = bulkEmployeeDetailsTextarea.value.trim();

            if (!branchName) {
                displayEmployeeManagementMessage('Branch Name for bulk entry is required.', true);
                return;
            }
            if (!detailsText) {
                displayEmployeeManagementMessage('Employee details cannot be empty.', true);
                return;
            }

            const employeeLines = detailsText.split('\n').map(line => line.trim()).filter(line => line);
            const employeesToAdd = [];

            for (const line of employeeLines) {
                const parts = line.split(',').map(part => part.trim());
                if (parts.length >= 2) { // Minimum Name,Code
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

    function displayEmployeeManagementMessage(message, isError = false) {
        employeeManagementMessage.textContent = message;
        employeeManagementMessage.className = isError ? 'message-container error' : 'message-container success';
        employeeManagementMessage.style.display = 'block';
        setTimeout(() => {
            employeeManagementMessage.style.display = 'none';
        }, 5000); // Hide after 5 seconds
    }


    // Initial data fetch and tab display when the page loads
    processData();
    showTab('allBranchSnapshotTabBtn');
});
