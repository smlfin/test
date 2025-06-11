document.addEventListener('DOMContentLoaded', () => {
    // *** Configuration ***
    // This URL is for your Canvassing Data sheet. Ensure it's correct and published as CSV.
    // NOTE: If you are still getting 404, this URL is the problem.
    const DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=696550092&single=true&output=csv"; 

    // IMPORTANT: Replace this with YOUR DEPLOYED GOOGLE APPS SCRIPT WEB APP URL
    // NOTE: If you are getting errors sending data, this URL is the problem.
    const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbzEYf0CKgwP0O4-z1lup1lDZImD1dQVEveLWsHwa_7T5ltndfIuRWXVZqFDj03_proD/exec"; // <-- PASTE YOUR NEWLY DEPLOYED WEB APP URL HERE

    // We will IGNORE MasterEmployees sheet for data fetching and report generation
    const EMPLOYEE_MASTER_DATA_URL = "UNUSED"; // This remains unused for front-end reporting

    // NEW: Performance optimization for detailed entry table
    const MAX_DISPLAY_ROWS = 500; // Limit the number of rows to display in detailed entries table

    // *** Header Definitions (Updated for 28 columns based on your sheet structure) ***
    // MAKE SURE THESE MATCH YOUR GOOGLE SHEET'S FIRST ROW HEADERS EXACTLY
    const HEADER_TIMESTAMP = "Timestamp"; // MM/DD/YYYY HH:MM:SS format
    const HEADER_DATE = "Date of Visit"; // DD/MM/YYYY format
    const HEADER_DATE_OF_ENTRY = "Date of Entry";
    const HEADER_BRANCH_NAME = "Branch Name";
    const HEADER_EMPLOYEE_NAME = "Employee Name";
    const HEADER_EMPLOYEE_CODE = "Employee Code";
    const HEADER_DESIGNATION = "Designation";
    const HEADER_MODE_OF_CONTACT = "Mode of contact";
    const HEADER_CUSTOMER_TYPE = "Customer Type";
    const HEADER_SOURCE = "Source";
    const HEADER_REASON_FOR_VISIT = "Reason for Visit";
    const HEADER_CUSTOMER_NAME = "Name of Customer";
    const HEADER_PHONE_NUMBER = "Phone Number";
    const HEADER_ADDRESS = "Address";
    const HEADER_WIFE_HUSBAND_NAME = "Name of wife/Husband";
    const HEADER_FAMILY_DETAILS_2 = "Family Deatils -2"; // Adjust if this is a real name
    const HEADER_WIFE_HUSBAND_JOB = "Job of wife/Husband";
    const HEADER_FAMILY_DETAILS_3 = "Family Deatils -3"; // Adjust if this is a real name
    const HEADER_CHILDREN_NAMES = "Names of Children";
    const HEADER_FAMILY_DETAILS_4 = "Family Deatils -4"; // Adjust if this is a real name
    const HEADER_CHILDREN_DETAILS = "Deatils of Children";
    const HEADER_LOAN_PRODUCT = "Product";
    const HEADER_LOAN_STATUS = "Status"; // Under Loan Details
    const HEADER_LOAN_DATE = "Date";    // Under Loan Details
    const HEADER_REMARKS = "Remarks";
    const HEADER_CALLBACK_DATE = "Date to call back";
    const HEADER_DONE_BY = "Done by";
    const HEADER_CUSTOMER_PROFILE_STATUS = "Status"; // Under Customer Profile
    const HEADER_CUSTOMER_PROFILE_INTEREST = "Interested"; // Under Customer Profile


    // Define which headers to display in the detailed table
    const HEADERS_TO_SHOW = [
        HEADER_DATE, HEADER_BRANCH_NAME, HEADER_EMPLOYEE_NAME, HEADER_DESIGNATION,
        HEADER_MODE_OF_CONTACT, HEADER_CUSTOMER_TYPE, HEADER_SOURCE, HEADER_REASON_FOR_VISIT,
        HEADER_CUSTOMER_NAME, HEADER_PHONE_NUMBER, HEADER_ADDRESS, HEADER_WIFE_HUSBAND_NAME,
        HEADER_WIFE_HUSBAND_JOB, HEADER_CHILDREN_NAMES, HEADER_CHILDREN_DETAILS,
        HEADER_LOAN_PRODUCT, HEADER_LOAN_STATUS, HEADER_LOAN_DATE, HEADER_REMARKS,
        HEADER_CALLBACK_DATE, HEADER_DONE_BY, HEADER_CUSTOMER_PROFILE_STATUS,
        HEADER_CUSTOMER_PROFILE_INTEREST
        // NOTE: HEADER_FAMILY_DETAILS_2, HEADER_FAMILY_DETAILS_3, HEADER_FAMILY_DETAILS_4 are included as constants
        // but not in HEADERS_TO_SHOW by default. Add if you want them in the detailed table.
    ];

    // *** DOM Elements ***
    const allBranchSnapshotTableBody = document.getElementById('allBranchSnapshotTableBody');
    const allStaffOverallPerformanceTableBody = document.getElementById('allStaffOverallPerformanceTableBody');
    const nonParticipatingBranchesList = document.getElementById('nonParticipatingBranchesList');
    const employeeManagementMessage = document.getElementById('employeeManagementMessage');
    const bulkAddEmployeeForm = document.getElementById('bulkAddEmployeeForm');
    const bulkEmployeeBranchNameInput = document.getElementById('bulkEmployeeBranchName');
    const bulkEmployeeDetailsInput = document.getElementById('bulkEmployeeDetails');
    const deleteEmployeeForm = document.getElementById('deleteEmployeeForm');
    const deleteEmployeeCodeInput = document.getElementById('deleteEmployeeCode');

    // Filter controls
    const branchSelect = document.getElementById('branchSelect');
    const employeeFilterPanel = document.getElementById('employeeFilterPanel');
    const employeeSelect = document.getElementById('employeeSelect');
    const customerTypeSelect = document.getElementById('customerTypeSelect');
    const reasonForVisitSelect = document.getElementById('reasonForVisitSelect');
    const loanProductSelect = document.getElementById('loanProductSelect');
    const dateRangePicker = document.getElementById('dateRangePicker');

    // Summary Breakdown
    const totalVisitsCount = document.getElementById('totalVisitsCount');
    const totalNewCustomersCount = document.getElementById('totalNewCustomersCount');
    const totalLoanEnquiriesCount = document.getElementById('totalLoanEnquiriesCount');
    const totalRDCustomersCount = document.getElementById('totalRDCustomersCount');
    const totalMISCCustomersCount = document.getElementById('totalMISCCustomersCount');
    const totalNonParticipatingBranchesCount = document.getElementById('totalNonParticipatingBranchesCount');
    const totalParticipatingBranchesCount = document.getElementById('totalParticipatingBranchesCount');


    // Detailed entries table
    const employeeDetailedEntriesTableBody = document.getElementById('employeeDetailedEntriesTableBody');
    const employeeDetailedEntriesSummary = document.getElementById('employeeDetailedEntriesSummary');


    // NEW: Message container for temporary status messages
    const statusMessageContainer = document.getElementById('statusMessage');

    // Global variables to store fetched data
    let allData = []; // Stores all parsed data from CSV
    let branches = new Set(); // Stores unique branch names
    let employees = new Map(); // Stores unique employee codes and names {code: name}
    let uniqueCustomerTypes = new Set();
    let uniqueReasonForVisits = new Set();
    let uniqueLoanProducts = new Set();

    // Store branch participation data
    let branchParticipationData = {}; // { branchName: { totalEmployees: X, participatingEmployees: Y, currentMonthVisits: Z } }


    // *** Helper Functions ***

    // Display temporary status messages
    function displayStatusMessage(message, isError = false) {
        statusMessageContainer.textContent = message;
        statusMessageContainer.className = `message-container ${isError ? 'error-message' : 'success-message'}`;
        statusMessageContainer.style.display = 'block';

        // Keep status messages for longer if it's an error or loading
        if (!isError && !message.includes('Loading')) {
            setTimeout(() => {
                statusMessageContainer.style.display = 'none';
                statusMessageContainer.textContent = '';
            }, 5000); // Message disappears after 5 seconds
        }
    }


    // Robust date parsing function for DD/MM/YYYY (e.g., "11/06/2025")
    // Use this for columns like 'Date of Visit' if they are strictly DD/MM/YYYY.
    function parseDateDDMMYYYY(dateString) {
        if (!dateString) return null;
        const parts = dateString.split(/[-\/\s:]+/);

        if (parts.length < 3) return null;

        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10) - 1; // Month is 0-indexed in JS
        const year = parseInt(parts[2], 10);

        let hour = 0;
        let minute = 0;
        let second = 0;

        if (parts.length >= 6) { // Handles DD/MM/YYYY HH:MM:SS
            hour = parseInt(parts[3], 10);
            minute = parseInt(parts[4], 10);
            second = parseInt(parts[5], 10);
        } else if (parts.length >= 5) { // Handles DD/MM/YYYY HH:MM
            hour = parseInt(parts[3], 10);
            minute = parseInt(parts[4], 10);
        } else if (parts.length >= 4) { // Handles DD/MM/YYYY HH
            hour = parseInt(parts[3], 10);
        }

        const date = new Date(year, month, day, hour, minute, second);
        // Basic validation to prevent invalid dates (e.g., Feb 30th)
        if (isNaN(date.getTime()) || date.getDate() !== day || date.getMonth() !== month || date.getFullYear() !== year) {
            return null;
        }
        return date;
    }

    // NEW: Robust date parsing function for MM/DD/YYYY HH:MM:SS (e.g., "6/11/2025 19:55:18")
    // Use this specifically for your 'Timestamp' column.
    function parseDateMMDDYYYY(dateString) {
        if (!dateString) return null;
        const parts = dateString.split(/[-\/\s:]+/); // Split by common date/time delimiters

        if (parts.length < 3) return null;

        const month = parseInt(parts[0], 10) - 1; // Month is 0-indexed in JS
        const day = parseInt(parts[1], 10);
        const year = parseInt(parts[2], 10);

        let hour = 0;
        let minute = 0;
        let second = 0;

        if (parts.length >= 6) { // Handles MM/DD/YYYY HH:MM:SS
            hour = parseInt(parts[3], 10);
            minute = parseInt(parts[4], 10);
            second = parseInt(parts[5], 10);
        } else if (parts.length >= 5) { // Handles MM/DD/YYYY HH:MM
            hour = parseInt(parts[3], 10);
            minute = parseInt(parts[4], 10);
        } else if (parts.length >= 4) { // Handles MM/DD/YYYY HH
            hour = parseInt(parts[3], 10);
        }

        const date = new Date(year, month, day, hour, minute, second);
        // Basic validation
        if (isNaN(date.getTime()) || date.getDate() !== day || date.getMonth() !== month || date.getFullYear() !== year) {
            return null;
        }
        return date;
    }

    // NEW: Flexible date parsing function for HEADER_TIMESTAMP
    // This will attempt to parse as MM/DD/YYYY HH:MM:SS first, then DD/MM/YYYY
    function parseFlexibleTimestampDate(dateString) {
        if (!dateString) return null;

        // Attempt to parse as MM/DD/YYYY HH:MM:SS (original Timestamp format)
        let parsedDate = parseDateMMDDYYYY(dateString);
        if (parsedDate && !isNaN(parsedDate.getTime())) {
            return parsedDate;
        }

        // If MM/DD/YYYY HH:MM:SS failed, attempt to parse as DD/MM/YYYY
        parsedDate = parseDateDDMMYYYY(dateString);
        if (parsedDate && !isNaN(parsedDate.getTime())) {
            return parsedDate;
        }

        return null; // Could not parse with either format
    }


    // Utility to format date to ISO-MM-DD
    // MODIFIED to use robust parseDateDDMMYYYY for initial date parsing
    const formatDate = (dateString) => {
        if (!dateString) return '';
        // Assuming this function is primarily used for HEADER_DATE (DD/MM/YYYY)
        const date = parseDateDDMMYYYY(dateString); // <--- MODIFIED HERE
        if (!date) return dateString; // If parsing failed, return original string or empty
        return date.toISOString().split('T')[0];
    };

    // Generic CSV parsing function
    function parseCSV(csvText) {
        const lines = csvText.trim().split('\n');
        if (lines.length === 0) return [];

        // IMPORTANT CHANGE 1: Get headers from the FIRST line (index 0)
        // This assumes your Google Sheet now exports a clean, single-row header CSV.
        const headers = parseCSVLine(lines[0]); // <--- MODIFIED LINE

        const data = [];

        // IMPORTANT CHANGE 2: Start processing data from the SECOND line (index 1)
        // This corresponds to Row 2 in your Google Sheet (your actual data entries)
        for (let i = 1; i < lines.length; i++) { // <--- MODIFIED LINE
            const line = lines[i].trim();
            if (!line) continue; // Skip empty lines

            const values = parseCSVLine(line);

            if (values.length !== headers.length) {
                console.warn(`Skipping malformed row ${i + 1}: Expected ${headers.length} columns, got ${values.length}. Line: "${line}"`);
                continue; // Skip malformed rows
            }

            const entry = {};
            headers.forEach((header, index) => {
                entry[header.trim()] = values[index] ? values[index].trim() : '';
            });
            data.push(entry);
        }
        return data;
    }

    // Helper to parse a single CSV line, handling commas within quotes
    function parseCSVLine(line) {
        const regex = /(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|([^,]*))/g;
        const values = [];
        let match;
        while ((match = regex.exec(line)) !== null) {
            let value = match[1] !== undefined ? match[1].replace(/\"\"/g, '\"') : match[2];
            values.push(value);
        }
        // Handle trailing empty fields if the line ends with a comma
        if (line.endsWith(',') && (match === null || regex.lastIndex === line.length)) {
             values.push('');
        }
        return values;
    }


    // Fetch data from Google Sheet CSV
    async function fetchData() {
        displayStatusMessage('Loading data...', false); // Show loading message
        try {
            const response = await fetch(DATA_URL);
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            const csvText = await response.text();
            allData = parseCSV(csvText);
            displayStatusMessage('Data loaded successfully!', false);
            return allData;
        } catch (error) {
            console.error("Error fetching or parsing data:", error);
            displayStatusMessage(`Failed to load data: ${error.message}. Please check DATA_URL and network.`, true);
            return [];
        }
    }

    // Initialize filters with unique values from data
    function initializeFilters(data) {
        branches.clear();
        employees.clear(); // Clear existing map
        uniqueCustomerTypes.clear();
        uniqueReasonForVisits.clear();
        uniqueLoanProducts.clear();

        data.forEach(entry => {
            if (entry[HEADER_BRANCH_NAME]) {
                branches.add(entry[HEADER_BRANCH_NAME]);
            }
            if (entry[HEADER_EMPLOYEE_CODE] && entry[HEADER_EMPLOYEE_NAME]) {
                employees.set(entry[HEADER_EMPLOYEE_CODE], entry[HEADER_EMPLOYEE_NAME]);
            }
            if (entry[HEADER_CUSTOMER_TYPE]) {
                uniqueCustomerTypes.add(entry[HEADER_CUSTOMER_TYPE]);
            }
            if (entry[HEADER_REASON_FOR_VISIT]) {
                uniqueReasonForVisits.add(entry[HEADER_REASON_FOR_VISIT]);
            }
            if (entry[HEADER_LOAN_PRODUCT]) {
                uniqueLoanProducts.add(entry[HEADER_LOAN_PRODUCT]);
            }
        });

        // Populate Branch Select
        branchSelect.innerHTML = '<option value="">-- Select a Branch --</option>';
        Array.from(branches).sort().forEach(branch => {
            const option = document.createElement('option');
            option.value = branch;
            option.textContent = branch;
            branchSelect.appendChild(option);
        });

        // Populate Employee Select (initially hidden)
        employeeSelect.innerHTML = '<option value="">-- Select an Employee --</option>';
        Array.from(employees.entries()).sort((a,b) => a[1].localeCompare(b[1])).forEach(([code, name]) => {
            const option = document.createElement('option');
            option.value = code; // Use code as value
            option.textContent = `${name} (${code})`; // Display name and code
            employeeSelect.appendChild(option);
        });

        // Populate Customer Type Select
        customerTypeSelect.innerHTML = '<option value="">-- All --</option>';
        Array.from(uniqueCustomerTypes).sort().forEach(type => {
            const option = document.createElement('option');
            option.value = type;
            option.textContent = type;
            customerTypeSelect.appendChild(option);
        });

        // Populate Reason for Visit Select
        reasonForVisitSelect.innerHTML = '<option value="">-- All --</option>';
        Array.from(uniqueReasonForVisits).sort().forEach(reason => {
            const option = document.createElement('option');
            option.value = reason;
            option.textContent = reason;
            reasonForVisitSelect.appendChild(option);
        });

        // Populate Loan Product Select
        loanProductSelect.innerHTML = '<option value="">-- All --</option>';
        Array.from(uniqueLoanProducts).sort().forEach(product => {
            const option = document.createElement('option');
            option.value = product;
            option.textContent = product;
            loanProductSelect.appendChild(option);
        });
    }

    // Filter functions
    const filterDataByBranch = (data, branch) => branch ? data.filter(entry => entry[HEADER_BRANCH_NAME] === branch) : data;
    const filterDataByEmployee = (data, employeeCode) => employeeCode ? data.filter(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode) : data;
    const filterDataByCustomerType = (data, customerType) => customerType ? data.filter(entry => entry[HEADER_CUSTOMER_TYPE] === customerType) : data;
    const filterDataByReasonForVisit = (data, reason) => reason ? data.filter(entry => entry[HEADER_REASON_FOR_VISIT] === reason) : data;
    const filterDataByLoanProduct = (data, product) => product ? data.filter(entry => entry[HEADER_LOAN_PRODUCT] === product) : data;

    const filterDataByDateRange = (data, dateRange) => {
        if (!dateRange || dateRange === "All Time") {
            return data;
        }

        const today = new Date();
        today.setHours(0, 0, 0, 0); // Set to start of today

        let startDate = null;
        let endDate = null;

        switch (dateRange) {
            case "Today":
                startDate = new Date(today);
                endDate = new Date(today);
                endDate.setHours(23, 59, 59, 999); // Set to end of today
                break;
            case "This Week": // Monday to Sunday of the current week
                const dayOfWeek = today.getDay(); // 0 for Sunday, 1 for Monday
                startDate = new Date(today);
                startDate.setDate(today.getDate() - (dayOfWeek === 0 ? 6 : dayOfWeek - 1)); // Go to Monday
                startDate.setHours(0, 0, 0, 0);
                endDate = new Date(startDate);
                endDate.setDate(startDate.getDate() + 6); // Go to Sunday
                endDate.setHours(23, 59, 59, 999);
                break;
            case "This Month":
                startDate = new Date(today.getFullYear(), today.getMonth(), 1);
                endDate = new Date(today.getFullYear(), today.getMonth() + 1, 0); // Last day of month
                endDate.setHours(23, 59, 59, 999);
                break;
            case "This Quarter": // Current quarter based on months
                const currentMonth = today.getMonth();
                const quarterStartMonth = Math.floor(currentMonth / 3) * 3;
                startDate = new Date(today.getFullYear(), quarterStartMonth, 1);
                endDate = new Date(today.getFullYear(), quarterStartMonth + 3, 0);
                endDate.setHours(23, 59, 59, 999);
                break;
            case "This Year":
                startDate = new Date(today.getFullYear(), 0, 1);
                endDate = new Date(today.getFullYear(), 11, 31);
                endDate.setHours(23, 59, 59, 999);
                break;
            case "Last 7 Days":
                startDate = new Date(today);
                startDate.setDate(today.getDate() - 6); // Today - 6 days = last 7 days including today
                startDate.setHours(0, 0, 0, 0);
                endDate = new Date(today);
                endDate.setHours(23, 59, 59, 999);
                break;
            case "Last 30 Days":
                startDate = new Date(today);
                startDate.setDate(today.getDate() - 29); // Today - 29 days = last 30 days including today
                startDate.setHours(0, 0, 0, 0);
                endDate = new Date(today);
                endDate.setHours(23, 59, 59, 999);
                break;
            case "Yesterday":
                startDate = new Date(today);
                startDate.setDate(today.getDate() - 1);
                startDate.setHours(0, 0, 0, 0);
                endDate = new Date(today);
                endDate.setDate(today.getDate() - 1);
                endDate.setHours(23, 59, 59, 999);
                break;
            default: // Handle custom date range if implemented
                if (dateRange.includes(' to ')) {
                    const [startStr, endStr] = dateRange.split(' to ');
                    startDate = new Date(startStr); // Assumes YYYY-MM-DD for custom range input
                    endDate = new Date(endStr);
                    endDate.setHours(23, 59, 59, 999);
                } else {
                    return data;
                }
        }

        if (!startDate || !endDate) {
            return data;
        }

        return data.filter(entry => {
            const entryDate = parseFlexibleTimestampDate(entry[HEADER_TIMESTAMP]);

            if (!entryDate) return false; // Skip if date cannot be parsed

            return entryDate >= startDate && entryDate <= endDate;
        });
    };


    // Main rendering functions
    function renderAllBranchSnapshot(filteredData) {
        allBranchSnapshotTableBody.innerHTML = '';
        const branchData = {}; // { 'Branch Name': { totalVisits: N, newCustomers: M, participatingEmployees: Set } }

        // Get all unique employees for initial participation count
        const allEmployees = new Map(); // employeeCode: branchName
        Array.from(branches).forEach(branch => {
            branchData[branch] = {
                totalVisits: 0,
                newCustomers: 0,
                participatingEmployees: new Set()
            };
        });

        // Fill branchData with counts from filtered data
        filteredData.forEach(entry => {
            const branch = entry[HEADER_BRANCH_NAME];
            const employeeCode = entry[HEADER_EMPLOYEE_CODE];
            const customerType = entry[HEADER_CUSTOMER_TYPE];

            if (branchData[branch]) {
                branchData[branch].totalVisits++;
                if (customerType === 'New') { // Assuming 'New' is the value for new customers
                    branchData[branch].newCustomers++;
                }
                if (employeeCode) {
                    branchData[branch].participatingEmployees.add(employeeCode);
                }
            }
        });

        // Populate table rows
        let grandTotalVisits = 0;
        let grandTotalNewCustomers = 0;
        let totalParticipatingBranches = 0;

        Array.from(branches).sort().forEach(branch => {
            const data = branchData[branch] || { totalVisits: 0, newCustomers: 0, participatingEmployees: new Set() };
            const row = allBranchSnapshotTableBody.insertRow();
            row.insertCell().textContent = branch;
            row.insertCell().textContent = data.totalVisits;
            row.insertCell().textContent = data.newCustomers;
            row.insertCell().textContent = data.participatingEmployees.size; // Count unique participating employees

            grandTotalVisits += data.totalVisits;
            grandTotalNewCustomers += data.newCustomers;
            if (data.totalVisits > 0) { // Consider a branch participating if it has at least one visit
                totalParticipatingBranches++;
            }
        });

        // Add totals row
        const totalsRow = allBranchSnapshotTableBody.insertRow();
        totalsRow.classList.add('totals-row');
        totalsRow.insertCell().textContent = 'Grand Total';
        totalsRow.insertCell().textContent = grandTotalVisits;
        totalsRow.insertCell().textContent = grandTotalNewCustomers;
        totalsRow.insertCell().textContent = '---'; // No sum for participating employees directly here
    }

    function renderAllStaffOverallPerformance(filteredData) {
        allStaffOverallPerformanceTableBody.innerHTML = '';
        const employeePerformance = {}; // { 'EMP001': { name: 'Alice', visits: N, newCustomers: M } }

        // Initialize all known employees (from allData) in the performance object
        employees.forEach((name, code) => {
            employeePerformance[code] = {
                name: name,
                visits: 0,
                newCustomers: 0,
                loanEnquiries: 0, // Assuming 'Loan Enquiry' status for loan products
                rdCustomers: 0, // Assuming 'RD' product for RD customers
                miscCustomers: 0 // Assuming 'MISC' for miscellaneous customers
            };
        });

        // Aggregate data based on filtered entries
        filteredData.forEach(entry => {
            const code = entry[HEADER_EMPLOYEE_CODE];
            const customerType = entry[HEADER_CUSTOMER_TYPE];
            const loanProduct = entry[HEADER_LOAN_PRODUCT];

            if (employeePerformance[code]) {
                employeePerformance[code].visits++;
                if (customerType === 'New') {
                    employeePerformance[code].newCustomers++;
                }
                if (loanProduct === 'Loan') { // Adjust this condition based on your actual loan product names
                    employeePerformance[code].loanEnquiries++;
                }
                if (loanProduct === 'RD') { // Adjust this condition for RD product name
                    employeePerformance[code].rdCustomers++;
                }
                if (loanProduct === 'MISC') { // Adjust this condition for MISC product name
                    employeePerformance[code].miscCustomers++;
                }
            }
        });

        // Populate table rows
        let grandVisits = 0;
        let grandNewCustomers = 0;
        let grandLoanEnquiries = 0;
        let grandRDCustomers = 0;
        let grandMISCCustomers = 0;

        // Sort employees by name for consistent display
        const sortedEmployeeCodes = Array.from(employees.keys()).sort((codeA, codeB) => {
            const nameA = employees.get(codeA) || '';
            const nameB = employees.get(codeB) || '';
            return nameA.localeCompare(nameB);
        });

        sortedEmployeeCodes.forEach(code => {
            const data = employeePerformance[code]; // Will be initialized even if no entries
            const row = allStaffOverallPerformanceTableBody.insertRow();
            row.insertCell().textContent = data.name;
            row.insertCell().textContent = code;
            row.insertCell().textContent = data.visits;
            row.insertCell().textContent = data.newCustomers;
            row.insertCell().textContent = data.loanEnquiries;
            row.insertCell().textContent = data.rdCustomers;
            row.insertCell().textContent = data.miscCustomers;

            grandVisits += data.visits;
            grandNewCustomers += data.newCustomers;
            grandLoanEnquiries += data.loanEnquiries;
            grandRDCustomers += data.rdCustomers;
            grandMISCCustomers += data.miscCustomers;
        });

        // Add totals row
        const totalsRow = allStaffOverallPerformanceTableBody.insertRow();
        totalsRow.classList.add('totals-row');
        totalsRow.insertCell().textContent = 'Grand Total';
        totalsRow.insertCell().textContent = '---'; // No code for total
        totalsRow.insertCell().textContent = grandVisits;
        totalsRow.insertCell().textContent = grandNewCustomers;
        totalsRow.insertCell().textContent = grandLoanEnquiries;
        totalsRow.insertCell().textContent = grandRDCustomers;
        totalsRow.insertCell().textContent = grandMISCCustomers;
    }

    function renderNonParticipatingBranches() {
        nonParticipatingBranchesList.innerHTML = '';
        const nonParticipating = Array.from(branches).filter(branch => {
            const branchVisits = allData.filter(entry => entry[HEADER_BRANCH_NAME] === branch);
            return branchVisits.length === 0;
        }).sort();

        if (nonParticipating.length === 0) {
            nonParticipatingBranchesList.innerHTML = '<p class="no-participation-message">All branches have participated!</p>';
        } else {
            nonParticipating.forEach(branch => {
                const li = document.createElement('li');
                li.textContent = branch;
                nonParticipatingBranchesList.appendChild(li);
            });
        }
    }


    // Function to calculate and render the summary breakdown
    function renderSummaryBreakdown(filteredData) {
        let totalVisits = filteredData.length;
        let totalNewCustomers = filteredData.filter(entry => entry[HEADER_CUSTOMER_TYPE] === 'New').length;
        let totalLoanEnquiries = filteredData.filter(entry => entry[HEADER_LOAN_PRODUCT] === 'Loan').length; // Adjust product name if needed
        let totalRDCustomers = filteredData.filter(entry => entry[HEADER_LOAN_PRODUCT] === 'RD').length;    // Adjust product name if needed
        let totalMISCCustomers = filteredData.filter(entry => entry[HEADER_LOAN_PRODUCT] === 'MISC').length; // Adjust product name if needed

        // Calculate participating vs non-participating branches based on filtered data
        const participatingBranchesSet = new Set(filteredData.map(entry => entry[HEADER_BRANCH_NAME]));
        let totalParticipatingBranches = participatingBranchesSet.size;
        let totalNonParticipatingBranches = branches.size - totalParticipatingBranches;


        totalVisitsCount.textContent = totalVisits;
        totalNewCustomersCount.textContent = totalNewCustomers;
        totalLoanEnquiriesCount.textContent = totalLoanEnquiries;
        totalRDCustomersCount.textContent = totalRDCustomers;
        totalMISCCustomersCount.textContent = totalMISCCustomers;
        totalParticipatingBranchesCount.textContent = totalParticipatingBranches;
        totalNonParticipatingBranchesCount.textContent = totalNonParticipatingBranches;
    }

    // Function to render detailed entries for a specific employee
    function renderEmployeeDetailedEntries(employeeCode, filteredData) {
        employeeDetailedEntriesTableBody.innerHTML = '';
        employeeDetailedEntriesSummary.innerHTML = ''; // Clear summary area

        const employeeName = employees.get(employeeCode) || 'Unknown Employee';
        const employeeEntries = filteredData.filter(entry => entry[HEADER_EMPLOYEE_CODE] === employeeCode);

        // Sort entries by Date of Visit (HEADER_DATE) in descending order
        employeeEntries.sort((a, b) => {
            const dateA = parseDateDDMMYYYY(a[HEADER_DATE]); // Using DD/MM/YYYY parser for HEADER_DATE
            const dateB = parseDateDDMMYYYY(b[HEADER_DATE]); // Using DD/MM/YYYY parser for HEADER_DATE
            return (dateB ? dateB.getTime() : 0) - (dateA ? dateA.getTime() : 0);
        });

        if (employeeEntries.length === 0) {
            employeeDetailedEntriesTableBody.innerHTML = `<tr><td colspan="${HEADERS_TO_SHOW.length}">No entries found for ${employeeName}.</td></tr>`;
            return;
        }

        // Create table header for detailed entries
        const thead = employeeDetailedEntriesTableBody.createTHead();
        const headerRow = thead.insertRow();
        HEADERS_TO_SHOW.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header;
            headerRow.appendChild(th);
        });

        // Add table body rows, respecting MAX_DISPLAY_ROWS
        const entriesToDisplay = employeeEntries.slice(0, MAX_DISPLAY_ROWS); // Limit display <--- MODIFIED
        entriesToDisplay.forEach(entry => { // Loop over the limited array <--- MODIFIED
            const row = employeeDetailedEntriesTableBody.insertRow();
            HEADERS_TO_SHOW.forEach(header => {
                const cell = row.insertCell();
                let cellValue = entry[header];
                if (header === HEADER_DATE || header === HEADER_CALLBACK_DATE) { // Apply date formatting
                    cellValue = formatDate(cellValue);
                }
                cell.textContent = cellValue;
            });
        });

        // Add a message if not all entries are displayed
        if (employeeEntries.length > MAX_DISPLAY_ROWS) { // <--- ADDED
            const moreRowsMessage = employeeDetailedEntriesTableBody.insertRow();
            const messageCell = moreRowsMessage.insertCell();
            messageCell.colSpan = HEADERS_TO_SHOW.length;
            messageCell.classList.add('info-message'); // Add a class for styling
            messageCell.textContent = `Displaying the most recent ${MAX_DISPLAY_ROWS} entries. Total entries: ${employeeEntries.length}.`;
        }

        // Display summary for the detailed entries
        const employeeTotalVisits = employeeEntries.length;
        const employeeNewCustomers = employeeEntries.filter(e => e[HEADER_CUSTOMER_TYPE] === 'New').length;
        const employeeLoanEnquiries = employeeEntries.filter(e => e[HEADER_LOAN_PRODUCT] === 'Loan').length;
        const employeeRDCustomers = employeeEntries.filter(e => e[HEADER_LOAN_PRODUCT] === 'RD').length;
        const employeeMISCCustomers = employeeEntries.filter(e => e[HEADER_LOAN_PRODUCT] === 'MISC').length;

        employeeDetailedEntriesSummary.innerHTML = `
            <h3>Summary for ${employeeName}</h3>
            <ul class="summary-list">
                <li>Total Visits: <span>${employeeTotalVisits}</span></li>
                <li>New Customers: <span>${employeeNewCustomers}</span></li>
                <li>Loan Enquiries: <span>${employeeLoanEnquiries}</span></li>
                <li>RD Customers: <span>${employeeRDCustomers}</span></li>
                <li>MISC Customers: <span>${employeeMISCCustomers}</span></li>
            </ul>
        `;
    }


    // *** Event Handlers ***

    // Main function to apply all filters and render reports
    function applyFiltersAndRender() {
        displayStatusMessage('Applying filters and rendering reports...', false); // Show message
        let filteredData = allData;

        // Get selected filter values
        const selectedBranch = branchSelect.value;
        const selectedEmployeeCode = employeeSelect.value;
        const selectedCustomerType = customerTypeSelect.value;
        const selectedReasonForVisit = reasonForVisitSelect.value;
        const selectedLoanProduct = loanProductSelect.value;
        const selectedDateRange = dateRangePicker.value;

        // Apply filters in sequence
        filteredData = filterDataByBranch(filteredData, selectedBranch);
        filteredData = filterDataByEmployee(filteredData, selectedEmployeeCode);
        filteredData = filterDataByCustomerType(filteredData, selectedCustomerType);
        filteredData = filterDataByReasonForVisit(filteredData, selectedReasonForVisit);
        filteredData = filterDataByLoanProduct(filteredData, selectedLoanProduct);
        filteredData = filterDataByDateRange(filteredData, selectedDateRange);


        // Render all reports with the filtered data
        renderAllBranchSnapshot(filteredData);
        renderAllStaffOverallPerformance(filteredData);
        renderSummaryBreakdown(filteredData);

        // Only render detailed employee entries if an employee is selected
        if (selectedEmployeeCode) {
            renderEmployeeDetailedEntries(selectedEmployeeCode, filteredData);
        } else {
            employeeDetailedEntriesTableBody.innerHTML = '<tr><td colspan="23">Please select an employee from the filter to see detailed entries.</td></tr>';
            employeeDetailedEntriesSummary.innerHTML = '';
        }
        displayStatusMessage('Reports updated.', false); // Clear message after rendering
    }


    // Event Listeners for Filter Changes
    branchSelect.addEventListener('change', applyFiltersAndRender);
    employeeSelect.addEventListener('change', applyFiltersAndRender);
    customerTypeSelect.addEventListener('change', applyFiltersAndRender);
    reasonForVisitSelect.addEventListener('change', applyFiltersAndRender);
    loanProductSelect.addEventListener('change', applyFiltersAndRender);
    dateRangePicker.addEventListener('change', applyFiltersAndRender);


    // Logic to show/hide employee filter based on tab selection
    const allTabButtons = document.querySelectorAll('.tab-button');
    const sections = {
        'allBranchSnapshotTabBtn': document.getElementById('allBranchSnapshotSection'),
        'allStaffOverallPerformanceTabBtn': document.getElementById('allStaffOverallPerformanceSection'),
        'nonParticipatingBranchesTabBtn': document.getElementById('nonParticipatingBranchesSection'),
        'employeeManagementTabBtn': document.getElementById('employeeManagementSection')
    };

    function showTab(tabId) {
        // Remove 'active' from all buttons and hide all sections
        allTabButtons.forEach(button => button.classList.remove('active'));
        for (const sectionId in sections) {
            sections[sectionId].style.display = 'none';
        }

        // Add 'active' to the clicked button and show its section
        document.getElementById(tabId).classList.add('active');
        sections[tabId].style.display = 'block';

        // Adjust visibility of employee filter panel
        if (tabId === 'allStaffOverallPerformanceTabBtn') {
            employeeFilterPanel.style.display = 'flex'; // Show filter
        } else {
            employeeFilterPanel.style.display = 'none'; // Hide filter
            employeeSelect.value = ''; // Reset employee selection when hiding
        }

        // Special rendering for non-participating branches as it doesn't use filters
        if (tabId === 'nonParticipatingBranchesTabBtn') {
            renderNonParticipatingBranches();
        }

        // Re-apply filters when switching tabs that use them
        if (tabId === 'allBranchSnapshotTabBtn' || tabId === 'allStaffOverallPerformanceTabBtn') {
            applyFiltersAndRender();
        }
    }

    allTabButtons.forEach(button => {
        button.addEventListener('click', () => showTab(button.id));
    });

    // Handle employee management messages
    function displayEmployeeManagementMessage(message, isError = false) {
        employeeManagementMessage.textContent = message;
        employeeManagementMessage.className = `message-container ${isError ? 'error-message' : 'success-message'}`;
        employeeManagementMessage.style.display = 'block';

        setTimeout(() => {
            employeeManagementMessage.style.display = 'none';
            employeeManagementMessage.textContent = '';
        }, 5000); // Message disappears after 5 seconds
    }

    // Function to send data to Google Apps Script Web App
    async function sendDataToGoogleAppsScript(action, data) {
        displayEmployeeManagementMessage('Processing request...', false);
        try {
            const response = await fetch(WEB_APP_URL, {
                method: 'POST',
                mode: 'cors', // Crucial for cross-origin requests
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                },
                body: new URLSearchParams({ action: action, data: JSON.stringify(data) })
            });

            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`Server error: ${response.status} - ${errorText}`);
            }

            const result = await response.json();
            if (result.status === 'SUCCESS') {
                displayEmployeeManagementMessage(result.message, false);
                // After successful operation, re-fetch data to update reports
                await processData(); // Re-fetch and re-render all data
                return true;
            } else {
                displayEmployeeManagementMessage(result.message, true);
                return false;
            }
        } catch (error) {
            console.error('Error sending data to Apps Script:', error);
            displayEmployeeManagementMessage(`Failed to complete action: ${error.message}`, true);
            return false;
        }
    }

    // Master function to fetch, initialize filters, and apply filters
    async function processData() {
        displayStatusMessage('Initializing dashboard...', false); // Initial message for the whole process
        const data = await fetchData();
        if (data.length > 0) {
            initializeFilters(data);
            applyFiltersAndRender(); // Initial render with all data
        } else {
            // Handle case where no data is fetched
            console.warn("No data fetched from the Google Sheet.");
            displayStatusMessage("No data available to display. Please check your Google Sheet and DATA_URL.", true);
            // Clear tables if no data
            allBranchSnapshotTableBody.innerHTML = '<tr><td colspan="4">No data available.</td></tr>';
            allStaffOverallPerformanceTableBody.innerHTML = '<tr><td colspan="7">No data available.</td></tr>';
            employeeDetailedEntriesTableBody.innerHTML = '<tr><td colspan="23">No data available.</td></tr>';
            nonParticipatingBranchesList.innerHTML = '<p class="no-participation-message">No data to determine participation.</p>';
            totalVisitsCount.textContent = '0';
            totalNewCustomersCount.textContent = '0';
            totalLoanEnquiriesCount.textContent = '0';
            totalRDCustomersCount.textContent = '0';
            totalMISCCustomersCount.textContent = '0';
            totalParticipatingBranchesCount.textContent = '0';
            totalNonParticipatingBranchesCount.textContent = '0';
        }
        displayStatusMessage('Dashboard ready.', false); // Final message
    }


    // Event Listener for Bulk Add Employee Form
    if (bulkAddEmployeeForm) {
        bulkAddEmployeeForm.addEventListener('submit', async (event) => {
            event.preventDefault();
            const branchName = bulkEmployeeBranchNameInput.value.trim();
            const employeeDetailsText = bulkEmployeeDetailsInput.value.trim();

            if (!branchName || !employeeDetailsText) {
                displayEmployeeManagementMessage('Branch Name and Employee Details are required.', true);
                return;
            }

            const lines = employeeDetailsText.split('\n').map(line => line.trim()).filter(line => line.length > 0);
            const employeesToAdd = [];

            for (const line of lines) {
                const parts = line.split(',').map(part => part.trim());
                if (parts.length < 2) { // Expect at least Name, Code
                    displayEmployeeManagementMessage(`Skipping malformed employee entry: "${line}". Format should be Name,Code,Designation.`, true);
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
    showTab('allBranchSnapshotTabBtn');
});
