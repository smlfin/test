// NEW: Report Menu Elements
const reportsMenuBtn = document.getElementById('reportsMenuBtn');
const reportsDropdown = document.getElementById('reportsDropdown');
const allBranchSnapshotBtn = document.getElementById('allBranchSnapshotBtn');

// NEW: Event listener for Reports dropdown menu
if (reportsMenuBtn) {
    reportsMenuBtn.addEventListener('click', function() {
        reportsDropdown.classList.toggle('show');
    });
}

// Close the dropdown if the user clicks outside of it
window.addEventListener('click', function(event) {
    if (!event.target.matches('.dropbtn')) {
        if (reportsDropdown && reportsDropdown.classList.contains('show')) {
            reportsDropdown.classList.remove('show');
        }
    }
});

// NEW: All Branch Snapshot functionality (example data and rendering)
if (allBranchSnapshotBtn) {
    allBranchSnapshotBtn.addEventListener('click', function(event) {
        event.preventDefault(); // Prevent default link behavior
        renderAllBranchSnapshot();
        // Hide other report sections if visible
        reportDisplay.innerHTML = ''; // Clear existing content in main report area
    });
}

function renderAllBranchSnapshot() {
    const reportDisplay = document.getElementById('reportDisplay');
    if (!reportDisplay) {
        console.error("reportDisplay element not found.");
        return;
    }

    // Example data for all branches
    // In a real application, you'd fetch this from your Google Sheet or a backend
    const allBranchData = [
        { branch: "Branch A", totalEmployees: 10, totalCustomersVisited: 500, averageVisitsPerEmployee: 50 },
        { branch: "Branch B", totalEmployees: 8, totalCustomersVisited: 400, averageVisitsPerEmployee: 50 },
        { branch: "Branch C", totalEmployees: 12, totalCustomersVisited: 720, averageVisitsPerEmployee: 60 },
        { branch: "Branch D", totalEmployees: 7, totalCustomersVisited: 350, averageVisitsPerEmployee: 50 }
    ];

    let html = `
        <div id="allBranchSnapshotReport">
            <h2>All Branch Snapshot</h2>
            <table>
                <thead>
                    <tr>
                        <th>Branch</th>
                        <th>Total Employees</th>
                        <th>Total Customers Visited</th>
                        <th>Average Visits/Employee</th>
                    </tr>
                </thead>
                <tbody>
    `;

    allBranchData.forEach(data => {
        html += `
            <tr>
                <td>${data.branch}</td>
                <td>${data.totalEmployees}</td>
                <td>${data.totalCustomersVisited}</td>
                <td>${data.averageVisitsPerEmployee}</td>
            </tr>
        `;
    });

    html += `
                </tbody>
            </table>
        </div>
    `;

    reportDisplay.innerHTML = html;
}