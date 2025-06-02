document.addEventListener('DOMContentLoaded', () => {
    // IMPORTANT: Replace with the actual public URL of your canvassing_data.json file from Google Drive
    const DATA_URL = "https://drive.google.com/uc?id=1mKgU6hdDKP8UauGRC6VM7DTR2nM9sOzX&export=download"; // Your verified DATA_URL

    const reportsContainer = document.getElementById('reportsContainer');
    if (!reportsContainer) {
        console.error("Error: Element with ID 'reportsContainer' not found in index.html. Cannot display reports.");
        // If the container is missing, we can't do anything, so we stop.
        return;
    }
    // Set initial loading message
    reportsContainer.innerHTML = "<p>Loading canvassing data...</p>";

    fetch(DATA_URL)
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.json();
        })
        .then(data => {
            console.log("Fetched raw data:", data); // Log fetched data for debugging

            if (!data || data.length === 0) {
                reportsContainer.innerHTML = "<p>No canvassing data available.</p>";
                return;
            }

            // Group data by Branch Name, then by Employee Name
            const groupedData = {};

            data.forEach(entry => {
                const branchName = entry['Branch Name'];
                const employeeName = entry['Employee Name'];

                if (!branchName) {
                    console.warn("Entry missing 'Branch Name':", entry);
                    return; // Skip entries without a branch name
                }
                if (!employeeName) {
                    console.warn("Entry missing 'Employee Name' for branch:", branchName, entry);
                    // For now, let's put entries without an employee under a generic 'Unassigned'
                    // Or you can skip them, depending on your data quality expectations
                    // For this example, we'll put them under 'Unassigned Employee'
                    employeeName = "Unassigned Employee";
                }

                if (!groupedData[branchName]) {
                    groupedData[branchName] = {}; // Initialize branch
                }
                if (!groupedData[branchName][employeeName]) {
                    groupedData[branchName][employeeName] = []; // Initialize employee for this branch
                }
                groupedData[branchName][employeeName].push(entry); // Add the entry to the employee's array
            });

            console.log("Grouped data:", groupedData); // Log grouped data for debugging

            // Clear loading message
            reportsContainer.innerHTML = '';

            // Render the grouped data
            for (const branch in groupedData) {
                const branchDiv = document.createElement('div');
                branchDiv.className = 'branch-group'; // Add a class for styling
                branchDiv.innerHTML = `<h2>Branch Name: ${branch}</h2>`;

                const employeeList = document.createElement('ul');
                employeeList.className = 'employee-list'; // Add a class for styling

                for (const employee in groupedData[branch]) {
                    const employeeLi = document.createElement('li');
                    employeeLi.className = 'employee-item'; // Add a class for styling
                    employeeLi.innerHTML = `<h3>Employee Name: ${employee}</h3>`;

                    const entryDetailsList = document.createElement('ul');
                    entryDetailsList.className = 'entry-details-list'; // Add a class for styling

                    // Loop through each entry for this employee
                    groupedData[branch][employee].forEach(entry => {
                        const entryLi = document.createElement('li');
                        entryLi.className = 'entry-detail-item'; // Add a class for styling

                        // Dynamically display all relevant fields for each entry
                        let detailHtml = `<h4>Canvassing Entry</h4>`;
                        detailHtml += `<p><strong>Timestamp:</strong> ${entry['Timestamp'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Date:</strong> ${entry['Date'] || 'N/A'}</p>`;
                        // No need for Branch Name/Employee Name again here as they are headers
                        detailHtml += `<p><strong>Employee Code:</strong> ${entry['Employee Code'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Designation:</strong> ${entry['Designation'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Activity Type:</strong> ${entry['Activity Type'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Type of Customer:</strong> ${entry['Type of Customer'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Lead Source:</strong> ${entry['Lead Source'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>How Contacted:</strong> ${entry['How Contacted'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Prospect Name:</strong> ${entry['Prospect Name'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Phone Number (Whatsapp):</strong> ${entry['Phone Numebr(Whatsapp)'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Address:</strong> ${entry['Address'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Profession:</strong> ${entry['Profession'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>DOB/WD:</strong> ${entry['DOB/WD'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Product Interested:</strong> ${entry['Prodcut Interested'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Remarks:</strong> ${entry['Remarks'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Next Follow-up Date:</strong> ${entry['Next Follow-up Date'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Relation With Staff:</strong> ${entry['Relation With Staff'] || 'N/A'}</p>`;

                        entryLi.innerHTML = detailHtml;
                        entryDetailsList.appendChild(entryLi);
                    });

                    employeeLi.appendChild(entryDetailsList);
                    employeeList.appendChild(employeeLi);
                }
                branchDiv.appendChild(employeeList);
                reportsContainer.appendChild(branchDiv);
            }

        })
        .catch(error => {
            console.error("Error fetching or processing data:", error);
            // Display a user-friendly error message on the page
            reportsContainer.innerHTML = "<p>Error loading data. Please try again later. Check browser console for details.</p>";
        });
});