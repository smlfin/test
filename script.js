document.addEventListener('DOMContentLoaded', () => {
    // NEW DATA URL: This will be the PUBLISHED CSV link from Google Sheets
    // Replace with the URL you copied from File > Share > Publish to web (as .csv)
    const DATA_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTO7LujC4VSa2wGkJ2YEYSN7UeXR221ny3THaVegYfNfRm2JQGg7QR9Bxxh9SadXtK8Pi6-psl2tGsb/pub?gid=696550092&single=true&output=csv"; // **IMPORTANT: Update this!**

    const reportsContainer = document.getElementById('reportsContainer');
    if (!reportsContainer) {
        console.error("Error: Element with ID 'reportsContainer' not found in index.html. Cannot display reports.");
        return;
    }
    reportsContainer.innerHTML = "<p>Loading canvassing data...</p>";

    fetch(DATA_URL)
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            // Fetch the response as text, because it's CSV, not JSON
            return response.text();
        })
        .then(csvText => {
            console.log("Fetched raw CSV data:", csvText); // Log raw CSV for debugging

            // --- Parse CSV to JSON (Array of Objects) ---
            // This is a simple CSV parser. For complex CSVs, consider a library.
            const lines = csvText.trim().split('\n'); // Split by line, remove empty last line
            if (lines.length <= 1) {
                // Only headers or no data
                reportsContainer.innerHTML = "<p>No canvassing data available.</p>";
                return;
            }

            const headers = lines[0].split(',').map(header => header.trim()); // Get headers from first line
            const data = [];

            for (let i = 1; i < lines.length; i++) {
                const values = lines[i].split(',').map(value => value.trim()); // Split values by comma
                let entry = {};
                headers.forEach((header, index) => {
                    entry[header] = values[index];
                });
                data.push(entry);
            }

            console.log("Parsed data (from CSV):", data); // Log parsed data for debugging

            if (!data || data.length === 0) {
                reportsContainer.innerHTML = "<p>No canvassing data available.</p>";
                return;
            }

            // --- Display total submissions count ---
            const totalSubmissions = data.length;
            // Assuming you still have 'totalSubmissions' span in index.html, otherwise remove this
            // You might want to add back a summary section in index.html
            // For now, let's just make sure the detailed report renders
            const totalSubmissionsSpan = document.getElementById('totalSubmissions');
            if (totalSubmissionsSpan) {
                totalSubmissionsSpan.textContent = totalSubmissions;
            }


            // Group data by Branch Name, then by Employee Name
            const groupedData = {};

            data.forEach(entry => {
                let branchName = entry['Branch Name'];
                let employeeName = entry['Employee Name'];

                // Handle potential missing values for grouping
                if (!branchName || branchName.trim() === '') {
                    branchName = "Unassigned Branch";
                }
                if (!employeeName || employeeName.trim() === '') {
                    employeeName = "Unassigned Employee";
                }

                if (!groupedData[branchName]) {
                    groupedData[branchName] = {};
                }
                if (!groupedData[branchName][employeeName]) {
                    groupedData[branchName][employeeName] = [];
                }
                groupedData[branchName][employeeName].push(entry);
            });

            console.log("Grouped data:", groupedData);

            // Clear loading message and prepare for rendering
            reportsContainer.innerHTML = '';

            // Render the grouped data
            for (const branch in groupedData) {
                const branchDiv = document.createElement('div');
                branchDiv.className = 'branch-group';
                branchDiv.innerHTML = `<h2>Branch Name: ${branch}</h2>`;

                const employeeList = document.createElement('ul');
                employeeList.className = 'employee-list';

                for (const employee in groupedData[branch]) {
                    const employeeLi = document.createElement('li');
                    employeeLi.className = 'employee-item';
                    employeeLi.innerHTML = `<h3>Employee Name: ${employee}</h3>`;

                    const entryDetailsList = document.createElement('ul');
                    entryDetailsList.className = 'entry-details-list';

                    groupedData[branch][employee].forEach(entry => {
                        const entryLi = document.createElement('li');
                        entryLi.className = 'entry-detail-item';

                        // Dynamically display all relevant fields for each entry
                        let detailHtml = `<h4>Canvassing Entry</h4>`;
                        detailHtml += `<p><strong>Timestamp:</strong> ${entry['Timestamp'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Date:</strong> ${entry['Date'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Employee Code:</strong> ${entry['Employee Code'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Designation:</strong> ${entry['Designation'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Activity Type:</strong> ${entry['Activity Type'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Type of Customer:</strong> ${entry['Type of Customer'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Lead Source:</strong> ${entry['Lead Source'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>How Contacted:</strong> ${entry['How Contacted'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Prospect Name:</strong> ${entry['Prospect Name'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Phone Numebr(Whatsapp):</strong> ${entry['Phone Numebr(Whatsapp)'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Address:</strong> ${entry['Address'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Profession:</strong> ${entry['Profession'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>DOB/WD:</strong> ${entry['DOB/WD'] || 'N/A'}</p>`;
                        detailHtml += `<p><strong>Prodcut Interested:</strong> ${entry['Prodcut Interested'] || 'N/A'}</p>`;
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
            const mainContent = document.querySelector('main') || document.body;
            mainContent.innerHTML = "<p>Error loading data. Please try again later. Check browser console for details.</p>";
        });
});