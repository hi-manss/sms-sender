let workbook, sheetData, currentRowIndex = 0;
const sentRows = [];
const failedRows = [];
const skippedRows = [];
const phoneNumber = "7718955555"; // Hardcoded number
let isExporting = false; // Flag to prevent multiple downloads

// Load the Excel file and display records in the table
document.getElementById('fileInput').addEventListener('change', (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        workbook = XLSX.read(data, { type: 'array' });
        // Read from A1 and skip empty rows
        sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 })
            .filter(row => row[0] != null && row[0].trim() !== ''); // Filter out null or empty rows
        if (sheetData.length > 0) {
            loadRecordTable();
            updateNotification("File loaded successfully! Click 'Start Processing SMS'.");
            document.getElementById('startButton').style.display = 'inline-block';
            currentRowIndex = 0; // Reset the index when loading a new file
        }
    };
    reader.readAsArrayBuffer(file);
});

// Load the records into the table
function loadRecordTable() {
    const tableBody = document.getElementById('recordTableBody');
    tableBody.innerHTML = ""; // Clear previous records

    sheetData.forEach((row, index) => {
        // Only process rows that contain a message
        if (row.length > 0) { // Ensure the row is not empty
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${row[0] || "N/A"}</td> <!-- Display the first cell -->
                <td class="status" id="status-${index}">Pending</td>
                <td>
                    <button id="actionBtn-${index}" class="retry-btn" style="display:none;" onclick="retryMessage(${index})">Retry</button>
                </td>
            `;
            tableBody.appendChild(tr);
        }
    });
}

// Start processing records one by one
function startProcess() {
    if (!sheetData || sheetData.length === 0) {
        updateNotification("No records to process.");
        return;
    }

    // Disable the Start Processing button
    document.getElementById('startButton').disabled = true;

    processNextRecord();
}

// Process the next record
function processNextRecord() {
    if (currentRowIndex >= sheetData.length) {
        showDownloadButtons(); // Show download buttons after processing
        return; // Exit if all records are processed
    }

    const message = sheetData[currentRowIndex][0]; // Get the message from the first column
    const statusCell = document.getElementById(`status-${currentRowIndex}`);
    const actionButton = document.getElementById(`actionBtn-${currentRowIndex}`);

    // Use phone link to send the message
    const smsLink = `sms:${phoneNumber}?body=${encodeURIComponent(message)}`;
    window.open(smsLink, '_blank'); // Open SMS app

    // Show confirmation popup
    document.getElementById("popupMessage").innerText = `Send message: "${message}" to ${phoneNumber}?`;
    document.getElementById("recordIndex").innerText = `Record Index: ${currentRowIndex + 1}`;
    document.getElementById("popupMessageBody").innerText = `"${message}"`; // Display message body

    // Set up button event listeners for popup
    document.getElementById("successBtn").onclick = () => {
        // Set status to success
        statusCell.textContent = "Success";
        statusCell.classList.add("success");
        sentRows.push(sheetData[currentRowIndex]);
        actionButton.style.display = 'none'; // Hide the retry button
        closePopup();

        currentRowIndex++; // Move to the next record
        processNextRecord(); // Continue processing
    };

    document.getElementById("failedBtn").onclick = () => {
        // Set status to failed
        statusCell.textContent = "Failed";
        statusCell.classList.add("fail");
        failedRows.push(sheetData[currentRowIndex]);
        actionButton.style.display = 'inline'; // Show the retry button
        closePopup();

        currentRowIndex++; // Move to the next record
        processNextRecord(); // Continue processing
    };

    document.getElementById("skipBtn").onclick = () => {
        // Set status to skipped
        statusCell.textContent = "Skipped";
        skippedRows.push(sheetData[currentRowIndex]);
        actionButton.style.display = 'inline'; // Show the retry button
        closePopup();

        currentRowIndex++; // Move to the next record
        processNextRecord(); // Continue processing
    };

    // Show the popup
    document.getElementById("messagePopup").style.display = "flex";
}

// Close the confirmation popup
function closePopup() {
    document.getElementById("messagePopup").style.display = "none";
}

// Retry failed or skipped messages
function retryMessage(index) {
    // Get the message to resend
    const message = sheetData[index][0]; // Get the message from the first column
    const statusCell = document.getElementById(`status-${index}`);
    const actionButton = document.getElementById(`actionBtn-${index}`);

    // Show confirmation popup
    document.getElementById("popupMessage").innerText = `Retry sending message: "${message}" to ${phoneNumber}?`;
    document.getElementById("recordIndex").innerText = `Record Index: ${index + 1}`;
    document.getElementById("popupMessageBody").innerText = `Message Body: "${message}"`; // Display message body

    // Set up button event listeners for popup (reuse the existing logic)
    document.getElementById("successBtn").onclick = () => {
        // Set status to success
        statusCell.textContent = "Success";
        statusCell.classList.add("success");
        sentRows.push(sheetData[index]);
        actionButton.style.display = 'none'; // Hide the retry button
        closePopup();
    };

    document.getElementById("failedBtn").onclick = () => {
        // Set status to failed
        statusCell.textContent = "Failed";
        statusCell.classList.add("fail");
        failedRows.push(sheetData[index]);
        actionButton.style.display = 'inline'; // Show the retry button
        closePopup();
    };

    document.getElementById("skipBtn").onclick = () => {
        // Set status to skipped
        statusCell.textContent = "Skipped";
        skippedRows.push(sheetData[index]);
        actionButton.style.display = 'inline'; // Show the retry button
        closePopup();
    };

    // Open the SMS link again
    const smsLink = `sms:${phoneNumber}?body=${encodeURIComponent(message)}`;
    window.open(smsLink, '_blank'); // Open SMS app

    // Show the popup
    document.getElementById("messagePopup").style.display = "flex"; // Ensure the popup is shown
}

// Show download buttons when processing is done
function showDownloadButtons() {
    updateNotification("Processing complete! You can download the processed data.");
    document.getElementById('downloadProcessedBtn').style.display = 'inline-block';
    document.getElementById('startOverButton').style.display = 'inline-block'; // Show Start Over button
}

// Update notification message
function updateNotification(message) {
    document.getElementById('notification').innerText = message;
}

// Function to handle start over confirmation
function confirmStartOver() {
    const confirmation = confirm("Are you sure you want to start over? This will reset all current progress.");
    if (confirmation) {
        resetAll();
    }
}

// Function to reset all data and UI elements
function resetAll() {
    // Clear the record table
    document.getElementById('recordTableBody').innerHTML = "";
    document.getElementById('notification').innerText = "";
    
    // Reset global variables
    workbook = null;
    sheetData = [];
    currentRowIndex = 0;
    sentRows.length = 0; // Clear the sentRows array
    failedRows.length = 0; // Clear the failedRows array
    skippedRows.length = 0; // Clear the skippedRows array
    
    // Hide the download button and start over button
    document.getElementById('downloadProcessedBtn').style.display = 'none';
    document.getElementById('startOverButton').style.display = 'none';

    // Enable the file input and start button
    document.getElementById('fileInput').value = ''; // Clear the file input
    document.getElementById('startButton').disabled = false; // Enable the Start Processing button
    document.getElementById('startButton').style.display = 'none'; // Hide the Start Processing button
}
function downloadProcessedData() {
    if (isExporting) {
        return; // Prevent multiple downloads
    }

    isExporting = true; // Set flag to prevent further downloads

    // Create a new workbook and add worksheets
    const workbook = XLSX.utils.book_new();
    
    // Prepare data for each sheet
    const sentData = sentRows.map(row => [row[0]]); // Only include the message in the row
    const failedData = failedRows.map(row => [row[0]]);
    const skippedData = skippedRows.map(row => [row[0]]);

    // Add Sent Messages sheet
    const sentSheet = XLSX.utils.aoa_to_sheet([["Sent Messages"], ...sentData]);
    XLSX.utils.book_append_sheet(workbook, sentSheet, "Sent Messages");

    // Add Failed Messages sheet
    const failedSheet = XLSX.utils.aoa_to_sheet([["Failed Messages"], ...failedData]);
    XLSX.utils.book_append_sheet(workbook, failedSheet, "Failed Messages");

    // Add Skipped Messages sheet
    const skippedSheet = XLSX.utils.aoa_to_sheet([["Skipped Messages"], ...skippedData]);
    XLSX.utils.book_append_sheet(workbook, skippedSheet, "Skipped Messages");

    // Create a download link for the workbook
    XLSX.writeFile(workbook, "processed_sms_data.xlsx");

    isExporting = false; // Reset flag
}

// Event listener for download button
document.getElementById('downloadProcessedBtn').addEventListener('click', downloadProcessedData);
