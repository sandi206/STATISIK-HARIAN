// Store the data in an array
const data = [];

// Handle form submission
document.getElementById('attendanceForm').addEventListener('submit', function(event) {
    event.preventDefault();

    const form = event.target;
    const newData = {
        employeeId: form.employeeId.value,
        pt: form.pt.value,
        department: form.department.value,
        division: form.division.value,
        team: form.team.value,
        shiftType: form.shiftType.value,
        attendance: form.attendance.value,
        absence: form.absence.value
    };

    // Add the new data to the array
    data.push(newData);
    form.reset();
    renderTable();
});

// Render the table
function renderTable() {
    const tableBody = document.querySelector('#attendanceTable tbody');
    tableBody.innerHTML = ''; // Clear previous table rows

    data.forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${row.employeeId}</td>
            <td>${row.pt}</td>
            <td>${row.department}</td>
            <td>${row.division}</td>
            <td>${row.team}</td>
            <td>${row.shiftType}</td>
            <td>${row.attendance}</td>
            <td>${row.absence}</td>
        `;
        tableBody.appendChild(tr);
    });
}

// Initialize Google API client
function handleClientLoad() {
    gapi.load('client:auth2', initClient);
}

function initClient() {
    gapi.auth2.init({
        client_id: 'YOUR_CLIENT_ID.apps.googleusercontent.com', // Replace with your Google Client ID
        scope: 'https://www.googleapis.com/auth/drive.file',
    }).then(function () {
        if (gapi.auth2.getAuthInstance().isSignedIn.get()) {
            console.log("User already signed in");
        } else {
            gapi.auth2.getAuthInstance().signIn();
        }
    });
}

// Upload file to Google Drive
function uploadFileToDrive(fileContent, fileName) {
    var fileMetadata = {
        'name': fileName,
        'mimeType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    };

    var media = {
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        body: fileContent
    };

    var request = gapi.client.drive.files.create({
        resource: fileMetadata,
        media: media,
        fields: 'id'
    });

    request.execute(function(response) {
        console.log('File uploaded with ID: ' + response.id);
    });
}

// Download the data as an Excel file
document.getElementById('downloadButton').addEventListener('click', function() {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, 'Data Kehadiran');

    const fileContent = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });

    // Upload the file to Google Drive
    uploadFileToDrive(fileContent, 'data_kehadiran.xlsx');
});
