<script>
    // Function to handle form submission and update the attendance table
    document.getElementById('attendanceForm').addEventListener('submit', function(event) {
        event.preventDefault(); // Prevent the form from submitting the traditional way

        // Collect form data
        const formData = new FormData(event.target);
        const data = {
            inputDate: formData.get('inputDate'),
            employeeId: formData.get('employeeId'),
            PT: formData.get('PT'),
            departemen: formData.get('Departemen'),
            division: formData.get('division'),
            team: formData.get('team'),
            shiftDetails: [],
        };

        // Loop through the shift rows and gather data
        document.querySelectorAll('#shiftTable tbody tr').forEach(function(row) {
            const shiftType = row.cells[0].textContent;
            const shiftData = {
                shiftType,
                absen: row.querySelector('[name$="Absen"]') ? row.querySelector('[name$="Absen"]').value : '',
                tunda: row.querySelector('[name$="Tunda"]') ? row.querySelector('[name$="Tunda"]').value : '',
                sakit: row.querySelector('[name$="Sakit"]') ? row.querySelector('[name$="Sakit"]').value : '',
                izin: row.querySelector('[name$="Izin"]') ? row.querySelector('[name$="Izin"]').value : '',
                cfv: row.querySelector('[name$="CFV"]') ? row.querySelector('[name$="CFV"]').value : '',
                cl: row.querySelector('[name$="CL"]') ? row.querySelector('[name$="CL"]').value : '',
                ct: row.querySelector('[name$="CT"]') ? row.querySelector('[name$="CT"]').value : '',
                off: row.querySelector('[name$="OFF"]') ? row.querySelector('[name$="OFF"]').value : ''
            };
            data.shiftDetails.push(shiftData);
        });

        // Append gathered data to the table
        const tableBody = document.querySelector('#attendanceTable tbody');
        
        const shiftStatusColumns = {
            "shiftRegulerHadir": "",
            "shiftRegulerAbsen": "",
            "shiftRegulerOff": "",
            "shiftTunda": "",
            "shiftSakit": "",
            "shiftIzin": "",
            "shiftCFV": "",
            "shiftCL": "",
            "shiftCT": ""
        };

        // Loop through each shift details to populate columns
        data.shiftDetails.forEach(shift => {
            if (shift.absen !== '') shiftStatusColumns.shiftRegulerAbsen = shift.absen;
            if (shift.off !== '') shiftStatusColumns.shiftRegulerOff = shift.off;
            if (shift.sakit !== '') shiftStatusColumns.shiftSakit = shift.sakit;
            if (shift.izin !== '') shiftStatusColumns.shiftIzin = shift.izin;
            if (shift.cfv !== '') shiftStatusColumns.shiftCFV = shift.cfv;
            if (shift.cl !== '') shiftStatusColumns.shiftCL = shift.cl;
            if (shift.ct !== '') shiftStatusColumns.shiftCT = shift.ct;
        });

        // Create the row to append to the table
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${data.inputDate}</td>
            <td>${data.employeeId}</td>
            <td>${data.PT}</td>
            <td>${data.departemen}</td>
            <td>${data.division}</td>
            <td>${data.team}</td>
            <td>${shiftStatusColumns.shiftRegulerHadir}</td>
            <td>${shiftStatusColumns.shiftRegulerAbsen}</td>
            <td>${shiftStatusColumns.shiftRegulerOff}</td>
            <td>${shiftStatusColumns.shiftTunda}</td>
            <td>${shiftStatusColumns.shiftSakit}</td>
            <td>${shiftStatusColumns.shiftIzin}</td>
            <td>${shiftStatusColumns.shiftCFV}</td>
            <td>${shiftStatusColumns.shiftCL}</td>
            <td>${shiftStatusColumns.shiftCT}</td>
        `;
        tableBody.appendChild(row);

        // Optionally, reset the form after adding the row
        event.target.reset();
    });

    // Function to export the attendance table data to Excel
    document.getElementById('downloadButton').addEventListener('click', function() {
        const table = document.getElementById('attendanceTable');
        
        // Convert the HTML table to an Excel workbook
        const workbook = XLSX.utils.table_to_book(table, { sheet: "Data Kehadiran" });
        const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' });

        // Save the Excel file
        function s2ab(s) { 
            const buf = new ArrayBuffer(s.length); 
            const view = new Uint8Array(buf); 
            for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF; 
            return buf; 
        }

        const blob = new Blob([s2ab(wbout)], { type: "application/octet-stream" });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'data_kehadiran.xlsx'; // Specify the filename for the Excel file
        link.click();
    });
</script>
