<script>
    // Function to handle form submission and update the attendance table
    document.getElementById('attendanceForm').addEventListener('submit', function(event) {
        event.preventDefault(); // Prevent the form from submitting the traditional way

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

        // Loop through each shift row and gather data
        document.querySelectorAll('#shiftTable tbody tr').forEach(function(row) {
            const shiftType = row.cells[0].textContent;
            const shiftData = {
                shiftType,
                absen: row.querySelector('[name$="Absen"]').value,
                tunda: row.querySelector('[name$="Tunda"]').value,
                sakit: row.querySelector('[name$="Sakit"]').value,
                izin: row.querySelector('[name$="Izin"]').value,
                cfv: row.querySelector('[name$="CFV"]').value,
                cl: row.querySelector('[name$="CL"]').value,
                ct: row.querySelector('[name$="CT"]').value,
                off: row.querySelector('[name$="OFF"]').value
            };
            data.shiftDetails.push(shiftData);
        });

        // Append gathered data to the table
        const tableBody = document.querySelector('#attendanceTable tbody');
        data.shiftDetails.forEach(shift => {
            // Add rows for each shift with respective status
            const shiftCombinations = [
                { type: `${shift.shiftType} & Hadir`, status: shift.absen === '' ? 'Hadir' : '' },
                { type: `${shift.shiftType} & Absen`, status: shift.absen !== '' ? 'Absen' : '' },
                { type: `${shift.shiftType} & Off`, status: shift.off !== '' ? 'Off' : '' },
                { type: `${shift.shiftType} & Tunda`, status: shift.tunda !== '' ? 'Tunda' : '' },
                { type: `${shift.shiftType} & Sakit`, status: shift.sakit !== '' ? 'Sakit' : '' },
                { type: `${shift.shiftType} & Izin`, status: shift.izin !== '' ? 'Izin' : '' },
                { type: `${shift.shiftType} & CFV`, status: shift.cfv !== '' ? 'CFV' : '' },
                { type: `${shift.shiftType} & CL`, status: shift.cl !== '' ? 'CL' : '' },
                { type: `${shift.shiftType} & CT`, status: shift.ct !== '' ? 'CT' : '' },
            ];

            // For each combination, add a row if there's a status
            shiftCombinations.forEach(combination => {
                if (combination.status !== '') {
                    const row = document.createElement('tr');
                    row.innerHTML = `
                        <td>${data.inputDate}</td>
                        <td>${data.employeeId}</td>
                        <td>${data.PT}</td>
                        <td>${data.departemen}</td>
                        <td>${data.division}</td>
                        <td>${data.team}</td>
                        <td>${combination.type}</td>
                        <td>${combination.status}</td>
                    `;
                    tableBody.appendChild(row);
                }
            });
        });

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
