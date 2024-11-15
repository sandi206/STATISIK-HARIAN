<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.0/xlsx.full.min.js"></script>

<script>
    document.getElementById('attendanceForm').addEventListener('submit', function(event) {
        event.preventDefault(); // Prevent form from submitting traditionally

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

        // Collect shift data
        document.querySelectorAll('#shiftTable tbody tr').forEach(function(row) {
            const shiftType = row.cells[0].textContent; // The first cell contains the shift type
            const shiftData = {
                shiftType,
                absen: row.querySelector('[name$="Hadir"]').value,
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

        // Append the collected data to the table
        const tableBody = document.querySelector('#attendanceTable tbody');
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${data.inputDate}</td>
            <td>${data.employeeId}</td>
            <td>${data.PT}</td>
            <td>${data.departemen}</td>
            <td>${data.division}</td>
            <td>${data.team}</td>
            <td>${data.shiftDetails.map(shift => shift.shiftType).join(', ')}</td>
            <td>${data.shiftDetails.map(shift => shift.absen).join(', ')}</td>
            <td>${data.shiftDetails.map(shift => shift.tunda).join(', ')}</td>
            <td>${data.shiftDetails.map(shift => shift.cfv).join(', ')}</td>
            <td>${data.shiftDetails.map(shift => shift.cl).join(', ')}</td>
            <td>${data.shiftDetails.map(shift => shift.ct).join(', ')}</td>
            <td>${data.shiftDetails.map(shift => shift.off).join(', ')}</td>
        `;
        tableBody.appendChild(row);

        // Show the download button
        document.getElementById('downloadButton').style.display = 'inline-block';

        // Function to generate Excel file when the download button is clicked
        document.getElementById('downloadButton').addEventListener('click', function() {
            const table = document.getElementById('attendanceTable');
            const workbook = XLSX.utils.table_to_book(table, { sheet: "Data Kehadiran" });
            const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' });

            function s2ab(s) {
                const buf = new ArrayBuffer(s.length);
                const view = new Uint8Array(buf);
                for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
                return buf;
            }

            // Create a Blob for the Excel file
            const blob = new Blob([s2ab(wbout)], { type: "application/octet-stream" });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = 'data_kehadiran.xlsx'; // Specify the name of the Excel file
            link.click(); // Trigger the download
        });
    });
</script>
