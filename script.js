    <script>
        // Menangani pengisian form
        document.getElementById('attendanceForm').addEventListener('submit', function(e) {
            e.preventDefault();  // Mencegah form submit secara default

            // Ambil data dari form
            const formData = new FormData(this);
            const rowData = {
                date: formData.get('inputDate'),
                employeeId: formData.get('employeeId'),
                pt: formData.get('PT'),
                department: formData.get('Departemen'),
                division: formData.get('division'),
                team: formData.get('team'),
                shiftData: {
                    shiftRegulerAbsen: formData.get('shiftRegulerAbsen'),
                    shiftRegulerTunda: formData.get('shiftRegulerTunda'),
                    shiftRegulerSakit: formData.get('shiftRegulerSakit'),
                    shiftRegulerIzin: formData.get('shiftRegulerIzin'),
                    shiftRegulerCFV: formData.get('shiftRegulerCFV'),
                    shiftRegulerCL: formData.get('shiftRegulerCL'),
                    shiftRegulerCT: formData.get('shiftRegulerCT'),
                    shiftRegulerOFF: formData.get('shiftRegulerOFF'),
                }
            };

            // Tambahkan data ke dalam tabel
            const table = document.querySelector('#attendanceTable tbody');
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${rowData.date}</td>
                <td>${rowData.employeeId}</td>
                <td>${rowData.pt}</td>
                <td>${rowData.department}</td>
                <td>${rowData.division}</td>
                <td>${rowData.team}</td>
                <td>Shift Reguler</td>
                <td>${rowData.shiftData.shiftRegulerAbsen}</td>
                <td>${rowData.shiftData.shiftRegulerTunda}</td>
                <td>${rowData.shiftData.shiftRegulerCFV}</td>
                <td>${rowData.shiftData.shiftRegulerCL}</td>
                <td>${rowData.shiftData.shiftRegulerCT}</td>
                <td>${rowData.shiftData.shiftRegulerOFF}</td>
            `;
            table.appendChild(row);
        });

        // Menangani download Excel
        document.getElementById('downloadButton').addEventListener('click', function() {
            const table = document.getElementById('attendanceTable');
            const wb = XLSX.utils.table_to_book(table, {sheet: "Attendance"});
            XLSX.writeFile(wb, "attendance_data.xlsx");
        });
    </script>
