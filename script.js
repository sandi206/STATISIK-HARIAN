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

                    shiftRegulerAbsen: formData.get('shift34Absen'),
                    shiftRegulerTunda: formData.get('shift34Tunda'),
                    shiftRegulerSakit: formData.get('shift34Sakit'),
                    shiftRegulerIzin: formData.get('shift34Izin'),
                    shiftRegulerCFV: formData.get('shift34CFV'),
                    shiftRegulerCL: formData.get('shift34CL'),
                    shiftRegulerCT: formData.get('shift34CT'),
                    shiftRegulerOFF: formData.get('shift34OFF'),

                    shiftRegulerAbsen: formData.get('shift33Absen'),
                    shiftRegulerTunda: formData.get('shift33Tunda'),
                    shiftRegulerSakit: formData.get('shift33Sakit'),
                    shiftRegulerIzin: formData.get('shift33Izin'),
                    shiftRegulerCFV: formData.get('shift33CFV'),
                    shiftRegulerCL: formData.get('shift33CL'),
                    shiftRegulerCT: formData.get('shift33CT'),
                    shiftRegulerOFF: formData.get('shift33OFF'),

                    shiftRegulerAbsen: formData.get('shift23Absen'),
                    shiftRegulerTunda: formData.get('shift23Tunda'),
                    shiftRegulerSakit: formData.get('shift23Sakit'),
                    shiftRegulerIzin: formData.get('shift23Izin'),
                    shiftRegulerCFV: formData.get('shift23CFV'),
                    shiftRegulerCL: formData.get('shift23CL'),
                    shiftRegulerCT: formData.get('shift23CT'),
                    shiftRegulerOFF: formData.get('shift33OFF'),

                    shiftRegulerAbsen: formData.get('shift12Absen'),
                    shiftRegulerTunda: formData.get('shift12Tunda'),
                    shiftRegulerSakit: formData.get('shift12Sakit'),
                    shiftRegulerIzin: formData.get('shift12Izin'),
                    shiftRegulerCFV: formData.get('shift12CFV'),
                    shiftRegulerCL: formData.get('shift12CL'),
                    shiftRegulerCT: formData.get('shift12CT'),
                    shiftRegulerOFF: formData.get('shift12OFF'),

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

                <td>3 shift 4 regu</td>
                <td>${rowData.shiftData.shift34Absen}</td>
                <td>${rowData.shiftData.shift34Tunda}</td>
                <td>${rowData.shiftData.shift34CFV}</td>
                <td>${rowData.shiftData.shift34CL}</td>
                <td>${rowData.shiftData.shift34CT}</td>
                <td>${rowData.shiftData.shift34OFF}</td>


                <td>3 shift 4 regu</td>
                <td>${rowData.shiftData.shift33Absen}</td>
                <td>${rowData.shiftData.shift33Tunda}</td>
                <td>${rowData.shiftData.shift33CFV}</td>
                <td>${rowData.shiftData.shift33CL}</td>
                <td>${rowData.shiftData.shift33CT}</td>
                <td>${rowData.shiftData.shift33OFF}</td>


                <td>3 shift 4 regu</td>
                <td>${rowData.shiftData.shift23Absen}</td>
                <td>${rowData.shiftData.shift23Tunda}</td>
                <td>${rowData.shiftData.shift23CFV}</td>
                <td>${rowData.shiftData.shift23CL}</td>
                <td>${rowData.shiftData.shift23CT}</td>
                <td>${rowData.shiftData.shift23OFF}</td>


                <td>3 shift 4 regu</td>
                <td>${rowData.shiftData.shift12Absen}</td>
                <td>${rowData.shiftData.shift12Tunda}</td>
                <td>${rowData.shiftData.shift12CFV}</td>
                <td>${rowData.shiftData.shift12CL}</td>
                <td>${rowData.shiftData.shift12CT}</td>
                <td>${rowData.shiftData.shift12OFF}</td>
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
