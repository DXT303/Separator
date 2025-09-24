<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="icon" type="image/png" href="logo.png">
    <title>TEAM FROGGY</title>

    <!-- XLSX.js Library -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.25/jspdf.plugin.autotable.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>

    <style>
        @import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;700&display=swap');

        body {
            font-family: 'Orbitron', sans-serif;
            background: url('https://images.unsplash.com/photo-1550751827-4bd374c3f58b') no-repeat center center fixed;
            background-size: cover;
            color: #0ff;
            text-align: center;
            padding: 50px;
            margin: 0;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            flex-direction: column;
        }

        .container {
            background: rgba(0, 0, 0, 0.85);
            padding: 40px;
            border-radius: 12px;
            box-shadow: 0 0 25px rgba(0, 255, 255, 0.8);
            display: inline-block;
            max-width: 600px;
            animation: fadeIn 1.5s ease-in-out;
            border: 2px solid #0ff;
        }

        .panel {
            display: none;
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            background: rgba(0, 0, 0, 0.95);
            padding: 50px;
            width: 80vw;
            height: 80vh;
            border-radius: 15px;
            box-shadow: 0 0 30px #0ff;
            color: #0ff;
            text-align: center;
            border: 2px solid #0ff;
            overflow: auto;
        }

        .close-button {
            position: absolute;
            top: 20px;
            right: 20px;
            font-size: 30px;
            color: #0ff;
            cursor: pointer;
            text-shadow: 0 0 10px #0ff;
            width: 50px;
            height: 50px;
            line-height: 50px;
            text-align: center;
            border-radius: 50%;
            transition: all 0.3s ease-in-out;
        }

        .action-button {
            background-color: #0ff;
            color: #000;
            font-size: 14px;
            padding: 6px 12px;
            border: 2px solid #0ff;
            border-radius: 6px;
            cursor: pointer;
            box-shadow: 0 0 10px #0ff;
            transition: all 0.3s ease-in-out;
            text-align: center;
        }

        .action-button:hover {
            background-color: #000;
            color: #0ff;
            box-shadow: 0 0 20px #0ff, 0 0 40px #0ff;
        }

        .close-button:hover {
            background-color: #000;
            color: #0ff;
            box-shadow: 0 0 25px #0ff;
            border-radius: 50%;
        }

        .button-container {
            display: flex;
            justify-content: center;
            gap: 20px;
            margin-top: 20px;
            flex-wrap: wrap;
        }

        .start-button, .file-input-label {
            background-color: #0ff;
            color: #000;
            font-size: 18px;
            padding: 12px 25px;
            border: 2px solid #0ff;
            border-radius: 8px;
            cursor: pointer;
            box-shadow: 0 0 15px #0ff;
            transition: all 0.3s ease-in-out;
            position: relative;
            overflow: hidden;
            display: inline-block;
            text-align: center;
        }

        .start-button:hover, .file-input-label:hover {
            background-color: #000;
            color: #0ff;
            box-shadow: 0 0 25px #0ff, 0 0 50px #0ff;
        }

        .start-button:disabled {
            background-color: #666;
            color: #999;
            cursor: not-allowed;
            box-shadow: none;
        }

        #fileInput {
            display: none;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            color: #0ff;
        }

        th, td {
            border: 1px solid #0ff;
            padding: 10px;
            text-align: left;
        }

        #downloadAllBtn {
            background-color: #ff6b00;
            border-color: #ff6b00;
            box-shadow: 0 0 15px #ff6b00;
            margin-top: 20px;
            font-weight: bold;
        }

        #downloadAllBtn:hover {
            background-color: #000;
            color: #ff6b00;
            box-shadow: 0 0 25px #ff6b00, 0 0 50px #ff6b00;
        }

        .glow-text {
            text-shadow: 0 0 10px #0ff, 0 0 20px #0ff, 0 0 30px #0ff;
            animation: glow 2s ease-in-out infinite alternate;
        }

        @keyframes glow {
            from { text-shadow: 0 0 5px #0ff, 0 0 10px #0ff, 0 0 15px #0ff; }
            to { text-shadow: 0 0 10px #0ff, 0 0 20px #0ff, 0 0 30px #0ff; }
        }

        @keyframes fadeIn {
            from { opacity: 0; transform: scale(0.9); }
            to { opacity: 1; transform: scale(1); }
        }
    </style>
</head>
<body>
    <div class="container">
        <h1><span class="glow-text">TEAM FROGGY</span></h1>
        <p>Cybernetics, Code, and Chaos. The Digital Nightmare begins.</p>
        
        <button class="start-button" onclick="showPanel('filePanel')">Start</button>
    </div>

    <div id="filePanel" class="panel">
        <span class="close-button" onclick="closePanel('filePanel')">&#10006;</span>
        <h1 id="fileTitle">Upload Your File</h1>
        <div class="button-container">
            <label for="fileInput" class="file-input-label">Choose File</label>
            <input type="file" id="fileInput" accept=".xlsx" onchange="updateFileName()">
            <button class="start-button" onclick="viewFile()">View File</button>
            <button class="start-button" onclick="manipulateFileContent()">Separate</button>
            <button class="start-button" onclick="PrintF()">Print</button>
            <button class="start-button" id="sendEmailBtn" onclick="SendEmail()">Send Email</button>
            <button class="start-button" id="calculate" onclick="Calculate()">Calculate</button>
        </div>
        <div id="fileContent"></div>
        <div id="downloadAllContainer" style="display: none;">
            <button class="start-button" id="downloadAllBtn" onclick="downloadAllFiles()">📁 Download All Files</button>
        </div>
    </div>

    <script>
        let separatedFiles = [];
        let allDepartments = [];

        async function PrintF() {
            const fileInput = document.getElementById('fileInput');

            if (fileInput.files.length === 0) {
                alert("Please upload a file first.");
                return;
            }

            const file = fileInput.files[0];
            const reader = new FileReader();

            reader.onload = function (event) {
                try {
                    console.log("Reading file...");

                    const data = new Uint8Array(event.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const sheet = workbook.Sheets[workbook.SheetNames[0]];

                    let jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

                    console.log("Raw JSON Data:", jsonData);

                    jsonData = jsonData.filter(row => row.some(cell => cell !== ""));
                    
                    let startIndex = jsonData.findIndex(row => row.includes("ID") && row.includes("Name"));
                    if (startIndex === -1) {
                        alert("Error: Could not find table headers in the file.");
                        return;
                    }

                    jsonData = jsonData.slice(startIndex);

                    console.log("Cleaned JSON Data:", jsonData);

                    if (!jsonData || jsonData.length < 2) {
                        alert("Error: The Excel file is empty or missing valid data.");
                        return;
                    }

                    const headers = jsonData[0];
                    const rows = jsonData.slice(1);

                    console.log("Headers:", headers);
                    console.log("Rows:", rows);

                    if (!headers || headers.length === 0) {
                        alert("Error: No headers found in the Excel file.");
                        return;
                    }

                    const departmentIndex = headers.indexOf("Department");
                    if (departmentIndex !== -1) {
                        rows.forEach(row => {
                            if (row[departmentIndex] && typeof row[departmentIndex] === "string") {
                                row[departmentIndex] = row[departmentIndex].replace("All Departments>", "");
                            }
                        });
                    }

                    const { jsPDF } = window.jspdf;
                    const pdf = new jsPDF({
                        orientation: "portrait",
                        unit: "pt",
                        format: [612, 792],
                        margin: 0
                    });

                    pdf.setFontSize(14);
                    pdf.text("KFL MANPOWER AGENCY", pdf.internal.pageSize.width / 2, 30, { align: "center" });

                    pdf.autoTable({
                        head: [headers],
                        body: rows,
                        startY: 50,
                        theme: "grid",
                        styles: { 
                            fontSize: 8, 
                            cellPadding: 3, 
                            halign: "center", 
                            lineWidth: 0.5,
                            lineColor: [0, 0, 0]
                        },
                        headStyles: { 
                            fillColor: [100, 100, 100], 
                            textColor: 255, 
                            fontStyle: "bold",
                            lineWidth: 0.75
                        },
                        alternateRowStyles: {
                            fillColor: [230, 247, 255]
                        },
                        tableLineColor: [0, 0, 0],
                        tableLineWidth: 0.5,
                        tableWidth: "auto",
                        margin: 0,
                        didDrawPage: function (data) {
                            if (data.pageCount === data.pageNumber) {
                                const totalPages = pdf.internal.getNumberOfPages();
                                const currentPage = pdf.internal.getCurrentPageInfo().pageNumber;

                                pdf.setFontSize(10);
                                let pageText = `${currentPage}`;

                                pdf.text(pageText, pdf.internal.pageSize.width / 2, pdf.internal.pageSize.height - 10, {
                                    align: 'center'
                                });
                            }
                        }
                    });

                    console.log("PDF created, downloading...");
                    pdf.save("Converted_File.pdf");

                } catch (error) {
                    console.error("Full Error:", error);
                    alert("An error occurred while generating the PDF. Check console for details.");
                }
            };

            reader.readAsArrayBuffer(file);
        }

        function filterTable() {
            let input = document.getElementById("searchInput");
            let filter = input.value.toLowerCase();
            let table = document.getElementById("dataTable");
            let rows = table.getElementsByTagName("tr");
            let notFoundDiv = document.getElementById("notFound");

            let found = false;
            for (let i = 1; i < rows.length; i++) {
                let td = rows[i].getElementsByTagName("td")[0];
                if (td) {
                    let txtValue = td.textContent || td.innerText;
                    if (txtValue.toLowerCase().includes(filter)) {
                        rows[i].style.display = "";
                        found = true;
                    } else {
                        rows[i].style.display = "none";
                    }
                }
            }

            notFoundDiv.style.display = found ? "none" : "block";
        }

        function showPanel(panelId) {
            document.getElementById(panelId).style.display = 'block';
        }

        function closePanel(panelId) {
            document.getElementById(panelId).style.display = 'none';
        }

        function updateFileName() {
            const fileInput = document.getElementById('fileInput');
            const fileName = fileInput.files.length > 0 ? fileInput.files[0].name : "Upload Your File";
            document.getElementById('fileTitle').textContent = fileName;
        }

        function viewFile() {
            const fileInput = document.getElementById('fileInput');
            const file = fileInput.files[0];

            if (!file) {
                alert('Please upload a file first.');
                return;
            }

            const reader = new FileReader();
            reader.onload = function(event) {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheet = workbook.Sheets[workbook.SheetNames[0]];
                const html = XLSX.utils.sheet_to_html(sheet);

                document.getElementById('fileContent').innerHTML = html;
            };

            reader.readAsArrayBuffer(file);
        }

        function manipulateFileContent() {
            const fileInput = document.getElementById('fileInput');
            const file = fileInput.files[0];

            if (!file) {
                alert('Please upload a file first.');
                return;
            }

            const reader = new FileReader();
            reader.onload = function(event) {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheet = workbook.Sheets[workbook.SheetNames[0]];

                if (!sheet) {
                    alert("No sheets found in the file.");
                    return;
                }

                const range = XLSX.utils.decode_range(sheet['!ref']);
                let departmentColumn = null;

                for (let col = range.s.c; col <= range.e.c; col++) {
                    let cellAddress = XLSX.utils.encode_cell({ r: 7, c: col });
                    let cellValue = sheet[cellAddress] ? sheet[cellAddress].v.toString().trim().toLowerCase() : '';

                    if (cellValue === 'department') {
                        departmentColumn = col;
                        break;
                    }
                }

                if (departmentColumn === null) {
                    alert('No "Department" column found in row 8.');
                    return;
                }

                let departments = {};
                for (let row = 8; row <= range.e.r; row++) {
                    let cellAddress = XLSX.utils.encode_cell({ r: row, c: departmentColumn });
                    let department = sheet[cellAddress] ? sheet[cellAddress].v.toString().trim() : '';

                    if (department) {
                        let upperDept = department.toUpperCase();

                        if (upperDept.includes("AGRI-EXIM>PROJECT-BASE")) {
                            departments["PROJECT-BASE"] = (departments["PROJECT-BASE"] || 0) + 1;
                        } else if (upperDept.includes("AGRI-EXIM") && !upperDept.includes("PROJECT-BASE")) {
                            departments["AGRI-EXIM"] = (departments["AGRI-EXIM"] || 0) + 1;
                        } else {
                            departments[department] = (departments[department] || 0) + 1;
                        }
                    }
                }

                // Store departments for download all function
                allDepartments = Object.keys(departments);
                separatedFiles = []; // Reset separated files

                let pivotHtml = `
                <div style="width: 100%; display: flex; justify-content: center; margin-bottom: 50px; margin-top: 20px;">
                    <input type="text" id="searchInput" onkeyup="filterTable()" placeholder="Search Department..." 
                        style="width: 100%; padding: 10px; border: 2px solid #0ff; background: #000; color: #0ff; 
                        font-size: 16px; text-align: center; box-sizing: border-box;">
                </div>

                <table id="dataTable" style="width: 100%; border-collapse: collapse;">
                    <tr>
                        <th>Department</th>
                        <th style="text-align: center;">Count</th>
                        <th style="text-align: center;">Action</th>
                    </tr>`;

                for (const department in departments) {
                    pivotHtml += `<tr>
                        <td>${department}</td>
                        <td style="text-align: center;">${departments[department]}</td>
                        <td style="text-align: center;">
                            <button class="action-button" style="display: block; margin: 0 auto;" onclick="handleAction('${department}')">
                                Download
                            </button>
                        </td>
                    </tr>`;
                }

                pivotHtml += `</table>
                <div id="notFound" style="display: none; text-align: center; margin-top: 60px;">
                    <img src="https://media.giphy.com/media/UoeaPqYrimha6rdTFV/giphy.gif"
                        alt="Confused Robot"   
                        style="width: 500px; height: auto; border-radius: 10px;">
                    <p style="color: #0ff; font-size: 22px; margin-top: 15px; font-weight: bold;">
                        No results found!
                    </p>
                </div>`;

                document.getElementById('fileContent').innerHTML = pivotHtml;
                
                // Show download all button
                document.getElementById('downloadAllContainer').style.display = 'block';
            };

            reader.readAsArrayBuffer(file);
        }

        async function downloadAllFiles() {
            if (allDepartments.length === 0) {
                alert('No departments found. Please separate the file first.');
                return;
            }

            const downloadBtn = document.getElementById('downloadAllBtn');
            const originalText = downloadBtn.textContent;
            
            try {
                downloadBtn.textContent = 'Creating Files...';
                downloadBtn.disabled = true;

                const fileInput = document.getElementById('fileInput');
                const file = fileInput.files[0];

                if (!file) {
                    alert('Please upload a file first.');
                    return;
                }

                // Create all files
                const files = await createAllFilteredFiles(file, allDepartments);
                
                if (files.length === 0) {
                    alert('No files were created.');
                    return;
                }

                downloadBtn.textContent = 'Zipping Files...';

                // Create ZIP file
                const zip = new JSZip();
                
                files.forEach(fileData => {
                    zip.file(fileData.name, fileData.data);
                });

                downloadBtn.textContent = 'Downloading...';

                // Generate and download ZIP
                const zipBlob = await zip.generateAsync({ type: 'blob' });
                
                const url = URL.createObjectURL(zipBlob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `All_Departments_Files_${new Date().toISOString().slice(0, 10)}.zip`;
                
                a.style.display = 'none';
                document.body.appendChild(a);
                a.click();
                
                setTimeout(() => {
                    document.body.removeChild(a);
                    URL.revokeObjectURL(url);
                }, 100);

                alert(`Successfully downloaded ${files.length} files in a ZIP archive!`);

            } catch (error) {
                console.error('Download All Error:', error);
                alert('Failed to create ZIP file: ' + error.message);
            } finally {
                downloadBtn.textContent = originalText;
                downloadBtn.disabled = false;
            }
        }

        async function createAllFilteredFiles(mainFile, departments) {
            return new Promise((resolve) => {
                const reader = new FileReader();
                reader.onload = function(event) {
                    const data = new Uint8Array(event.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const sheet = workbook.Sheets[workbook.SheetNames[0]];
                    const range = XLSX.utils.decode_range(sheet['!ref']);
                    let departmentColumn = null;
                    let headers = [];

                    for (let col = range.s.c; col <= range.e.c; col++) {
                        let cellAddress = XLSX.utils.encode_cell({ r: 7, c: col });
                        let cellValue = sheet[cellAddress] ? sheet[cellAddress].v.toString().trim().toLowerCase() : '';
                        if (cellValue === 'department') {
                            departmentColumn = col;
                        }
                        headers.push(sheet[cellAddress] ? sheet[cellAddress].v : '');
                    }

                    const now = new Date();
                    const exportTime = `${String(now.getDate()).padStart(2, '0')}-${String(now.getMonth() + 1).padStart(2, '0')}-${now.getFullYear()} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
                    
                    const files = [];
                    
                    departments.forEach(dept => {
                        const headerInfo = [
                            ["KFL MANPOWER AGENCY SERVER 3"],
                            ["Transaction"],
                            [`Export Time: ${exportTime}`],
                            ["Operator: IT DEPARTMENT"],
                            [],
                            [`Department: ${dept}`],
                            [],
                        ];

                        let filteredData = [...headerInfo, headers];

                        for (let row = 8; row <= range.e.r; row++) {
                            let cellAddress = XLSX.utils.encode_cell({ r: row, c: departmentColumn });
                            let cellValue = sheet[cellAddress] ? sheet[cellAddress].v.toString().trim() : '';

                            if (cellValue.toUpperCase().includes(dept.toUpperCase())) {
                                let rowData = [];
                                for (let col = range.s.c; col <= range.e.c; col++) {
                                    let cell = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
                                    rowData.push(cell ? cell.v : '');
                                }
                                filteredData.push(rowData);
                            }
                        }

                        const newSheet = XLSX.utils.aoa_to_sheet(filteredData);
                        const newWorkbook = XLSX.utils.book_new();
                        XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Filtered Data");
                        
                        const wbout = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
                        files.push({
                            name: `${dept.replace(/[^a-zA-Z0-9]/g, '_')}_Data.xlsx`,
                            data: wbout
                        });
                    });
                    
                    resolve(files);
                };
                reader.readAsArrayBuffer(mainFile);
            });
        }

        function handleAction(departmentName) {
            const fileInput = document.getElementById('fileInput');
            const file = fileInput.files[0];

            if (!file) {
                alert('Please upload a file first.');
                return;
            }

            const reader = new FileReader();
            reader.onload = function (event) {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheet = workbook.Sheets[workbook.SheetNames[0]];

                if (!sheet) {
                    alert("No sheets found in the file.");
                    return;
                }

                const range = XLSX.utils.decode_range(sheet['!ref']);
                let departmentColumn = null;
                let headers = [];

                for (let col = range.s.c; col <= range.e.c; col++) {
                    let cellAddress = XLSX.utils.encode_cell({ r: 7, c: col });
                    let cellValue = sheet[cellAddress] ? sheet[cellAddress].v.toString().trim().toLowerCase() : '';

                    if (cellValue === 'department') {
                        departmentColumn = col;
                    }
                    headers.push(sheet[cellAddress] ? sheet[cellAddress].v : '');
                }

                if (departmentColumn === null) {
                    alert('No "Department" column found in row 8.');
                    return;
                }

                const now = new Date();
                const exportTime = `${String(now.getDate()).padStart(2, '0')}-${String(now.getMonth() + 1).padStart(2, '0')}-${now.getFullYear()} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;

                const headerInfo = [
                    ["KFL MANPOWER AGENCY SERVER 3"],
                    ["Transaction"],
                    [`Export Time: ${exportTime}`],
                    ["Operator: IT DEPARTMENT"],
                    [],
                    [`Department: ${departmentName}`],
                    [],
                ];

                let filteredData = [...headerInfo, headers];

                for (let row = 8; row <= range.e.r; row++) {
                    let cellAddress = XLSX.utils.encode_cell({ r: row, c: departmentColumn });
                    let cellValue = sheet[cellAddress] ? sheet[cellAddress].v.toString().trim() : '';

                    if (cellValue.toUpperCase().includes(departmentName.toUpperCase())) {
                        let rowData = [];
                        for (let col = range.s.c; col <= range.e.c; col++) {
                            let cell = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
                            rowData.push(cell ? cell.v : '');
                        }
                        filteredData.push(rowData);
                    }
                }

                if (filteredData.length === headerInfo.length + 1) {
                    alert(`No data found for department: ${departmentName}`);
                    return;
                }

                const newSheet = XLSX.utils.aoa_to_sheet(filteredData);
                const newWorkbook = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Filtered Data");

                const colWidths = [];
                const maxCols = filteredData[0].length;

                for (let col = 0; col < maxCols; col++) {
                    let maxLength = 0;
                    filteredData.forEach((row) => {
                        const cellValue = row[col] ? row[col].toString() : '';
                        if (cellValue.length > maxLength) {
                            maxLength = cellValue.length;
                        }
                    });
                    colWidths.push({ wch: maxLength + 2 });
                }
                newSheet['!cols'] = colWidths;

                const wbout = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
                const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
                
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `${departmentName.replace(/[^a-zA-Z0-9]/g, '_')}_Data.xlsx`;
                
                a.style.display = 'none';
                document.body.appendChild(a);
                a.click();
                
                setTimeout(() => {
                    document.body.removeChild(a);
                    URL.revokeObjectURL(url);
                }, 100);

                if (typeof separatedFiles !== 'undefined') {
                    separatedFiles.push({
                        name: `${departmentName.replace(/[^a-zA-Z0-9]/g, '_')}_Data.xlsx`,
                        blob: blob
                    });
                }
            };

            reader.readAsArrayBuffer(file);
        }

        // Placeholder functions - add your implementations
        function SendEmail() {
            alert('Send Email function - implement your email sending logic here');
        }

        function Calculate() {
            alert('Calculate function - implement your calculation logic here');
        }
    </script>
</body>
</html>