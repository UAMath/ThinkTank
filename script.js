let pyodide;
let selectedFile = null;

const dropZone = document.getElementById("drop-zone");
const fileInput = document.getElementById("file-input");
const fileInfo = document.getElementById("file-info");
const convertBtn = document.getElementById("convert-btn");
const status = document.getElementById("status");

async function initPyodide() {
    status.innerText = "Initializing Python...";
    pyodide = await loadPyodide();
    
    await pyodide.loadPackage("micropip");
    const micropip = pyodide.pyimport("micropip");
    await micropip.install("openpyxl");

    // Fetch and load master_sheet.csv into Pyodide's virtual filesystem
    try {
        const masterResponse = await fetch("master_sheet.csv");
        const masterCsvText = await masterResponse.text();
        pyodide.FS.writeFile("master_sheet.csv", masterCsvText);
    } catch (err) {
        console.error("Failed to load master_sheet.csv:", err);
    }

    // Load process_data.py
    const response = await fetch("process_data.py");
    const pythonCode = await response.text();
    await pyodide.runPythonAsync(pythonCode);

    status.innerText = "Python Ready! Drag a file above.";
}
initPyodide();

// Drag & Drop Event Listeners
dropZone.addEventListener("click", () => fileInput.click());
fileInput.addEventListener("change", (e) => handleFile(e.target.files[0]));

dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("hover");
});
dropZone.addEventListener("dragleave", () => dropZone.classList.remove("hover"));
dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("hover");
    if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
});

function handleFile(file) {
    selectedFile = file;
    fileInfo.innerText = `Selected: ${file.name}`;
    if (pyodide) convertBtn.disabled = false;
}

// Convert on Button Click
convertBtn.addEventListener("click", async () => {
    if (!selectedFile) return;

    status.innerText = "Processing file...";
    convertBtn.disabled = true;

    try {
        const textContent = await selectedFile.text();

        const convertFunc = pyodide.globals.get("convert_tsv_to_excel");
        const bytes = convertFunc(textContent);

        const blob = new Blob([bytes.toJs()], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "bookings_output.xlsx";
        a.click();
        URL.revokeObjectURL(url);

        status.innerText = "Done! Excel file downloaded.";
    } catch (err) {
        status.innerText = "Error processing file! Check console for details.";
        console.error(err);
    } finally {
        convertBtn.disabled = false;
    }
});