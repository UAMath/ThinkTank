let pyodide;
let selectedFile = null;

const dropZone = document.getElementById("drop-zone");
const fileInput = document.getElementById("file-input");
const fileInfo = document.getElementById("file-info");
const convertBtn = document.getElementById("convert-btn");
const status = document.getElementById("status");

// initialize Pyodide, install openpyxl, and load the process_data.py file
async function initPyodide() {
    pyodide = await loadPyodide();
    await pyodide.loadPackage("micropip");
    const micropip = pyodide.pyimport("micropip");
    await micropip.install("openpyxl");

    // fetch process_data.py script and load it into Pyodide
    const response = await fetch("process_data.py");
    const pythonCode = await response.text();
    await pyodide.runPythonAsync(pythonCode);

    status.innerText = "Python Ready! Drag a file above.";
}
initPyodide();

// drag & Drop Event Listeners
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

// convert on Button Click
convertBtn.addEventListener("click", async () => {
    if (!selectedFile) return;

    status.innerText = "Processing file...";
    convertBtn.disabled = true;

    try {
        const textContent = await selectedFile.text();

        // Pass TSV content directly to the Python function defined in process_data.py
        const convertFunc = pyodide.globals.get("convert_tsv_to_excel");
        const bytes = convertFunc(textContent);

        // Download Excel File
        const blob = new Blob([bytes.toJs()], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "bookings_output.xlsx";
        a.click();
        URL.revokeObjectURL(url);

        status.innerText = "Done! Excel file downloaded.";
    } catch (err) {
        status.innerText = "Error processing file!";
        console.error(err);
    } finally {
        convertBtn.disabled = false;
    }
});