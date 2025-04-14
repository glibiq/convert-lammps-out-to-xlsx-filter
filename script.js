document.addEventListener("DOMContentLoaded", () => {
    let allHeaders = new Set();
    let selectedHeaders = {};
    let orderedHeaders = [];
    let dataRows = [];

    document.getElementById('fileInput').addEventListener('change', function(event) {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function(e) {
            const lines = e.target.result.split("\n");
            dataRows = [];
            allHeaders.clear();
            orderedHeaders = [];

            let tempHeaders = []; // Stores headers for each block

            // Extract headers from all header lines and data
            lines.forEach(line => {
                if (line.startsWith("#")) {
                    tempHeaders = line.substring(1).trim().split(/\s+/);
                    tempHeaders.forEach(h => allHeaders.add(h));
                } else if (line.trim() !== "") {
                    let values = line.trim().split(/\s+/);
                    if (tempHeaders.length === values.length) {
                        dataRows.push({ headers: tempHeaders, values });
                    }
                }
            });

            orderedHeaders = Array.from(allHeaders).filter(h => h !== "Timestep"); // Remove Timestep from selection
            renderHeaderSelection();
        };

        reader.readAsText(file);
    });

    function renderHeaderSelection() {
        const container = document.getElementById("headerSelection");
        container.innerHTML = ""; // Clear previous selections

        orderedHeaders.forEach(header => {
            const checkbox = document.createElement("input");
            checkbox.type = "checkbox";
            checkbox.id = header;
            checkbox.checked = true;
            checkbox.addEventListener("change", toggleDownloadButton);

            selectedHeaders[header] = checkbox;

            const label = document.createElement("label");
            label.innerHTML = header;
            label.setAttribute("for", header);

            const div = document.createElement("div");
            div.appendChild(checkbox);
            div.appendChild(label);
            container.appendChild(div);
        });

        toggleDownloadButton();
    }

    function toggleDownloadButton() {
        const anyChecked = Object.values(selectedHeaders).some(cb => cb.checked);
        document.getElementById("downloadBtn").disabled = !anyChecked;
    }

    window.selectAll = function() {
        Object.values(selectedHeaders).forEach(cb => cb.checked = true);
        toggleDownloadButton();
    }

    window.unselectAll = function() {
        Object.values(selectedHeaders).forEach(cb => cb.checked = false);
        toggleDownloadButton();
    }

    window.downloadXLSX = function () {
        if (dataRows.length === 0) {
            alert("No data to process. Please upload a valid file.");
            return;
        }
    
        const minTotal = parseInt(document.getElementById("minMoleculeSum")?.value || "3", 10);
        const selected = orderedHeaders.filter(h => selectedHeaders[h].checked);
    
        const moleculeSums = {};
        dataRows.forEach(({ headers: rowHeaders, values }) => {
            selected.forEach(header => {
                const idx = rowHeaders.indexOf(header);
                if (idx >= 0) {
                    const value = parseFloat(values[idx]) || 0;
                    moleculeSums[header] = (moleculeSums[header] || 0) + value;
                }
            });
        });
    
        // Separate headers based on threshold
        const includedHeaders = selected.filter(h => moleculeSums[h] >= minTotal);
        const excludedHeaders = selected.filter(h => moleculeSums[h] < minTotal);
    
        const headers = ["Timestep", ...includedHeaders];
    
        const processedData = dataRows.map(({ headers: rowHeaders, values }) => {
            const rowData = {};
            rowData["Timestep"] = values[0] || "0";
    
            includedHeaders.forEach(header => {
                const idx = rowHeaders.indexOf(header);
                rowData[header] = idx >= 0 && idx < values.length ? values[idx] : "0";
            });
    
            return rowData;
        });
    
        // Show excluded warning
        const warningEl = document.getElementById("excludedWarning");
        if (excludedHeaders.length > 0) {
            warningEl.innerText = `⚠️ Molecules excluded due to low total count (< ${minTotal}): ${excludedHeaders.join(", ")}`;
        } else {
            warningEl.innerText = "";
        }
    
        generateExcel(headers, processedData);
    }
    

    function generateExcel(headers, data) {
        const ws = XLSX.utils.json_to_sheet(data, { header: headers });

        // Freeze the first row (headers)
        ws["!autofilter"] = { ref: `A1:${String.fromCharCode(65 + headers.length - 1)}1` };
        ws["!rows"] = [{ hidden: false }]; // Ensure headers are visible

        // Auto-adjust column widths
        ws["!cols"] = headers.map(header => ({ wch: header.length + 5 }));

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Data");

        XLSX.writeFile(wb, "Converted_File.xlsx");
    }
});
