// script.js
async function convertirPDFaExcel() {
    const fileInput = document.getElementById('fileInput');
    const resultDiv = document.getElementById('result');
    const loadingDiv = document.getElementById('loading');

    resultDiv.innerHTML = '';
    loadingDiv.classList.remove('d-none');

    if (fileInput.files.length === 0) {
        alert("Por favor, selecciona un archivo.");
        loadingDiv.classList.add('d-none');
        return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = async function(e) {
        const typedarray = new Uint8Array(e.target.result);

        try {
            const pdf = await pdfjsLib.getDocument(typedarray).promise;
            let rows = [];
            for (let i = 1; i <= pdf.numPages; i++) {
                const page = await pdf.getPage(i);
                const textContent = await page.getTextContent();
                let pageRows = extractRowsFromPage(textContent.items);
                rows = rows.concat(pageRows);
            }

            const worksheet = XLSX.utils.aoa_to_sheet(rows);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "PDF Data");

            XLSX.writeFile(workbook, "convertido.xlsx");
            resultDiv.innerHTML = "<div class='alert alert-success'>El archivo PDF ha sido convertido a Excel con éxito.</div>";
        } catch (error) {
            resultDiv.innerHTML = "<div class='alert alert-danger'>Error al procesar el archivo PDF. Por favor, intenta de nuevo.</div>";
            console.error("Error al procesar el archivo:", error);
        } finally {
            loadingDiv.classList.add('d-none');
        }
    };

    reader.onerror = function() {
        resultDiv.innerHTML = "<div class='alert alert-danger'>Error al leer el archivo</div>";
        loadingDiv.classList.add('d-none');
    };

    reader.readAsArrayBuffer(file);
}

function extractRowsFromPage(items) {
    const rows = [];
    let currentRow = [];
    let currentY = null;

    items.forEach(item => {
        if (currentY === null || Math.abs(item.transform[5] - currentY) > 10) {
            if (currentRow.length > 0) {
                rows.push(currentRow);
                currentRow = [];
            }
            currentY = item.transform[5];
        }
        currentRow.push(item.str);
    });

    if (currentRow.length > 0) {
        rows.push(currentRow);
    }

    return rows;
}
