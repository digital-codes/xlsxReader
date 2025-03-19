import JSZip from 'jszip';

async function getWorksheetNames(input: Blob | ArrayBuffer): Promise<string[]> {
    const zip = new JSZip();
    const buffer = input instanceof Blob ? await input.arrayBuffer() : input; // Convert Blob to ArrayBuffer if needed
    const zipContent = await zip.loadAsync(buffer);

    const workbookXML = await zipContent.file('xl/workbook.xml')?.async('string');
    if (!workbookXML) throw new Error('Workbook metadata not found.');
    console.log("XML:",workbookXML)
    const parser = new DOMParser();
    const workbookDoc = parser.parseFromString(workbookXML, 'application/xml');
    const sheets = workbookDoc.getElementsByTagName('sheet');

    const sheetNames: string[] = [];
    for (let i = 0; i < sheets.length; i++) {
        const sheet = sheets[i];
        const name = sheet.getAttribute('name');
        if (name) sheetNames.push(name);
    }

    return sheetNames;
}

async function readWorksheetByName(input: Blob | ArrayBuffer, sheetName: string): Promise<any[]> {
    const zip = new JSZip();
    const buffer = input instanceof Blob ? await input.arrayBuffer() : input; // Convert Blob to ArrayBuffer if needed
    const zipContent = await zip.loadAsync(buffer);
    console.log(zipContent.files);

    // Read workbook metadata to find the sheet ID corresponding to the name
    const workbookXML = await zipContent.file('xl/workbook.xml')?.async('string');
    if (!workbookXML) throw new Error('Workbook metadata not found.');

    const parser = new DOMParser();
    const workbookDoc = parser.parseFromString(workbookXML, 'application/xml');
    const sheets = workbookDoc.getElementsByTagName('sheet');

    let sheetId: string | null = null;
    for (let i = 0; i < sheets.length; i++) {
        const sheet = sheets[i];
        if (sheet.getAttribute('name') === sheetName) {
            sheetId = sheet.getAttribute('sheetId');
            break;
        }
    }
    if (!sheetId) throw new Error(`Worksheet "${sheetName}" not found.`);

    console.log("Sheet ID:", sheetId)

    // Map sheet ID to file path (e.g., "sheet1.xml")
    const sheetPath = `xl/worksheets/sheet${sheetId}.xml`;
    console.log("Sheet Path:", sheetPath)

    const sheetXML = await zipContent.file(sheetPath)?.async('string'); 
    if (!sheetXML) throw new Error(`Data for worksheet "${sheetName}" not found.`);
    console.log("Sheet XML:", sheetXML, "Length:", sheetXML.length)


    // Parse shared strings for textual cell values
    const sharedStringsXML = await zipContent.file('xl/sharedStrings.xml')?.async('string');
    const sharedStrings: string[] = [];
    if (sharedStringsXML) {
        const sharedStringsDoc = parser.parseFromString(sharedStringsXML, 'application/xml');
        const siElements = sharedStringsDoc.getElementsByTagName('si');
        for (let i = 0; i < siElements.length; i++) {
            sharedStrings.push(siElements[i].textContent || '');
        }
    }

    // Parse the worksheet data
    const sheetDoc = parser.parseFromString(sheetXML, 'application/xml');
    const rows = sheetDoc.getElementsByTagName('row');
    const sheetData: any[] = [];

    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const rowData: any[] = [];
        const cells = row.getElementsByTagName('c'); // 'c' is the cell element in XLSX
        for (let j = 0; j < cells.length; j++) {
            const cell = cells[j];
            const type = cell.getAttribute('t'); // Check the cell type
            const valueElement = cell.getElementsByTagName('v')[0]; // 'v' is the value element

            let value: any = valueElement ? valueElement.textContent : null;

            if (type === 's' && value !== null) {
                // Shared string
                const sharedIndex = parseInt(value, 10);
                value = sharedStrings[sharedIndex] || null;
            } else if (type === 'inlineStr') {
                // Inline string
                const inlineStrElement = cell.getElementsByTagName('t')[0];
                value = inlineStrElement ? inlineStrElement.textContent : null;
            }

            rowData.push(value);
        }
        sheetData.push(rowData);
    }

    return sheetData;
}

export { getWorksheetNames, readWorksheetByName };