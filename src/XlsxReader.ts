import JSZip from 'jszip';

async function getWorksheetNames(input: Blob | ArrayBuffer): Promise<string[]> {
    const zip = new JSZip();
    const buffer = input instanceof Blob ? await input.arrayBuffer() : input; // Convert Blob to ArrayBuffer if needed
    const zipContent = await zip.loadAsync(buffer);

    const workbookXML = await zipContent.file('xl/workbook.xml')?.async('string');
    if (!workbookXML) throw new Error('Workbook metadata not found.');
    //console.log("XML:",workbookXML)
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


const getColumnIndex = (cellRef: string): number => {
    const match = cellRef.match(/[A-Z]+/);
    if (!match) return -1;
    const letters = match[0];
    let colNum = 0;
    for (let i = 0; i < letters.length; i++) {
        colNum = colNum * 26 + (letters.charCodeAt(i) - 64);
    }
    return colNum - 1; // zero-based index
};

const parseSheetData = (rows: HTMLCollectionOf<Element>, sharedStrings: string[]): any[][] => {
    const sheetData: any[][] = [];

    // take the first row as header
    const headerRow = rows[0];
    const headerCells = headerRow.getElementsByTagName('c');
    const expectedColumnCount = headerCells.length;
    const header: string[] = [];
    for (let j = 0; j < headerCells.length; j++) {
        const cell = headerCells[j];
        const cellRef = cell.getAttribute('r');
        if (!cellRef) continue;
        const colIndex = getColumnIndex(cellRef);
        const type = cell.getAttribute('t');
        const valueElement = cell.getElementsByTagName('v')[0];
        let value: any = valueElement ? valueElement.textContent : null;
        if (type === 's' && value !== null) {
            const sharedIndex = parseInt(value, 10);
            value = sharedStrings[sharedIndex] || null;
        } else if (type === 'inlineStr') {
            const inlineStrElement = cell.getElementsByTagName('t')[0];
            value = inlineStrElement ? inlineStrElement.textContent : null;
        }
        if (colIndex >= 0 && colIndex < expectedColumnCount) {
            header[colIndex] = value;
        }
    }
    console.log("Header:", header)
    sheetData.push(header);

    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const rowData: any[] = new Array(expectedColumnCount).fill(null);
        const cells = row.getElementsByTagName('c');

        for (let j = 0; j < cells.length; j++) {
            const cell = cells[j];
            const cellRef = cell.getAttribute('r');
            if (!cellRef) continue;

            const colIndex = getColumnIndex(cellRef);
            const type = cell.getAttribute('t');
            const valueElement = cell.getElementsByTagName('v')[0];
            let value: any = valueElement ? valueElement.textContent : null;

            if (type === 's' && value !== null) {
                const sharedIndex = parseInt(value, 10);
                value = sharedStrings[sharedIndex] || null;
            } else if (type === 'inlineStr') {
                const inlineStrElement = cell.getElementsByTagName('t')[0];
                value = inlineStrElement ? inlineStrElement.textContent : null;
            }

            if (colIndex >= 0 && colIndex < expectedColumnCount) {
                rowData[colIndex] = value;
            }
        }

        sheetData.push(rowData);
    }

    return sheetData;
};


async function readWorksheetByName(input: Blob | ArrayBuffer, sheetName: string): Promise<any[]> {
    const zip = new JSZip();
    const buffer = input instanceof Blob ? await input.arrayBuffer() : input; // Convert Blob to ArrayBuffer if needed
    const zipContent = await zip.loadAsync(buffer);
    //console.log(zipContent.files);
    const zipFiles = Object.keys(zipContent.files);
    //console.log("Files:", zipFiles)
    if (!zipFiles.includes('xl/workbook.xml')) throw new Error('Workbook metadata not found.');

    // Read workbook metadata to find the sheet ID corresponding to the name
    const workbookXML = await zipContent.file('xl/workbook.xml')?.async('string');
    if (!workbookXML) throw new Error('Workbook metadata not found.');
    //console.log("Workbook XML:", workbookXML)



    const parser = new DOMParser();

    const workbookDoc = parser.parseFromString(workbookXML, 'application/xml');
    const sheets = workbookDoc.getElementsByTagName('sheet');
    console.log("#sheets:", sheets.length)

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

    const sheetXML_ = zipContent.file(sheetPath)
    //console.log("Sheet XML:", sheetXML_)

    const sheetXML = await zipContent.file(sheetPath)?.async('text');
    if (!sheetXML) throw new Error(`Data for worksheet "${sheetName}" not found.`);
    //console.log("Sheet XML:", sheetXML, "Length:", sheetXML.length)

    // Parse shared strings for textual cell values
    const sharedStringsXML_ = zipContent.file('xl/sharedStrings.xml') //?.async('text');
    //console.log("Shared strings1:", sharedStringsXML_)
    const sharedStringsXML = sharedStringsXML_ ? await sharedStringsXML_.async('text') : null;
    //console.log("Shared strings2:", sharedStringsXML)
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

    const sheetData: any[] = parseSheetData(rows, sharedStrings);

    return sheetData;
}

export { getWorksheetNames, readWorksheetByName };