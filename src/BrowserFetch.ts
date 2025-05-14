import axios from 'axios';

async function fetchExcelFile(url: string): Promise<Blob> {
    try {
        const response = await axios.get(url, {
            responseType: 'blob', // Fetch as Blob for browser compatibility
            validateStatus: (status:number) => status >= 200 && status < 300,
        });
        return response.data;
    } catch (error: any) {
        if (axios.isAxiosError(error)) {
            const statusCode = error.response?.status;
            const statusText = error.response?.statusText;
            throw new Error(`Failed to fetch file: ${statusCode} ${statusText}`);
        }
        throw new Error(`Unknown error: ${error.message}`);
    }
}

export { fetchExcelFile };

/* Example browser usage:
import { fetchExcelFile } from './BrowserFetch';
const url = 'path/to/your/excel/file.xlsx';
try:
    const xlsx = await fetchExcelFile(url)
    console.log('File fetched successfully:', xlsx);
    // You can now use the blob to read the Excel file
} catch (error) {
    console.error('Error fetching file:', error);
}
*/
