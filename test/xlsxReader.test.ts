import { getWorksheetNames, readWorksheetByName } from "../src/XlsxReader";
import { describe, expect } from "vitest"

import { beforeEach, vi } from 'vitest';
import fs from 'fs';
import path from 'path';

//const xlsxExample = path.resolve(__dirname, '../assets/example.xlsx');
const xlsxExample = 'example.xlsx'

// Shared variable to store first test results
let tables: string[] = [];
let xlsData: ArrayBuffer = new ArrayBuffer()

describe('xlsxReader', () => {
  beforeAll(() => {
    console.log("Test start")
    tables = []
    xlsData = new ArrayBuffer()
  }) 

  beforeEach(() => {
    vi.stubGlobal('fetch', async (url) => {
      // Extract filename from URL
      console.log("Mocking fetch to ", url)
      const filename = path.basename(url);
      const filePath = path.resolve(__dirname, '../assets', filename);

      if (fs.existsSync(filePath)) {
        const fileBuffer = fs.readFileSync(filePath);
        console.log("File loaded, length", fileBuffer.length)
        return new Response(fileBuffer.buffer, { status: 200, headers: { 'Content-Type': 'application/octet-stream' } });
      }

      return new Response(null, { status: 404 });
    });
  });

  describe('Load worksheet', () => {
    console.log("go ...")
    test('should load workseet and retrieve table names', async () => {
      try {
        const xls = await fetch(xlsxExample);
        xlsData = await xls.arrayBuffer();
        if (!(xlsData instanceof ArrayBuffer)) {
          throw new Error("Fetched data is not an ArrayBuffer");
        } else {
          console.log("ArrayBuffer length:", xlsData.byteLength)
        }
        const names = await getWorksheetNames(xlsData)
        console.log("Names:", names)
        expect(names.length).toEqual(2);
        tables = names
      } catch (error) {
        console.error("Error:", error)
        throw error;
      }
    });

    test('Second test: Iterate over TABLES and validate', () => {
      expect(tables.length).toBeGreaterThan(0); // Ensure first test ran

      tables.forEach(async (item) => {
        console.log("Item:", item)
        try {
          const tbl = await readWorksheetByName(xlsData, item)
          console.log("Table length:", tbl.length)
        }
        catch (error) {
          console.error("Error:", error)
          throw error;
        }
      });
    });
  })
})
