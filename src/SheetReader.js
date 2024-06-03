import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import './SheetReader.css';

const SheetReader = () => {
  const [tables, setTables] = useState([]);

  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      // informasi merge cells
      const merges = worksheet['!merges'] || [];
      console.log(merges);

      // konvert seluruh sheet ke JSON & skip 9 baris pertama
      const jsonSheet = XLSX.utils.sheet_to_json(worksheet, { header: 1, blankrows: false }).slice(9);

      // // sel yang di-merge
      // merges.forEach((merge) => {
      //   const startRow = merge.s.r;
      //   const startCol = merge.s.c;
      //   const endRow = merge.e.r;
      //   const endCol = merge.e.c;

      //   const startCell = worksheet[XLSX.utils.encode_cell({ r: startRow, c: startCol })];
      //   if (!startCell) return;

      //   const value = startCell.v;
      //   for (let R = startRow; R <= endRow; R++) {
      //     for (let C = startCol; C <= endCol; C++) {
      //       console.log('Cell ' + R + C);
      //       const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
      //       if (!worksheet[cellAddress]) worksheet[cellAddress] = { v: value };
      //     }
      //   }
      // });

      console.log(jsonSheet);

      const tables = [];
      let currentTable = [];
      let currentTitle = '';

      jsonSheet.forEach((row) => {
        // memulai tabel baru jika ditemukan judul baru yang sesuai pola dan tidak sama dengan 'TOTAL'
        // pola seperti A1A, A1B, A2, dst
        // console.log(row);
        if (row.length > 1 && typeof row[1] === 'string' && /^[A-Z0-9]+$/.test(row[1]) && row[1] !== 'TOTAL') {
          if (currentTable.length > 0) {
            tables.push({ title: currentTitle, data: currentTable });
            currentTable = [];
          }
          const kode = row[1];
          const judul = row[2];
          currentTitle = `${kode} - ${judul}`; // format judul (Kode - Keterangan)
        } else if (
          row.length > 0 &&
          row[2] === "kegiatan ini tidak diperhitungkan dalam sks kelebihan beban kerja, namun diperhitungkan dalam perhitungan lainnya"
        ) {
          return;  //skip baris kalau memiliki deskripsi di atas
        } else {
          // tambahkan baris ke tabel saat ini
          currentTable.push(row);
        }
      });

      // taambahkan tabel terakhir jika ada
      if (currentTable.length > 0) {
        tables.push({ title: currentTitle, data: currentTable });
      }

      setTables(tables);
    };
    reader.readAsArrayBuffer(file);
  };

  return (
    <div>
      <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
      {tables.map((table, index) => (
        <div key={index}>
          <h3>{table.title}</h3>
          <table border="1">
            <thead>
              <tr>
                {table.data[0].map((header, i) => (
                  <th key={i}>{header}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {table.data.slice(1).map((row, i) => (
                <tr key={i}>
                  {row.map((cell, j) => (
                    <td key={j}>{cell}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ))}
    </div>
  );
};

export default SheetReader;
