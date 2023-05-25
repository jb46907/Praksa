const express = require('express');
const bodyParser = require('body-parser');
const port = 3000;
const app = express();
const fs = require('fs');
const ExcelJS = require('exceljs');



app.use(bodyParser.json());

async function createExcel() {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Nalog za isplatu');
  const imageFilePath = 'logo.png';
  const imageFile = fs.readFileSync(imageFilePath);

  const imageId = workbook.addImage({
    buffer: imageFile,
    extension: 'png',
  });

  const imageDimensions = {
    width: 20,
    height: 20, 
  };
  const imagePosition = {
    col: 1,
    row: 1,
  };


  worksheet.addImage(imageId, {
    tl: imagePosition,
    br: {
      col: imagePosition.col + imageDimensions.width / 10,
      row: imagePosition.row + imageDimensions.height / 10,
    },
  });

  const data = require('./data.json');

  worksheet.mergeCells('A5:C5');
  const subject = data.subjects[0];
  const mergedCell = worksheet.getCell('A5');
  mergedCell.value = subject ? `Predmet: ${subject.name} (${subject.code})` : 'Predmet: N/A';
  
  worksheet.mergeCells('A6:I11');
  const cellA6 = worksheet.getCell('A6');
  cellA6.value = {
    richText: [
      { text: '                                                         NALOG ZA ISPLATU\n', font: { bold: true, size: 18 } },
      {
        text: 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.',
      },
    ],
  };
  cellA6.alignment = {
    vertical: 'middle',
    horizontal: 'left',
    wrapText: true,
  };

 worksheet.mergeCells('A12:B12');
 worksheet.mergeCells('H12:I12');

 worksheet.getCell('A12').value = 'Katedra';
 worksheet.getCell('C12').value = 'Studij';
 worksheet.getCell('D12').value = 'ak. god.';
 worksheet.getCell('E12').value = 'stud. god.';
 worksheet.getCell('F12').value = 'pocetak turnusa';
 worksheet.getCell('G12').value = 'kraj turnusa';
 worksheet.getCell('H12').value = 'broj sati predviden programom';

 data.katedre.forEach((katedra, index) => {
   const rowNumber = 13 + index;
   const row = worksheet.getRow(rowNumber);

   worksheet.mergeCells(`A${rowNumber}:B${rowNumber}`);
   worksheet.getCell(`A${rowNumber}`).value = katedra['ime'];
   worksheet.getCell(`C${rowNumber}`).value = katedra['studij'];
   worksheet.getCell(`D${rowNumber}`).value = katedra['ak. god.'];
   worksheet.getCell(`E${rowNumber}`).value = katedra['stud. god.'];
   worksheet.getCell(`F${rowNumber}`).value = katedra['pocetak turnusa'];
   worksheet.getCell(`G${rowNumber}`).value = katedra['kraj turnusa'];
   worksheet.mergeCells(`H${rowNumber}:I${rowNumber}`);
   worksheet.getCell(`H${rowNumber}`).value =
     ' P:' +
     katedra['pred'] +
     ' S:' +
     +katedra['sem'] +
     ' V:' +
     +katedra['vjez'];

   row.alignment = { horizontal: 'left' };
   worksheet.getCell(`H${rowNumber}`).alignment = { horizontal: 'center' };
   worksheet.getCell(`A${rowNumber}`).alignment = { horizontal: 'center' };

   worksheet.getRow(rowNumber).eachCell((cell) => {
     cell.border = {
       top: { style: 'medium' },
       left: { style: 'thin' },
       bottom: { style: 'medium' },
       right: { style: 'medium' },
     };
     cell.alignment = { horizontal: 'center', vertical: 'middle' };
   });
  });
  
  const mergeRanges = [
    'A15:A16',
    'B15:B16',
    'C15:C16',
    'D15:D16',
    'E15:G15',
    'H15:H16',
    'I15:I16',
    'J15:J16',
    'K15:M15',
    'N15:N16',
  ];
  mergeRanges.forEach((range) => {
    worksheet.mergeCells(range);
  });

  worksheet.getCell('A15').value = 'Redni broj';
  worksheet.getCell('B15').value = 'Ime i Prezime';
  worksheet.getCell('C15').value = 'Zvanje';
  worksheet.getCell('D15').value = 'Status';
  worksheet.getCell('E15').value = 'Sati Nastave';
  worksheet.getCell('E16').value = 'pred';
  worksheet.getCell('F16').value = 'sem';
  worksheet.getCell('G16').value = 'vjez';
  worksheet.getCell('H15').value = 'Bruto satnica predavanja (EUR)';
  worksheet.getCell('I15').value = 'Bruto satnica seminari (EUR)';
  worksheet.getCell('J15').value = 'Bruto satnica vjezbe (EUR)';
  worksheet.getCell('K15').value = 'Bruto iznos';
  worksheet.getCell('K16').value = 'pred';
  worksheet.getCell('L16').value = 'sem';
  worksheet.getCell('M16').value = 'vjez';
  worksheet.getCell('N15').value = 'Ukupno za isplatu (EUR)';

  const headerRows = [15, 16, 12];

  headerRows.forEach((rowNumber) => {
    const row = worksheet.getRow(rowNumber);
    row.eachCell((cell) => {
      cell.font = { bold: true };
      cell.alignment = {
        horizontal: 'center',
        vertical: 'middle',
        wrapText: true,
      };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'E7E7E7' },
      };
      cell.border = {
        top: { style: 'medium' },
        left: { style: 'medium' },
        bottom: { style: 'medium' },
        right: { style: 'medium' },
      };
    });
  });

  const cellHeights = {
    A12: 40,
    A16: 55,
    E16: 55,
    F16: 55,
  };
  for (const [cellRef, height] of Object.entries(cellHeights)) {
    const cell = worksheet.getCell(cellRef);
    worksheet.getRow(cell.row).height = height;
  }  

  const startRow = 17;

  if (Array.isArray(data.profesori) && data.profesori.length > 0) {
    data.profesori.forEach((professor, index) => {
      const row = startRow + index;

      console.log('Processing professor:', professor);
  
      worksheet.getCell(`A${row}`).value = index + 1;
      worksheet.getCell(`B${row}`).value = professor['NastavnikSuradnikNaziv'];
      worksheet.getCell(`C${row}`).value = professor['Titula'];
      worksheet.getCell(`D${row}`).value = professor['NazivNastavnikStatus'];
      worksheet.getCell(`E${row}`).value = professor['PlaniraniSatiPredavanja'];
      worksheet.getCell(`F${row}`).value = professor['PlaniraniSatiSeminari'];
      worksheet.getCell(`G${row}`).value = professor['PlaniraniSatiVjezbe'];
      worksheet.getCell(`H${row}`).value = professor['NormaPlaniraniSatiPredavanja'];
      worksheet.getCell(`I${row}`).value = professor['NormaPlaniraniSatiSeminari'];
      worksheet.getCell(`J${row}`).value = professor['NormaPlaniraniSatiVjezbe'];
  
      const calculateCellValue = (normaKey, planiraniKey) => {
        const norma = professor[normaKey];
        const planirani = professor[planiraniKey];
        return norma * planirani;
      };
  
      worksheet.getCell(`K${row}`).value = calculateCellValue('NormaPlaniraniSatiPredavanja', 'PlaniraniSatiPredavanja');
      worksheet.getCell(`L${row}`).value = calculateCellValue('NormaPlaniraniSatiSeminari', 'PlaniraniSatiSeminari');
      worksheet.getCell(`M${row}`).value = calculateCellValue('NormaPlaniraniSatiVjezbe', 'PlaniraniSatiVjezbe');
  
      const sumCellValues = ['K', 'L', 'M'].reduce((sum, cellKey) => {
        return sum + worksheet.getCell(`${cellKey}${row}`).value;
      }, 0);
  
      worksheet.getCell(`N${row}`).value = sumCellValues;
  
      // Apply specific formatting to cells
      const rowCells = worksheet.getRow(row).eachCell((cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      });
    });
  } else {
    console.log('No data found in the "profesori" array.');
  }
  
  
  let sumBrutoPredavanja = 0;
  let sumBrutoSeminara = 0;
  let sumBrutoVjezbe = 0;
  let sumSatiPredavanja = 0;
  let sumSatiSeminara = 0;
  let sumSatiVjezbe = 0;
  let sumTotal = 0;
  
  const rowNumber = data.profesori.length + 17;
  
  worksheet.mergeCells(`A${rowNumber}:C${rowNumber}`);
  worksheet.getCell(`A${rowNumber}`).value = 'UKUPNO';
  
  const calculateSumFormula = (column, startRow, endRow) => ({
    formula: `SUM(${column}${startRow}:${column}${endRow})`,
    result: 0,
  });
  
  const updateCell = (column, cellRef, value) => {
    const cell = worksheet.getCell(cellRef);
    cell.value = value;
  
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  };
  
  const updateFormulaCell = (column, cellRef, startRow, endRow, resultVariable) => {
    const formula = calculateSumFormula(column, startRow, endRow);
    formula.result = resultVariable;
    updateCell(column, cellRef, formula);
  };
  
  updateFormulaCell('E', `E${rowNumber}`, 17, rowNumber - 1, sumSatiPredavanja);
  updateFormulaCell('F', `F${rowNumber}`, 17, rowNumber - 1, sumSatiSeminara);
  updateFormulaCell('G', `G${rowNumber}`, 17, rowNumber - 1, sumSatiVjezbe);
  
  updateFormulaCell('K', `K${rowNumber}`, 17, rowNumber - 1, sumBrutoPredavanja);
  updateFormulaCell('L', `L${rowNumber}`, 17, rowNumber - 1, sumBrutoSeminara);
  updateFormulaCell('M', `M${rowNumber}`, 17, rowNumber - 1, sumBrutoVjezbe);
  
  const totalSumFormula = calculateSumFormula('N', 17, rowNumber - 1);
  totalSumFormula.result = sumTotal;
  updateCell('N', `N${rowNumber}`, totalSumFormula);
  

  const row = worksheet.getRow(rowNumber);

  row.eachCell((cell) => {
    cell.border = {
      top: { style: 'medium' },
      left: { style: 'medium' },
      bottom: { style: 'medium' },
      right: { style: 'medium' },
    };
    cell.alignment = { horizontal: 'center' };
    cell.font = { bold: true };
  });

  const borderStyle = {
    top: { style: 'medium' },
    left: { style: 'medium' },
    bottom: { style: 'medium' },
    right: { style: 'medium' },
  };

  worksheet.getCell(`D${rowNumber}`).border = borderStyle;
  worksheet.getCell(`H${rowNumber}`).border = borderStyle;
  worksheet.getCell(`I${rowNumber}`).border = borderStyle;
  worksheet.getCell(`J${rowNumber}`).border = borderStyle;

  const columnWidths = [6, 18.43, 21.14, 21.14, 6.14, 7.86, 8.14, 10.14, 10, 10.14];
  worksheet.columns.forEach((column, index) => {
    column.width = columnWidths[index] || 8.43;
  });

  const mergeCellsData = [
    { ref: 'A28:C29', value: 'Prodekanica za nastavu i studentska pitanja\n', name: 'ImePrezime', index: 0 },
    { ref: 'A34:C35', value: 'Prodekan za financije i upravljanje\n', name: 'ImePrezime', index: 1 },
    { ref: 'J34:L35', value: 'Dekan\n', name: 'ImePrezime', index: 2 },
  ];
  
  mergeCellsData.forEach(({ ref, value, name, index }) => {
    worksheet.mergeCells(ref);
    const cell = worksheet.getCell(ref.split(':')[0]);
    cell.value = {
      richText: [
        { text: value },
        { text: `Prof. dr. sc. ${data.dekani[index][name]}` },
      ],
    };
    cell.alignment = {
      vertical: 'middle',
      horizontal: 'left',
      wrapText: true,
    };
  });
  

  const fileName = 'Nalog za isplatu.xlsx';
  await workbook.xlsx.writeFile(fileName);
  console.log(`Excel file created successfully: ${fileName}.`);
}

app.get('/professor/:professorId', async (req, res) => {
  try {
    const professorId = parseInt(req.params.professorId);
    const data = await fs.promises.readFile('data.json', 'utf8');
    const professors = JSON.parse(data).professors;
    
    const professor = professors.find((prof) => prof.id === professorId);
    if (professor) {
      res.send(professor);
    } else {
      res.status(404).send('Professor not found');
    }
  } catch (error) {
    res.status(500).send('An error occurred while retrieving professor data');
  }
});


app.get('/katedra/:katedraId', async (req, res) => {
  try {
    const katedraId = parseInt(req.params.katedraId);
    const data = await fs.promises.readFile('data.json', 'utf8');
    const katedre = JSON.parse(data).katedre;

    const katedra = katedre.find((katedra) => katedra.id === katedraId);
    if (katedra) {
      res.send(katedra);
    } else {
      res.status(404).send('Katedra not found');
    }
  } catch (error) {
    res.status(500).send('An error occurred while retrieving katedra data');
  }
});


app.post('/create', async (req, res) => {
  try {
    const data = await fs.promises.readFile('data.json', 'utf8');
    const jsonData = JSON.parse(data);

    await createExcel(jsonData);

    res.send('Excel file created successfully.');
  } catch (error) {
    console.error('Error:', error);
    res.status(500).send('An error occurred while generating the Excel file.');
  }
});

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
