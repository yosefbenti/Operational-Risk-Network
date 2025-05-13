

const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');
const app = express();

// Configure Multer for file uploads
const upload = multer({ dest: 'uploads/' });

app.use(express.static('public'));
app.use(express.urlencoded({ extended: true }));
app.set('view engine', 'ejs');

let summaryData = {};

// Function to calculate summary from the workbook
function calculateSummary(workbook, type) {
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const jsonData = xlsx.utils.sheet_to_json(sheet);

  const scores = jsonData.map(row => {
    const key = `average_${type}`;
    return parseFloat(row[key]);
  }).filter(score => !isNaN(score));

  const total = scores.length;
  const above90 = scores.filter(score => score >= 90);
  const between85 = scores.filter(score => score >= 85 && score < 90);
  const below85 = scores.filter(score => score < 85);

  function average(arr) {
    return arr.length ? (arr.reduce((a, b) => a + b, 0) / arr.length).toFixed(2) : '0.00';
  }

  const result = {
    total,
    overallAverage: average(scores),
    categories: {
      above90Count: above90.length,
      between85Count: between85.length,
      below85Count: below85.length,
      above90Percent: total ? ((above90.length / total) * 100).toFixed(2) + '%' : '0%',
      between85Percent: total ? ((between85.length / total) * 100).toFixed(2) + '%' : '0%',
      below85Percent: total ? ((below85.length / total) * 100).toFixed(2) + '%' : '0%',
    },
    categoryAverages: {
      above90: average(above90),
      between85: average(between85),
      below85: average(below85),
    }
  };

  summaryData[type] = result;
}

// Route to render the upload form
app.get('/', (req, res) => {
  res.render('index', { data: summaryData });
});

// Helper to handle upload and cleanup
function handleUpload(req, res, type) {
  if (!req.file) {
    return res.status(400).send('No file uploaded.');
  }

  try {
    const workbook = xlsx.readFile(req.file.path);
    calculateSummary(workbook, type);
    fs.unlinkSync(req.file.path); // Delete the file after processing
    res.redirect('/');
  } catch (err) {
    console.error('Error processing file:', err);
    res.status(500).send('Error processing file. Please ensure the Excel format is correct.');
  }
}

// Route to handle City file upload
app.post('/upload-city', upload.single('excelFile'), (req, res) => {
  handleUpload(req, res, 'city');
});

// Route to handle Outline file upload
app.post('/upload-outline', upload.single('excelFile'), (req, res) => {
  handleUpload(req, res, 'outline');
});

// Route to handle Excel file download
app.post('/download', (req, res) => {
  const wb = xlsx.utils.book_new();
  let hasData = false;

  for (let type of ['city', 'outline']) {
    if (summaryData[type]) {
      const data = summaryData[type];
      const sheetData = [
        ['Category', 'Count', 'Average', 'Percentage'],
        ['Total Branches', data.total, data.overallAverage, '100%'],
        ['≥ 90%', data.categories.above90Count, data.categoryAverages.above90, data.categories.above90Percent],
        ['85% - 89.99%', data.categories.between85Count, data.categoryAverages.between85, data.categories.between85Percent],
        ['< 85%', data.categories.below85Count, data.categoryAverages.below85, data.categories.below85Percent],
      ];
      const ws = xlsx.utils.aoa_to_sheet(sheetData);
      xlsx.utils.book_append_sheet(wb, ws, `${type.charAt(0).toUpperCase() + type.slice(1)} Summary`);
      hasData = true;
    }
  }

  if (!hasData) {
    return res.status(400).send('No summary data available. Please upload a file first.');
  }

  const filename = 'summary_output.xlsx';
  const filepath = path.join(__dirname, filename);
  xlsx.writeFile(wb, filepath);

  res.download(filepath, filename, err => {
    if (err) {
      console.error('Download error:', err);
    } else {
      // Optionally remove the file after download
      fs.unlink(filepath, err => {
        if (err) console.error('Error deleting downloaded file:', err);
      });
    }
  });
});

// Start the server
app.listen(3000, () => {
  console.log('✅ Server started at http://localhost:3000');
});








/*const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');
const app = express();

// Configure Multer for file uploads
const upload = multer({ dest: 'uploads/' });

app.use(express.static('public'));
app.use(express.urlencoded({ extended: true }));
app.set('view engine', 'ejs');

let summaryData = {};

// Function to calculate summary from the workbook
function calculateSummary(workbook, type) {
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const jsonData = xlsx.utils.sheet_to_json(sheet);

  const scores = jsonData.map(row =>
    parseFloat(row[`average_${type}`])
  ).filter(score => !isNaN(score));

  const total = scores.length;
  const above90 = scores.filter(score => score >= 90);
  const between85 = scores.filter(score => score >= 85 && score < 90);
  const below85 = scores.filter(score => score < 85);

  function average(arr) {
    return arr.length ? (arr.reduce((a, b) => a + b, 0) / arr.length).toFixed(2) : '0.00';
  }

  const result = {
    total,
    overallAverage: average(scores),
    categories: {
      above90Count: above90.length,
      between85Count: between85.length,
      below85Count: below85.length,
      above90Percent: ((above90.length / total) * 100).toFixed(2) + '%',
      between85Percent: ((between85.length / total) * 100).toFixed(2) + '%',
      below85Percent: ((below85.length / total) * 100).toFixed(2) + '%',
    },
    categoryAverages: {
      above90: average(above90),
      between85: average(between85),
      below85: average(below85),
    }
  };

  summaryData[type] = result;
}

// Route to render the upload form
app.get('/', (req, res) => {
  res.render('index', { data: summaryData });
});

// Route to handle City file upload
app.post('/upload-city', upload.single('excelFile'), (req, res) => {
  if (!req.file) {
    return res.status(400).send('No file uploaded.');
  }

  const workbook = xlsx.readFile(req.file.path);
  calculateSummary(workbook, 'city');
  res.redirect('/');
});

// Route to handle Outline file upload
app.post('/upload-outline', upload.single('excelFile'), (req, res) => {
  if (!req.file) {
    return res.status(400).send('No file uploaded.');
  }

  const workbook = xlsx.readFile(req.file.path);
  calculateSummary(workbook, 'outline');
  res.redirect('/');
});

// Route to handle Excel file download
app.post('/download', (req, res) => {
  const wb = xlsx.utils.book_new();
  let hasData = false;

  for (let type of ['city', 'outline']) {
    if (summaryData[type]) {
      const data = summaryData[type];
      const sheetData = [
        ['Category', 'Count', 'Average', 'Percentage'],
        ['Total Branches', data.total, data.overallAverage, '100%'],
        ['≥ 90%', data.categories.above90Count, data.categoryAverages.above90, data.categories.above90Percent],
        ['85% - 89.99%', data.categories.between85Count, data.categoryAverages.between85, data.categories.between85Percent],
        ['< 85%', data.categories.below85Count, data.categoryAverages.below85, data.categories.below85Percent],
      ];
      const ws = xlsx.utils.aoa_to_sheet(sheetData);
      xlsx.utils.book_append_sheet(wb, ws, `${type.charAt(0).toUpperCase() + type.slice(1)} Summary`);
      hasData = true;
    }
  }

  if (!hasData) {
    return res.status(400).send('No summary data available. Please upload a file first.');
  }

  const filename = 'summary_output.xlsx';
  const filepath = path.join(__dirname, filename);
  xlsx.writeFile(wb, filepath);
  res.download(filepath, filename, (err) => {
    if (err) {
      console.error('Error downloading the file:', err);
    }
  });
});

// Start the server
app.listen(3000, () => {
  console.log('Server started at http://localhost:3000');
});*/












// const express = require('express');
// const multer = require('multer');
// const xlsx = require('xlsx');
// const path = require('path');
// const fs = require('fs');

// const app = express();
// app.use(express.static('public'));
// app.set('view engine', 'ejs');
// app.use(express.urlencoded({ extended: true }));

// const storage = multer.diskStorage({
//   destination: (req, file, cb) => cb(null, 'uploads/'),
//   filename: (req, file, cb) => cb(null, file.originalname)
// });
// const upload = multer({ storage });

// app.get('/', (req, res) => {
//   res.render('index', { data: null });
// });

// app.post('/upload', upload.single('excelFile'), (req, res) => {
//   const workbook = xlsx.readFile(req.file.path);
//   const sheet = workbook.Sheets[workbook.SheetNames[0]];
//   const jsonData = xlsx.utils.sheet_to_json(sheet);

//   const categories = {
//     above90: [],
//     between85: [],
//     below85: []
//   };

//   let totalScore = 0;

//   jsonData.forEach(entry => {
//     const score = parseFloat(entry.average_outline);
//     totalScore += score;

//     if (score >= 90) {
//       categories.above90.push(score);
//     } else if (score >= 85) {
//       categories.between85.push(score);
//     } else {
//       categories.below85.push(score);
//     }
//   });

//   const totalCount = jsonData.length;

//   const overallAverage = totalCount > 0 ? (totalScore / totalCount).toFixed(2) : 0;

//   const avgAbove90 = categories.above90.length > 0
//     ? (categories.above90.reduce((a, b) => a + b, 0) / categories.above90.length).toFixed(2)
//     : 'N/A';
//   const avgBetween85 = categories.between85.length > 0
//     ? (categories.between85.reduce((a, b) => a + b, 0) / categories.between85.length).toFixed(2)
//     : 'N/A';
//   const avgBelow85 = categories.below85.length > 0
//     ? (categories.below85.reduce((a, b) => a + b, 0) / categories.below85.length).toFixed(2)
//     : 'N/A';

//   const result = {
//     total: totalCount,
//     overallAverage,
//     categories: {
//       above90Count: categories.above90.length,
//       between85Count: categories.between85.length,
//       below85Count: categories.below85.length,
//       above90Percent: totalCount > 0 ? ((categories.above90.length / totalCount) * 100).toFixed(2) + '%' : '0%',
//       between85Percent: totalCount > 0 ? ((categories.between85.length / totalCount) * 100).toFixed(2) + '%' : '0%',
//       below85Percent: totalCount > 0 ? ((categories.below85.length / totalCount) * 100).toFixed(2) + '%' : '0%',
//     },
//     categoryAverages: {
//       above90: avgAbove90,
//       between85: avgBetween85,
//       below85: avgBelow85
//     }
//   };

//   res.render('index', { data: result });
// });

// app.post('/download', (req, res) => {
//   const data = JSON.parse(req.body.summaryData);

//   const sheetData = [
//     ['Category', 'Count', 'Average', 'Percentage'],
//     ['Total Branches', data.total, data.overallAverage, '100%'],
//     ['≥ 90%', data.categories.above90Count, data.categoryAverages.above90, data.categories.above90Percent],
//     ['85% - 89.99%', data.categories.between85Count, data.categoryAverages.between85, data.categories.between85Percent],
//     ['< 85%', data.categories.below85Count, data.categoryAverages.below85, data.categories.below85Percent]
//   ];

//   const wb = xlsx.utils.book_new();
//   const ws = xlsx.utils.aoa_to_sheet(sheetData);
//   xlsx.utils.book_append_sheet(wb, ws, 'Summary');

//   const filePath = path.join(__dirname, 'downloads', 'summary.xlsx');
//   fs.mkdirSync(path.join(__dirname, 'downloads'), { recursive: true });
//   xlsx.writeFile(wb, filePath);

//   res.download(filePath, 'summary.xlsx');
// });

// const PORT = 3000;
// app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
