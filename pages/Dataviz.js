document.getElementById('fileInput').addEventListener('change', handleFileUpload);
document.getElementById('generateReport').addEventListener('click', generateReport);
const saveBtn = document.getElementById('saveBtn');
const downloadBtn = document.getElementById('downloadBtn');
let hot;
let originalData;

// Initial dummy data for the report
const dummyData = [
    ['Category', 'Value 1', 'Value 2', 'Value 3'],
    ['A', 10, 20, 30],
    ['B', 15, 25, 35],
    ['C', 20, 30, 40]
];

generateDummyReport(dummyData);

function handleFileUpload(event) {
    const file = event.target.files[0];
    const fileInfo = document.getElementById('fileInfo');
    if (file) {
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
            loadSpreadsheet(worksheet);
        };
        reader.readAsArrayBuffer(file);
        fileInfo.textContent = `File uploaded: ${file.name}`;
        fileInfo.style.display = 'block';
    } else {
        fileInfo.textContent = 'No file uploaded.';
        fileInfo.style.display = 'block';
    }
}


function loadSpreadsheet(data) {
    const container = document.getElementById('spreadsheet');
    if (hot) {
        hot.destroy();
        hot = null;
    }
    hot = new Handsontable(container, {
        data: data,
        rowHeaders: true,
        colHeaders: true,
        contextMenu: true,
        height:400,
        licenseKey: 'non-commercial-and-evaluation'
    });
    document.getElementById('spreadsheet').style.display = 'hidden';
}

function showUploadedData(data) {
    const container = document.getElementById('excelData');
    if (hot) {
        hot.destroy();
    }
    hot = new Handsontable(container, {
        data: data,
        rowHeaders: true,
        colHeaders: true,
        contextMenu: true,
        licenseKey: 'non-commercial-and-evaluation'
    });
    document.getElementById('uploadedData').style.display = 'block';
}
function generateReport() {
    const data = hot.getData();
    if (data.length > 1) {
        const headers = data[0];
        const values = data.slice(1);
        
        const categories = headers.slice(1); // assuming first column is category
        const series = [];

        values.forEach(row => {
            series.push({
                category: row[0],
                values: row.slice(1)
            });
        });

        generateCharts(categories, series);
    }
}

function generateDummyReport(data) {
    const headers = data[0];
    const values = data.slice(1);
    
    const categories = headers.slice(1);
    const series = values.map(row => ({
        category: row[0],
        values: row.slice(1)
    }));

    generateCharts(categories, series);
}

function generateCharts(categories, series) {
    generateBarChart(categories, series);
    generatePieChart(categories, series);
    generateLineChart(categories, series);
    generateScatterPlot(categories, series);
    generateHistogram(series);
    generateAreaChart(categories, series);
}

function generateBarChart(categories, series) {
    const traces = series.map(s => ({
        x: categories,
        y: s.values,
        name: s.category,
        type: 'bar'
    }));

    const layout = {
        title: 'Bar Chart'
    };
    Plotly.newPlot('barChart', traces, layout);
}

function generatePieChart(categories, series) {
    const trace = {
        labels: series.map(s => s.category),
        values: series.map(s => s.values.reduce((a, b) => a + b, 0)),
        type: 'pie'
    };
    const layout = {
        title: 'Pie Chart'
    };
    Plotly.newPlot('pieChart', [trace], layout);
}

function generateLineChart(categories, series) {
    const traces = series.map(s => ({
        x: categories,
        y: s.values,
        name: s.category,
        type: 'scatter',
        mode: 'lines+markers'
    }));

    const layout = {
        title: 'Line Chart'
    };
    Plotly.newPlot('lineChart', traces, layout);
}

function generateScatterPlot(categories, series) {
    const traces = series.map(s => ({
        x: categories,
        y: s.values,
        mode: 'markers',
        type: 'scatter',
        name: s.category
    }));

    const layout = {
        title: 'Scatter Plot'
    };
    Plotly.newPlot('scatterPlot', traces, layout);
}

function generateHistogram(series) {
    const data = series.flatMap(s => s.values);
    const trace = {
        x: data,
        type: 'histogram',
    };

    const layout = {
        title: 'Histogram'
    };
    Plotly.newPlot('histogram', [trace], layout);
}

function generateAreaChart(categories, series) {
    const traces = series.map(s => ({
        x: categories,
        y: s.values,
        fill: 'tozeroy',
        type: 'scatter',
        mode: 'none',
        name: s.category
    }));

    const layout = {
        title: 'Area Chart'
    };
    Plotly.newPlot('areaChart', traces, layout);
}
