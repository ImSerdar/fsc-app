const axios = require('axios');
const ExcelJS = require('exceljs');
const stream = require('stream');

function calculateFuelSurchargePercentage(price) {
    let surchargePercentage = 0;

    if (price >= 180.9 && price <= 181.8) {
        surchargePercentage = 26;
    } else if (price > 181.8) {
        surchargePercentage = 26 + (Math.floor(price - 180.9) * 0.25);  
    } else if (price < 180.9) {
        surchargePercentage = (Math.floor((price - 69.9) / 2) * 0.25) + 12;
    }

    return surchargePercentage.toFixed(2);  // Round to two decimal places
    
}

module.exports = async function (context, req) {
    context.log('JavaScript HTTP trigger function processed a request.');

    const url = 'https://charting.kalibrate.com/WPPS/Diesel/Retail%20(Incl.%20Tax)/WEEKLY/2023/Diesel_Retail%20(Incl.%20Tax)_WEEKLY_2023.xlsx';

    try {
        // Download Excel file from the URL
        const response = await axios.get(url, { responseType: 'arraybuffer' });

        // Create a readable stream from the downloaded data
        const readableStream = new stream.Readable();
        readableStream.push(response.data);
        readableStream.push(null);

        // Load workbook from stream
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.read(readableStream);

        // Get the first worksheet
        const worksheet = workbook.getWorksheet(1);

       

        // Build HTML table from Excel data
        let htmlTable = '<table>\n<tr><th>Effective Date</th><th>Price</th><th>LTL</th><th>TL</th></tr>\n';

        // Get the third row for "Week Ending" values
        const weekEndingRow = worksheet.getRow(3);
        let weekEndings = [];
        weekEndingRow.eachCell((cell, colNumber) => {
            if (colNumber > 1) { // Skip columns before the actual data
                let [month, day] = cell.text.split('/');
                let date = new Date(new Date().getFullYear(), month - 1, +day + 3);  // Assume current year, add 3 days for Friday
                let newWeekEnding = `${(date.getMonth() + 1).toString().padStart(2, '0')}/${date.getDate().toString().padStart(2, '0')}`;  // Format date as MM/DD
                weekEndings.push(newWeekEnding);
            }
        });

        // Find the row starting with "Vancouver" and iterate through it for prices
        let prices = [];
        let foundEmptyCell = false;

        worksheet.eachRow((row, rowNumber) => {
            if (row.getCell(1).text.startsWith('VANCOUVER*')) {
                row.eachCell((cell, colNumber) => {
                    if (colNumber > 1) {  // Skip columns before the actual data
                        if(cell.text === '' || cell.text == null) {
                            foundEmptyCell = true;
                        }
                        if (foundEmptyCell) {
                            return;  // Stop processing if an empty cell has been found
                        }
                        prices.push(cell.text);
                    }
                });
                if (foundEmptyCell) {
                    return;  // Stop processing further rows if an empty cell has been found
                }
            }
        });

        // Determine the starting index to get the last 6 rows
            let startIndex = prices.length - 6;
            if (startIndex < 0) startIndex = 0; 
        // Combine week endings and prices into the HTML table
        for (let i = startIndex; i < prices.length; i++) {
            if (prices[i]) { // Check if there is data in the prices field
                const price = parseFloat(prices[i]);  // Assume prices are in string format, convert to float
                const ltl = calculateFuelSurchargePercentage(price);  // Calculate LTL
                const tl = parseFloat(ltl) + 15;
                htmlTable += `<tr><td>${weekEndings[i]}</td><td>${prices[i]}</td><td>${ltl}%</td><td>${tl}%</td></tr>\n`;
            } else {
                break; // Stop the loop if there is no more data in the prices field
            }
        }

        htmlTable += '</table>'; // Close the HTML table

        // CSS for dark mode and centering the table
        
        // Combine CSS and HTML table
        //const htmlContent = css + htmlTable;


        const htmlStructure = `
                <!DOCTYPE html>
                <html lang="en">
                <head>
                    <meta charset="UTF-8">
                    <meta http-equiv="X-UA-Compatible" content="IE=edge">
                    <meta name="viewport" content="width=device-width, initial-scale=1.0">
                    <style>
                    body {
                        background-color: #2e2e2e !important;
                        color: #f5f5f5;
                        font-family: Arial, sans-serif;
                    }
                
                    table {
                        margin: 0 auto;
                        border-collapse: collapse;
                        width: 40em;
                        border: solid;
                    }
                
                    th {
                        background-color: #444;
                        padding: 10px;
                        text-align: center;
                        border: solid;
                    }
                
                    td {
                        padding: 10px;
                        text-align: left;
                        text-align: center;
                        border: solid;
                    }
                    </style>
                    <title>Table Data</title>
                </head>
                <body>
                    <h1>Weekly Average diesel fuel prices in Vancouver, BC</h1>
                    ${htmlTable}
                    
                    <a href='https://charting.kalibrate.com/WPPS/Diesel/Retail%20(Incl.%20Tax)/WEEKLY/2023/Diesel_Retail%20(Incl.%20Tax)_WEEKLY_2023.xlsx'> Source </a>
                </body>
                </html>
            `;


        // Print the HTML content to the console
        //context.log(htmlContent);

        context.res = {
            // status: 200, /* Defaults to 200 */
            headers: {
                'Content-Type': 'text/html'
            },
            body: htmlStructure
        };
    } catch (error) {
        context.log(error);

        context.res = {
            status: 500,
            body: "An error occurred while processing your request."
        };
    }
};