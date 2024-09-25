const propertiesReader = require('properties-reader');
const ExcelJS= require('exceljs');

const fs = require('fs');

// Define the log file path

let properties = propertiesReader("./config.properties");
const logFilePath = properties.get("logFilePath");
const excelFilePath = properties.get("excelFilePath");
const workbook = new ExcelJS.Workbook();

//new worksheet to the workbook
const Metrics_sheet = workbook.addWorksheet('Metrics');
//new worksheet to the workbook
const performanceMetrics_sheet = workbook.addWorksheet('Performance Sub Metrics');


function logg(content, severity = "INFO") {
    fs.appendFileSync(logFilePath, new Date().toISOString() + ` [${severity}] : ` + content + "\n");
    console.log(content)
}

function deleteLogAndExcelFile() {
    if (fs.existsSync(logFilePath)) {
        fs.unlinkSync(logFilePath); // This deletes the file
        logg("Existing Log file is deleted", "INFO");
    } else {
        logg("Existing Log file is not found", "INFO");
    }
    // Delete the old Excel file if it exists
    if (fs.existsSync(excelFilePath)) {
        fs.unlinkSync(excelFilePath);// This deletes the old Excel file
        logg('Existing Excel file deleted', "INFO");
    }
    else {
        logg('Existing Excel file is not found', "INFO");
    }
}

function compareReportsMetrics(from, to) {
    let metricCount = 0;
    let scoresData = [];
    const metricNames = [
        "performance",
        "accessibility",
        "best-practices",
        "seo"
    ];
    //Difference between the Metrics
    const calculationDifference = (from, to) => {
        from = from * 100;
        to = to * 100;
        return to - from;
    };
    //Percentage Calculation
    const calculationPercentageDifference = (from, to) => {
        const per = ((to - from) / from) * 100;
        return Math.round(per * 100) / 100;
    };

    //Audits Difference
    for (let metricObj in from["categories"]) {

        if (metricNames.includes(metricObj)) {
            if (from["categories"][metricObj].score != undefined && to["categories"][metricObj].score != undefined) {
                if (from["categories"][metricObj].score === "" || from["categories"][metricObj].score === null) {
                    logg(from["categories"][metricObj].title + " value is not present in previous file", "ERROR");
                }
                else{
                    logg(from["categories"][metricObj].title + " value is present in previous file", "INFO");
                }
                if (to["categories"][metricObj].score === "" || to["categories"][metricObj].score === null) {
                    logg(to["categories"][metricObj].title + " value is not present in recent file", "ERROR");
                }
                else{
                    logg(to["categories"][metricObj].title + " value is present in Recent file", "INFO");
                }

                const Difference = calculationDifference(
                    from["categories"][metricObj].score,
                    to["categories"][metricObj].score
                );
                const percentageDifference = calculationPercentageDifference(
                    from["categories"][metricObj].score,
                    to["categories"][metricObj].score
                );         

                metricCount++;
                // Adding key-value pairs as objects to the array
                let obj = {
                    'Sno:': metricCount, 'MetricName': from["categories"][metricObj].title, 'Previous Report Score': from["categories"][metricObj].score * 100,
                    'Current Report Score': to["categories"][metricObj].score * 100,
                    'Score Disparity(Calculated)': Difference, 'Percentage Variance': percentageDifference + '%',
                    'Standards': '90 to 100 ->Green\n50 to 89 ->Orange\n0 to 49 ->Red'
                };
                scoresData.push(obj);
                
            } else {
                if (from["categories"][metricObj].score === undefined) {
                    logg(from["categories"][metricObj].title + " value is undefined in previous file", "ERROR");
                } else {
                    logg(to["categories"][metricObj].title + " value is undefined in recent file", "ERROR");
                }
            }
        }
    }
    return scoresData;
}

function comparePerformanceMetrics(from, to) {
    try {
        let count = 0;
        let scoresDataTwo = [];
        const subMetrics = [
            "first-contentful-paint",
            "speed-index",
            "largest-contentful-paint",
            "total-blocking-time",
            "cumulative-layout-shift"
        ];
        //Performance child audits
        for (let auditObj in from["audits"]) {
            if (subMetrics.includes(auditObj)) {
                let latest = "";
                let Previous = "";
                if (from["audits"][auditObj].displayValue != undefined && to["audits"][auditObj].displayValue != undefined) {
                    if (from["audits"][auditObj].displayValue.toString() === "" || from["audits"][auditObj].displayValue.toString() === null) {
                        logg(from["audits"][auditObj].title + " value is not present in previous file", "ERROR");
                    }
                    else{
                        logg(from["audits"][auditObj].title + " value is present in previous file", "INFO");
                    }
                    if (to["audits"][auditObj].displayValue.toString() === "" || to["audits"][auditObj].displayValue.toString() === null) {
                        logg(to["audits"][auditObj].title + " value is not present in recent file", "ERROR");
                    }
                    else{
                        logg(to["audits"][auditObj].title + " value is present in Recent file", "INFO");
                    }
                    count++;
                    // for latest report display value
                    const removeSec = to["audits"][auditObj].displayValue.toString().replace(/\s+/g, '');
                    if (removeSec.includes('s') || removeSec.includes("ms")) {
                        const sec = /[ms]/gi;
                        latest = removeSec.replace(sec, '');
                    }
                    // for previous report display value
                    const removeMilliSec = from["audits"][auditObj].displayValue.toString().replace(/\s+/g, '');
                    if (removeMilliSec.includes('s') || removeMilliSec.includes("ms")) {
                        const sec = /[ms]/gi;
                        Previous = removeMilliSec.replace(sec, '');
                    }
                    // Time Difference
                    function timeVariation(latest, Previous) {
                        if (from["audits"][auditObj].title === "Total Blocking Time") {
                            if (latest < Previous) {
                                return (Previous - latest) / 1000;
                            }
                            if (latest > Previous) {
                                return (latest - Previous) / 1000;
                            }
                            return "no change";
                        }
                        if (latest < Previous) {
                            return Previous - latest;
                        }
                        if (latest > Previous) {
                            return latest - Previous;
                        }
                        return "no change";
                    }
                    let Variation = timeVariation(latest, Previous);


                    let objs = "";
                    switch (from["audits"][auditObj].title) {
                        case "First Contentful Paint":
                            objs = {
                                'Sno': count, 'Performance Sub-metrics': from["audits"][auditObj].title, 'Previous Report Score': from["audits"][auditObj].score * 100,
                                'Current Report Score': to["audits"][auditObj].score * 100,
                                'Duration Previous Report(In Sec)': Previous, 'Duration Current Report(In Sec)': latest,
                                'Variation(Calculated)': Variation,
                                'Standards': '0 to 1.8 seconds ->Green\n1.8 to 3 seconds ->Orange\n3+ seconds ->Red'
                            };
                            scoresDataTwo.push(objs);
                            break;
                        case "Speed Index":
                            objs = {
                                'Sno': count, 'Performance Sub-metrics': from["audits"][auditObj].title, 'Previous Report Score': from["audits"][auditObj].score * 100,
                                'Current Report Score': to["audits"][auditObj].score * 100,
                                'Duration Previous Report(In Sec)': Previous, 'Duration Current Report(In Sec)': latest,
                                'Variation(Calculated)': Variation,
                                'Standards': '0 to 3.4 seconds->Green\n3.4 to 5.8 seconds->Orange\n5.8+ seconds ->Red'
                            };
                            scoresDataTwo.push(objs);
                            break;
                        case "Largest Contentful Paint":
                            objs = {
                                'Sno': count, 'Performance Sub-metrics': from["audits"][auditObj].title, 'Previous Report Score': from["audits"][auditObj].score * 100,
                                'Current Report Score': to["audits"][auditObj].score * 100,
                                'Duration Previous Report(In Sec)': Previous, 'Duration Current Report(In Sec)': latest,
                                'Variation(Calculated)': Variation,
                                'Standards': '0 to 2.5 seconds->Green\n2.5 to 4 seconds->Orange\n4+ seconds->Red'
                            };
                            scoresDataTwo.push(objs);
                            break;
                        case "Total Blocking Time":
                            objs = {
                                'Sno': count, 'Performance Sub-metrics': from["audits"][auditObj].title, 'Previous Report Score': from["audits"][auditObj].score * 100,
                                'Current Report Score': to["audits"][auditObj].score * 100,
                                'Duration Previous Report(In Sec)': Previous / 1000, 'Duration Current Report(In Sec)': latest / 1000,
                                'Variation(Calculated)': Variation,
                                'Standards': '0 to 2 seconds->Green\n2 to 6 seconds->Orange\n6+-seconds>Red'
                            };
                            scoresDataTwo.push(objs);
                            break;
                        case "Cumulative Layout Shift":
                            objs = {
                                    'Sno': count, 'Performance Sub-metrics': from["audits"][auditObj].title, 'Previous Report Score': from["audits"][auditObj].score,
                                    'Current Report Score': to["audits"][auditObj].score,
                                    'Duration Previous Report(In Sec)': from["audits"][auditObj].displayValue, 'Duration Current Report(In Sec)': to["audits"][auditObj].displayValue,
                                    'Variation(Calculated)': (to["audits"][auditObj].displayValue -from["audits"][auditObj].displayValue)/10,
                                    'Standards': 'Note - this is a unitless metrics\n0 to 0.10 seconds->Green\n0.10 to 0.25 seconds->Orange\n0.25+-seconds>Red'
                                };
                            scoresDataTwo.push(objs);                           
                            break;
                    }
                } else {
                    if (from["audits"][auditObj].displayValue === undefined) {
                        logg(from["audits"][auditObj].title + " value is undefined in previous file", "ERROR");
                    } else {
                        logg(to["audits"][auditObj].title + " value is undefined in recent file", "ERROR");
                    }
                }
            }
        }
        return scoresDataTwo;
    } catch (err) {
        logg(err);
    }
}

// Apply color coding to the 'Score' column based on conditions
function applyColorCodingToPerformanceSubMetricsInSheet() {
    performanceMetrics_sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // Skip the header row    
        const scoreCell = row.getCell('Duration Current Report(In Sec)');
        const metricCell = row.getCell('Performance Sub-metrics');
        if (metricCell == "First Contentful Paint") {
            if (scoreCell.value > 0 && scoreCell.value <= 1.8) {
                scoreCell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: '41df46' }, // Green for 0 > score < 1.8
                };
            } else if (scoreCell.value >= 1.9 && scoreCell.value <= 3) {
                scoreCell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FF8000' }, // Orange for 1.9 >= score < 3
                };
            } else if (scoreCell.value > 3) {
                scoreCell.fill = {
                    type: 'pattern',
                    header: false,
                    pattern: 'solid',
                    fgColor: { argb: 'e3330c' }, // Red for score > 3
                };
            }
        }
        if (metricCell == "Speed Index") {
            if (scoreCell.value > 0 && scoreCell.value <= 3.4) {
                scoreCell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: '41df46' }, // Green for 0 >score < 3.4
                };
            } else if (scoreCell.value >= 3.5 && scoreCell.value <= 5.8) {
                scoreCell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FF8000' }, // Orange for 3.5 >= score < 5.8
                };
            } else if (scoreCell.value > 5.8) {
                scoreCell.fill = {
                    type: 'pattern',
                    header: false,
                    pattern: 'solid',
                    fgColor: { argb: 'e3330c' }, // Red for score > 5.8
                };
            }
        }
        if (metricCell == "Largest Contentful Paint") {
            if (scoreCell.value >= 0 && scoreCell.value <= 2.5) {
                scoreCell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: '41df46' }, // Green for 0 >= score < 2.5
                };
            } else if (scoreCell.value >= 2.6 && scoreCell.value <= 4) {
                scoreCell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FF8000' }, // Orange for 2.6 >= score < 4
                };
            } else if (scoreCell.value > 4) {
                scoreCell.fill = {
                    type: 'pattern',
                    header: false,
                    pattern: 'solid',
                    fgColor: { argb: 'e3330c' }, // Red for score < 4
                };
            }
        }
        if (metricCell == "Total Blocking Time") {
            if (scoreCell.value > 0 && scoreCell.value <= 2) {
                scoreCell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: '41df46' }, // Green for 0 > score < 2
                };
            } else if (scoreCell.value >= 2.1 && scoreCell.value <= 6) {
                scoreCell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FF8000' }, // Orange for 2.1 >= score < 6
                };
            } else if (scoreCell.value > 6) {
                scoreCell.fill = {
                    type: 'pattern',
                    header: false,
                    pattern: 'solid',
                    fgColor: { argb: 'e3330c' }, // Red for score > 6
                };
            }
        }
        if (metricCell == "Cumulative Layout Shift") {
            if (scoreCell.value > 0 && scoreCell.value <= 0.10) {
                scoreCell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: '41df46' }, // Green for 0 > score < 0.10
                };
            } else if (scoreCell.value >= 0.11 && scoreCell.value <= 0.25) {
                scoreCell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FF8000' }, // Orange for 0.11 >= score < 0.25
                };
            } else if (scoreCell.value > 0.25) {
                scoreCell.fill = {
                    type: 'pattern',
                    header: false,
                    pattern: 'solid',
                    fgColor: { argb: 'e3330c' }, // Red for score > 0.25
                };
            }
        }
    }
    )
};

function checkJsonFilePath(oldfilePath, newfilePath) {
    // Attempt to read the Previous file
    try {
        fs.readFileSync(oldfilePath, 'utf-8');
        logg(`File read successfully: ${oldfilePath}`);
        // Check if the file path present in config.properties
        if (oldfilePath == "") {
            logg("Previous json file path is not available in config.properties", "ERROR");
        } else {
            // Check if the file exists
            if (!fs.existsSync(oldfilePath))
                logg(`Previous json file is not exist: ${oldfilePath}`, "ERROR");
        }
    } catch (error) {
        logg(`Error reading the previous file: ${oldfilePath}` + error, "ERROR");
    }

    // Attempt to read the Recent file
    try {
        fs.readFileSync(newfilePath, 'utf-8');
        logg(`File read successfully: ${newfilePath}`);
        if (newfilePath == "") {
            logg("Recent json file path is not available in config.properties", "ERROR");
        } else {
            // Check if the file exists
            if (!fs.existsSync(newfilePath))
                logg(`Recent json file is not exist: ${newfilePath}`);
        }
    } catch (error) {
        logg(`Error reading the recent file: ${newfilePath}`, error);
    }
}

// Apply color coding to the 'Score' column based on Standards
function applyColorCodingToMetricsSheet() {
    // Add a new worksheet to the workbook
    Metrics_sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // Skip the header row  
        const scoreCell = row.getCell('Current Report Score');
   
        // Apply different colors based on the score value
        if (scoreCell.value >= 90 && scoreCell.value < 100) {
            scoreCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: '41df46' }, // Green for score >= 90
            };           
        } 
        
        else if (scoreCell.value >= 50 && scoreCell.value <= 89) {
            scoreCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF8000' }, // Orange for 50 <= score < 90
            };
        } else if (scoreCell.value >= 0 && scoreCell.value < 50) {
            scoreCell.fill = {
                type: 'pattern',
                header: false,
                pattern: 'solid',
                fgColor: { argb: 'e3330c' }, // Red for score < 49
            };
        }
        
    }
)
};


// Apply color coding to the 'Score Disparity' column based on the Previous report
function applyColorCodingToScoreDisparity() {
    // Add a new worksheet to the workbook
    Metrics_sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // Skip the header row        
        const variationCell = row.getCell('Score Disparity(Calculated)');       
        const previousReportScore = row.getCell('Previous Report Score');
        const currentReportScore = row.getCell('Current Report Score');

        // Apply different colors based on the score value
        if (currentReportScore > previousReportScore) {
            variationCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: '9dee91' }, 
            };           
        } 
        
        else if (currentReportScore == previousReportScore) {
            variationCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'f3864b' }, // Orange
            };
        } else if (currentReportScore < previousReportScore) {
            variationCell.fill = {
                type: 'pattern',
                header: false,
                pattern: 'solid',
                fgColor: { argb: 'ee4b4b' }, // Red 
            };
        }
   }
    
)};


// Apply color coding to the 'Score Variance ' column based on the Previous report
function applyColorCodingToScoreVariance() {
    // Add a new worksheet to the workbook
    performanceMetrics_sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // Skip the header row        
        const variationCell = row.getCell('Variation(Calculated)');
        const previousReportDuration = row.getCell('Duration Previous Report(In Sec)');
        const currentReportDuration = row.getCell('Duration Current Report(In Sec)');

        // Apply different colors based on the score value
        if (currentReportDuration < previousReportDuration) {
            variationCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: '9dee91' }, // Green 
            };           
        } 
        
        else if (currentReportDuration == previousReportDuration) {
            variationCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'f3864b' }, // Orange 
            };
        } else if (currentReportDuration > previousReportDuration) {
            variationCell.fill = {
                type: 'pattern',
                header: false,
                pattern: 'solid',
                fgColor: { argb: 'ee4b4b' }, // Red 
            };
        }              
    }
)};

function saveWorkBook() {
    workbook.xlsx.writeFile(excelFilePath)
        .then(() => {
            const filePath = excelFilePath;
            logg(`Excel file created successfully.: ${filePath}`);
        })
        .catch((err) => {
            logg(`Error creating Excel file : ` + err, "ERROR");
        });
    logg(`Log file created successfully.: ` + logFilePath);
}

function addMetricsInSheet() {
    try {
        // Define the columns for sheet one
        Metrics_sheet.columns = [
            { header: 'Sno', key: 'Sno:', width: 10 },
            { header: 'MetricName', key: 'MetricName', width: 30 },
            { header: 'Previous Report Score', key: 'Previous Report Score', width: 35 },
            { header: 'Current Report Score', key: 'Current Report Score', width: 35 },
            { header: 'Score Disparity(Calculated)', key: 'Score Disparity(Calculated)', width: 35 },
            { header: 'Percentage Variance', key: 'Percentage Variance', width: 35 },
            { header: 'Standards', key: 'Standards', width: 35 },
        ];

        // Format the header row
        Metrics_sheet.getRow(1).font = { bold: true, color: false };
        Metrics_sheet.getRow(1).alignment = { horizontal: 'left', wrapText: true };
        Metrics_sheet.getColumn(7).alignment = { horizontal: 'left', wrapText: true };       
       
               // Add rows from the object array
        compareReportsMetrics(require(properties.get("previousReportPath")), require(properties.get("currentReportPath"))).forEach((item) => {
            Metrics_sheet.addRow(item);
        });
    } catch (err) {
        logg(err);
    }
}

function addPerformanceSubMetricsInSheet() {
    try {
        // Define the columns
        performanceMetrics_sheet.columns = [
            { header: 'Sno', key: 'Sno', width: 10 },
            { header: 'Performance Sub-metrics', key: 'Performance Sub-metrics', width: 25 },
            { header: 'Previous Report Score', key: 'Previous Report Score', width: 25 },
            { header: 'Current Report Score', key: 'Current Report Score', width: 25 },
            { header: 'Duration Previous Report(In Sec)', key: 'Duration Previous Report(In Sec)', width: 25 },
            { header: 'Duration Current Report(In Sec)', key: 'Duration Current Report(In Sec)', width: 25 },
            { header: 'Variation(Calculated)', key: 'Variation(Calculated)', width: 25 },
            { header: 'Standards', key: 'Standards', width: 30 },
        ];
        // Add rows from the object array
        comparePerformanceMetrics(require(properties.get("previousReportPath")), require(properties.get("currentReportPath"))).forEach((item) => {
            performanceMetrics_sheet.addRow(item);      
        });
        // Format the header row
        performanceMetrics_sheet.getRow(1).font = { bold: true };
        performanceMetrics_sheet.getRow(1).alignment = { horizontal: 'left', wrapText: true };
        performanceMetrics_sheet.getColumn(8).alignment = { horizontal: 'left', wrapText: true };
    } catch (err) {
        logg(err);
    }
}

function addNotesToMetricsSheet(){
    
    let lastRowNumber = Metrics_sheet.lastRow ? Metrics_sheet.lastRow.number : 1; // Default to row 1 if no rows exist
    
    // Insert the Notes text into the next row (which will be the last)
    let lastRow = Metrics_sheet.getRow(lastRowNumber + 1);
    let lastSecondRow = Metrics_sheet.getRow(lastRowNumber + 2);
    let lastThirdRow = Metrics_sheet.getRow(lastRowNumber + 3);
    lastRow.getCell(1).value = '*** Current score results are color-coded based on the standards to be achieved.';
    lastSecondRow.getCell(1).value='Score Disparity indicates the comparison between the previous and current reports.';
    lastThirdRow.getCell(1).value='Red - Below standards; Orange - Needs improvement; Green - Meets standards.';
    
    // Merge columns A to D in that row (collapse columns)
    Metrics_sheet.mergeCells(`A${lastRowNumber + 1}:F${lastRowNumber + 1}`);
    Metrics_sheet.mergeCells(`A${lastRowNumber + 2}:F${lastRowNumber + 2}`);
    Metrics_sheet.mergeCells(`A${lastRowNumber + 3}:F${lastRowNumber + 3}`);
    };

function addNotesToPerformanceMetricsSheet(){
    
    let lastRowNumber = performanceMetrics_sheet.lastRow ? performanceMetrics_sheet.lastRow.number : 1; // Default to row 1 if no rows exist
        
    // Insert the Notes text into the next row (which will be the last)
    let lastRow = performanceMetrics_sheet.getRow(lastRowNumber + 1);
    let lastSecondRow = performanceMetrics_sheet.getRow(lastRowNumber + 2);
    let lastThirdRow = performanceMetrics_sheet.getRow(lastRowNumber + 3);
    lastRow.getCell(1).value = '*** Duration Current Report results are color-coded based on the standards to be achieved.';
    lastSecondRow.getCell(1).value='Variation indicates the comparison between the previous and current reports.';
    lastThirdRow.getCell(1).value='Red - Below standards; Orange - Needs improvement; Green - Meets standards.';
        
    // Merge columns A to C in that row (collapse columns)
    performanceMetrics_sheet.mergeCells(`A${lastRowNumber + 1}:G${lastRowNumber + 1}`);
    performanceMetrics_sheet.mergeCells(`A${lastRowNumber + 2}:G${lastRowNumber + 2}`);
    performanceMetrics_sheet.mergeCells(`A${lastRowNumber + 3}:G${lastRowNumber + 3}`);
    };
        





// Delete the old log file if it exists
deleteLogAndExcelFile();

//method for file paths and reading the files      
checkJsonFilePath(properties.get("previousReportPath"), properties.get("currentReportPath"));

//Add Metrics and Notes in the Excel Sheet
addMetricsInSheet();
addPerformanceSubMetricsInSheet();
addNotesToMetricsSheet();
addNotesToPerformanceMetricsSheet();

// Apply the common color coding to sheet one
applyColorCodingToMetricsSheet();
applyColorCodingToScoreDisparity();
applyColorCodingToScoreVariance();
applyColorCodingToPerformanceSubMetricsInSheet();

// Save the workbook to disk
saveWorkBook();