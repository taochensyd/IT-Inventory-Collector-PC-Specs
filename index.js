const express = require('express');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3000;

app.use(bodyParser.json());

app.post('/api/v1/pcinfo', async (req, res) => {
    const data = req.body;

    const workbook = new ExcelJS.Workbook();

    const filePath = path.join('C:\\data', 'pc_info.xlsx');

    let currentId = 1;  // start with default ID

    // Check if the Excel file exists
    if (fs.existsSync(filePath)) {
        // Load the existing workbook
        await workbook.xlsx.readFile(filePath);

        // Get the main sheet and retrieve the last ID
        const mainSheet = workbook.getWorksheet('Main');
        if (mainSheet) {
            const lastRow = mainSheet.lastRow;
            if (lastRow && lastRow.getCell(1).value) {
                currentId = lastRow.getCell(1).value + 1;
            }
        }
    } else {
        // If the file does not exist, create the sheets and headers
        const mainSheet = workbook.addWorksheet('Main');
        mainSheet.addRow(['ID', 'PC Name', 'Logged In User', 'CPU Name', 'Cores', 'Frequency (GHz)', 'Threads', 'DIMM Slots Used', 'RAM Size (GB)', 'RAM Frequency (MHz)', 'Drive Letter', 'Used Size (GB)', 'Total Size (GB)', 'Disk Type', 'IP Address']);

        const cpuSheet = workbook.addWorksheet('CPU');
        cpuSheet.addRow(['ID', 'CPU Name', 'Cores', 'Frequency (GHz)', 'Threads']);

        const ramSheet = workbook.addWorksheet('RAM');
        ramSheet.addRow(['ID', 'DIMM Slots Used', 'Size (GB)', 'Frequency (MHz)']);

        const diskSheet = workbook.addWorksheet('Disk');
        diskSheet.addRow(['ID', 'Drive Letter', 'Used Size (GB)', 'Total Size (GB)', 'Disk Type']);

        const networkSheet = workbook.addWorksheet('Network');
        networkSheet.addRow(['ID', 'IP Address']);
    }

    // Now add the data using currentId

    // Main Sheet Data (only appending the first disk as an example)
    const mainSheet = workbook.getWorksheet('Main');
    const cDrive = data.disks.find(d => d.driveLetter === 'C:');
    mainSheet.addRow([
        currentId,
        data.pcName || 'N/A',
        data.loggedInUser,
        data.cpuName.trim(),
        data.cpuCores,
        data.cpuFrequencyGHz,
        data.cpuThreads,
        data.dimmSlotsUsed,
        data.ramSizeGB,
        data.ramFrequencyMHz,
        cDrive ? cDrive.driveLetter : 'N/A',
        cDrive ? cDrive.usedSizeGB : 'N/A',
        cDrive ? cDrive.totalSizeGB : 'N/A',
        cDrive ? cDrive.diskType : 'N/A',
        data.ipAddress || 'N/A'
    ]);

    // CPU Sheet Data
    const cpuSheet = workbook.getWorksheet('CPU');
    cpuSheet.addRow([currentId, data.cpuName.trim(), data.cpuCores, data.cpuFrequencyGHz, data.cpuThreads]);

    // RAM Sheet Data
    const ramSheet = workbook.getWorksheet('RAM');
    ramSheet.addRow([currentId, data.dimmSlotsUsed, data.ramSizeGB, data.ramFrequencyMHz]);

    // Disk Sheet Data
    const diskSheet = workbook.getWorksheet('Disk');
    for (let disk of data.disks) {
        diskSheet.addRow([currentId, disk.driveLetter, disk.usedSizeGB, disk.totalSizeGB, disk.diskType]);
    }

    // Network Sheet Data
    const networkSheet = workbook.getWorksheet('Network');
    networkSheet.addRow([currentId, data.ipAddress || 'N/A']);

    // Save the workbook to the file
    await workbook.xlsx.writeFile(filePath);

    res.status(200).send('Data processed and Excel file updated/created successfully in C:\\data');

});

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
