const fs = require('fs');
const path = require('path');
const multer = require('multer');
const { extractUserData } = require('../services/userDataParser');
const { mapToTemplate } = require('../services/templateMapper');

// Configure multer for file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        const uploadPath = path.join(__dirname, '../../uploads');
        cb(null, uploadPath);
    },
    filename: (req, file, cb) => {
        cb(null, `${Date.now()}_${file.originalname}`);
    }
});

const upload = multer({ storage });

/**
 * Read user file and return column headers for preview
 */
async function previewUserFile(req, res) {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }
        
        const userData = extractUserData(req.file.path);
        
        // Return sample data for preview (first 5 rows)
        const preview = userData.slice(0, 5);
        
        res.json({
            success: true,
            recordCount: userData.length,
            preview: preview,
            columns: Object.keys(preview[0] || {})
        });
        
        // Clean up uploaded file
        fs.unlink(req.file.path, () => {});
        
    } catch (error) {
        console.error('Preview error:', error);
        res.status(500).json({ error: error.message });
    }
}

/**
 * Process mapping and export file
 */
async function processMapping(req, res) {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }
        
        const { extractUserData } = require('../services/userDataParser');
        const { mapToTemplate } = require('../services/templateMapper');
        const fs = require('fs');
        const path = require('path');
        
        const filePath = req.file.path;
        const fileName = req.file.originalname;
        const baseName = path.parse(fileName).name;
        
        const userData = extractUserData(filePath);
        
        if (userData.length === 0) {
            return res.status(400).json({ error: 'No valid records found in file' });
        }
        
        // Re-number STT
        userData.forEach((record, index) => {
            record.STT = index + 1;
        });
        
        const templatePath = path.join(__dirname, '../../template/template.xlsx');
        const outputPath = path.join(__dirname, '../../uploads', `output_${Date.now()}.xlsx`);
        
        await mapToTemplate(userData, templatePath, outputPath);
        
        // Clean up
        fs.unlink(filePath, () => {});
        
        // ============ GENERATE DYNAMIC FILENAME ============
        const exportFileName = `Import_${baseName}_${userData.length}CC.xlsx`;
        
        // Send file
        res.download(outputPath, exportFileName, (err) => {
            if (err) console.error('Error sending file:', err);
            setTimeout(() => {
                if (fs.existsSync(outputPath)) fs.unlink(outputPath, () => {});
            }, 5000);
        });
        
    } catch (error) {
        console.error('Process error:', error);
        res.status(500).json({ error: error.message });
    }
}

module.exports = {
    upload,
    previewUserFile,
    processMapping
};