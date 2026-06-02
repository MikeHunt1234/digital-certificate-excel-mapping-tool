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
        
        // Extract user data
        const userData = extractUserData(req.file.path);
        
        if (userData.length === 0) {
            return res.status(400).json({ error: 'No data found in uploaded file' });
        }
        
        // Template path
        const templatePath = path.join(__dirname, '../../template/template.xlsx');
        
        if (!fs.existsSync(templatePath)) {
            return res.status(500).json({ error: 'Template file not found. Please ensure template.xlsx is in the /template folder.' });
        }
        
        // Output path
        const outputPath = path.join(__dirname, '../../uploads', `output_${Date.now()}.xlsx`);
        
        // Map to template
        const recordCount = await mapToTemplate(userData, templatePath, outputPath);
        
        // Send file
        res.download(outputPath, 'mapped_data.xlsx', (err) => {
            if (err) console.error('Download error:', err);
            // Clean up output file after download
            setTimeout(() => {
                if (fs.existsSync(outputPath)) fs.unlink(outputPath, () => {});
            }, 5000);
        });
        
        // Clean up uploaded file
        fs.unlink(req.file.path, () => {});
        
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