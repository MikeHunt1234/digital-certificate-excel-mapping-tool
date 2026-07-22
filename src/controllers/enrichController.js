const fs = require('fs');
const path = require('path');
const { enrichDataFromFiles } = require('../services/dataEnricher');

/**
 * Process enrichment of two Excel files
 */
async function processEnrichment(req, res) {
    try {
        if (!req.files || !req.files.targetFile || !req.files.sourceFile) {
            return res.status(400).json({ 
                error: 'Both target file and source file are required' 
            });
        }
        
        const targetFile = req.files.targetFile[0];
        const sourceFile = req.files.sourceFile[0];
        
        console.log(`📄 Target file: ${targetFile.originalname}`);
        console.log(`📄 Source file: ${sourceFile.originalname}`);
        
        // Output path
        const outputPath = path.join(__dirname, '../../uploads', `${Date.now()}.xlsx`);
        
        // Process enrichment
        const result = await enrichDataFromFiles(
            targetFile.path,
            sourceFile.path,
            outputPath
        );
        
        // Clean up uploaded files
        fs.unlink(targetFile.path, () => {});
        fs.unlink(sourceFile.path, () => {});
        
        // Get filename from target file
        const baseName = path.parse(targetFile.originalname).name;
        const exportFileName = `enriched_${baseName}_${result.totalRecords}CC.xlsx`;
        
        // Send file
        res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(exportFileName)}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.sendFile(outputPath, (err) => {
            if (err) console.error('Error sending file:', err);
            setTimeout(() => {
                if (fs.existsSync(outputPath)) fs.unlink(outputPath, () => {});
            }, 5000);
        });
        
    } catch (error) {
        console.error('Enrichment error:', error);
        
        // Clean up uploaded files if they exist
        if (req.files) {
            Object.values(req.files).forEach(fileArray => {
                if (Array.isArray(fileArray)) {
                    fileArray.forEach(f => {
                        if (f.path && fs.existsSync(f.path)) fs.unlink(f.path, () => {});
                    });
                }
            });
        }
        
        res.status(500).json({ error: error.message });
    }
}

module.exports = {
    processEnrichment
};