// src/controllers/enrichController.js
const fs = require('fs');
const path = require('path');
const { enrichDataFromFiles } = require('../services/dataEnricher');

/**
 * Process enrichment with multiple source files
 */
async function processEnrichment(req, res) {
    try {
        if (!req.files || !req.files.targetFile || !req.files.sourceFiles) {
            return res.status(400).json({ 
                error: 'Both target file and at least one source file are required' 
            });
        }
        
        const targetFile = req.files.targetFile[0];
        const sourceFiles = req.files.sourceFiles; // Array of files
        
        console.log(`📄 Target file: ${targetFile.originalname}`);
        console.log(`📄 Source files: ${sourceFiles.length} file(s)`);
        sourceFiles.forEach((f, i) => {
            console.log(`   ${i + 1}. ${f.originalname} (${(f.size / 1024).toFixed(1)} KB)`);
        });
        
        // Get user mapping from request body (optional)
        let userMapping = null;
        if (req.body.mapping) {
            try {
                userMapping = JSON.parse(req.body.mapping);
            } catch (e) {
                console.warn('Invalid mapping JSON, using auto-detection');
            }
        }
        
        // Output path
        const outputPath = path.join(__dirname, '../../uploads', `enriched_${Date.now()}.xlsx`);
        
        // Get array of source file paths
        const sourcePaths = sourceFiles.map(f => f.path);
        
        // Process enrichment with multiple source files
        const result = await enrichDataFromFiles(
            targetFile.path,
            sourcePaths, // Pass array of source file paths
            outputPath,
            {
                userMapping: userMapping,
                overwriteExisting: true
            }
        );
        
        // Clean up uploaded files
        fs.unlink(targetFile.path, () => {});
        sourceFiles.forEach(f => fs.unlink(f.path, () => {}));
        
        // Use the original target filename
        const baseName = path.parse(targetFile.originalname).name;
        const exportFileName = `${baseName}.xlsx`;
        
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