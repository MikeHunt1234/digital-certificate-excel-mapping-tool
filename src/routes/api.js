const express = require('express');
const router = express.Router();
const multer = require('multer');
const path = require('path');
const fs = require('fs');

// Import controllers
const { upload, previewUserFile, processMapping } = require('../controllers/mappingController');
const { mergeMultipleFiles } = require('../controllers/excelController');
const { previewLegacyFile, processLegacyFiles } = require('../controllers/legacyController');

// Ensure uploads directory exists
const uploadsDir = path.join(__dirname, '../../uploads');
if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir, { recursive: true });
}

// Configure multer for file uploads
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, uploadsDir);
    },
    filename: function (req, file, cb) {
        cb(null, `${Date.now()}_${file.originalname}`);
    }
});

const uploadMiddleware = multer({ 
    storage: storage,
    limits: { fileSize: 50 * 1024 * 1024 } // 50MB limit
});

// ============ MODERN MODE ROUTES ============

// Preview modern file (get column info)
router.post('/preview', uploadMiddleware.single('userFile'), previewUserFile);

// Process modern mapping and download
router.post('/process', uploadMiddleware.single('userFile'), processMapping);

// Process multiple modern files
router.post('/merge-multiple', uploadMiddleware.array('userFiles'), mergeMultipleFiles);

// ============ LEGACY MODE ROUTES ============

// Preview legacy file (single file preview)
router.post('/legacy/preview', uploadMiddleware.single('userFile'), previewLegacyFile);

// Process multiple legacy files with overrides
router.post('/legacy/process', uploadMiddleware.array('userFiles'), processLegacyFiles);

module.exports = router;