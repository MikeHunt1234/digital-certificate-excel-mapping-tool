const express = require('express');
const router = express.Router();
const { upload, previewUserFile, processMapping } = require('../controllers/mappingController');
const { mergeMultipleFiles } = require('../controllers/excelController');

// Preview user file (get column info)
router.post('/preview', upload.single('userFile'), previewUserFile);

// Process mapping and download
router.post('/process', upload.single('userFile'), processMapping);

// Process multiple excel files
router.post('/merge-multiple', upload.array('userFiles'), mergeMultipleFiles);

module.exports = router;