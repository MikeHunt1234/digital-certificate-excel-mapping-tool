const express = require('express');
const router = express.Router();
const { upload, previewUserFile, processMapping } = require('../controllers/mappingController');

// Preview user file (get column info)
router.post('/preview', upload.single('userFile'), previewUserFile);

// Process mapping and download
router.post('/process', upload.single('userFile'), processMapping);

module.exports = router;