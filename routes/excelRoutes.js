const express = require("express");
const multer = require("multer");
const { uploadEmptyPackage, uploadBuyingLabel, uploadScanLabel, uploadEtsyOrder, uploadEtsyStatement, uploadEtsyFFCost } = require("../controllers/excelController");

const router = express.Router();
const upload = multer({ dest: "uploads/" });

router.post("/upload-excel/empty-package", upload.single("file"), uploadEmptyPackage);
router.post("/upload-excel/buying-label", upload.single("file"), uploadBuyingLabel);
router.post("/upload-excel/scan-label", upload.single("file"), uploadScanLabel);

router.post("/upload-excel/etsy-statement", upload.single("file"), uploadEtsyStatement);
router.post("/upload-excel/etsy-cost", upload.single("file"), uploadEtsyFFCost);
router.post("/upload-excel/etsy-order", upload.single("file"), uploadEtsyOrder);

module.exports = router;
