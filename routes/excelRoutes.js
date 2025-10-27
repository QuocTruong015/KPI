const express = require("express");
const multer = require("multer");
const { uploadEmptyPackage, uploadBuyingLabel, uploadScanLabel, uploadEtsyOrder, uploadEtsyStatement, uploadEtsyFFCost, uploadEtsyProfit, uploadAmzTransaction, uploadAmzFFCost, uploadAmzOrder, uploadAmzProfit, uploadWebOrder, uploadWebCost1, uploadWebCost2, uploadWebTotalCost, uploadWebProfit } = require("../controllers/excelController");

const router = express.Router();
const upload = multer({ dest: "uploads/" });

router.post("/upload-excel/empty-package", upload.single("file"), uploadEmptyPackage);
router.post("/upload-excel/buying-label", upload.single("file"), uploadBuyingLabel);
router.post("/upload-excel/scan-label", upload.single("file"), uploadScanLabel);

router.post("/upload-excel/etsy-statement", upload.single("file"), uploadEtsyStatement);
router.post("/upload-excel/etsy-cost", upload.single("file"), uploadEtsyFFCost);
router.post("/upload-excel/etsy-order", upload.single("file"), uploadEtsyOrder);
router.post("/upload-excel/etsy-profit",upload.single("file"), uploadEtsyProfit);

router.post("/upload-excel/amz-transaction", upload.single("file"), uploadAmzTransaction);
router.post("/upload-excel/amz-cost", upload.single("file"), uploadAmzFFCost);
router.post("/upload-excel/amz-order", upload.single("file"), uploadAmzOrder);
router.post("/upload-excel/amz-profit", upload.single("file"), uploadAmzProfit);

router.post("/upload-excel/web-order", upload.single("file"), uploadWebOrder);
router.post('/upload-excel/web-cost1', upload.single('file'), uploadWebCost1);
router.post('/upload-excel/web-cost2', upload.single('file'), uploadWebCost2);
router.post('/upload-excel/web-total-cost', upload.single('file'), uploadWebTotalCost);
router.post('/upload-excel/web-profit', upload.single('file'), uploadWebProfit);

module.exports = router;
