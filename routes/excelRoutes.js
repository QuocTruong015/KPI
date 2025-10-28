const express = require("express");
const multer = require("multer");
const { uploadEmptyPackage, uploadBuyingLabel, uploadScanLabel, uploadEtsyOrder, 
    uploadEtsyStatement, uploadEtsyFFCost, uploadEtsyProfit, uploadAmzTransaction, uploadAmzFFCost, 
    uploadAmzOrder, uploadAmzProfit, uploadWebOrder, uploadWebCost1, uploadWebCost2, uploadWebTotalCost, 
    uploadWebProfit, uploadMerchOrder, uploadMerchSku, uploadMerchProfitByDesignerAndRD, exportAllProfit } = require("../controllers/excelController");
const { uploadKpiTargetFile } = require("../controllers/kpiController");
const { calculateCombinedKPI } = require("../controllers/kpiController");
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

router.post('/upload-excel/merch-order', upload.single('file'), uploadMerchOrder);
router.post('/upload-excel/merch-sku', upload.single('file'), uploadMerchSku);
router.post('/upload-excel/merch-profit', upload.single('file'), uploadMerchProfitByDesignerAndRD);

router.post('/export-all', upload.single('file'), exportAllProfit);
router.post('/KPI/upload-kpi-target', upload.single('file'), uploadKpiTargetFile);
router.post(
  '/kpi-combined',
  upload.fields([
    { name: 'profit_file', maxCount: 1 },
    { name: 'target_file', maxCount: 1 }
  ]),
  calculateCombinedKPI
);

module.exports = router;
