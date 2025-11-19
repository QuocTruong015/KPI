const express = require("express");
const multer = require("multer");
const { uploadEmptyPackage, uploadBuyingLabel, uploadScanLabel, uploadEtsyOrder, 
    uploadEtsyStatement, uploadEtsyFFCost, uploadEtsyProfit, uploadAmzTransaction, uploadAmzFFCost, 
    uploadAmzOrder, uploadAmzProfit, uploadWebOrder, uploadWebCost1, uploadWebCost2, uploadWebTotalCost, 
    uploadWebProfit, uploadMerchOrder, uploadMerchSku, uploadMerchProfitByDesignerAndRD, exportAllProfit, 
    uploadEtsyStore, uploadAmzStore, uploadWebStore, uploadEtsyCustom, uploadAmzCustom } = require("../controllers/excelController");
const { uploadKpiTargetFile, calculateCombinedKPI } = require("../controllers/kpiController");
const { uploadPhoneCaseCost , uploadPhoneCaseRev, uploadPhoneCaseProfit, uploadTracking, uploadCanvasRev, 
  uploadFulfillmentPosterCost, uploadServiceStaff2} = require("../controllers/excelController");
const router = express.Router();
const upload = multer({ dest: "uploads/" });

router.post("/upload-excel/empty-package", upload.single("file"), uploadEmptyPackage);
router.post("/upload-excel/buying-label", upload.single("file"), uploadBuyingLabel);
router.post("/upload-excel/scan-label", upload.single("file"), uploadScanLabel);
router.post("/upload-excel/phone-case-cost", upload.single("file"), uploadPhoneCaseCost);
router.post("/upload-excel/phone-case-rev", upload.single("file"), uploadPhoneCaseRev);
router.post("/upload-excel/phone-case-profit",
  upload.fields([
    { name: "revFile", maxCount: 1 },
    { name: "costFile", maxCount: 1 },
  ]),
  uploadPhoneCaseProfit
);
router.post("/upload-excel/tracking", upload.single("file"), uploadTracking);
router.post("/upload-excel/canvas-rev", upload.single("file"), uploadCanvasRev);
router.post("/upload-excel/fulfillment-poster-cost", upload.fields([
  { name: "file1", maxCount: 1 },
  { name: "file2", maxCount: 1 },
  { name: "file3", maxCount: 1 },
]), uploadFulfillmentPosterCost);
router.post("/upload-excel/service-staff-2", upload.fields([
  { name: "file1", maxCount: 1 },
  { name: "file2", maxCount: 1 },
]), uploadServiceStaff2);

router.post("/upload-excel/etsy-statement", upload.single("file"), uploadEtsyStatement);
router.post("/upload-excel/etsy-cost", upload.single("file"), uploadEtsyFFCost);
router.post("/upload-excel/etsy-order", upload.single("file"), uploadEtsyOrder);
router.post("/upload-excel/etsy-profit",upload.single("file"), uploadEtsyProfit);
router.post("/upload-excel/etsy-store", upload.single("file"), uploadEtsyStore);
router.post("/upload-excel/etsy-custom", upload.single("file"), uploadEtsyCustom);

router.post("/upload-excel/amz-transaction", upload.single("file"), uploadAmzTransaction);
router.post("/upload-excel/amz-cost", upload.single("file"), uploadAmzFFCost);
router.post("/upload-excel/amz-order", upload.single("file"), uploadAmzOrder);
router.post("/upload-excel/amz-profit", upload.single("file"), uploadAmzProfit);
router.post("/upload-excel/amz-store", upload.single("file"), uploadAmzStore);
router.post("/upload-excel/amz-custom", upload.single("file"), uploadAmzCustom);

router.post("/upload-excel/web-order", upload.single("file"), uploadWebOrder);
router.post('/upload-excel/web-cost1', upload.single('file'), uploadWebCost1);
router.post('/upload-excel/web-cost2', upload.single('file'), uploadWebCost2);
router.post('/upload-excel/web-total-cost', upload.single('file'), uploadWebTotalCost);
router.post('/upload-excel/web-profit', upload.single('file'), uploadWebProfit);
router.post('/upload-excel/web-store', upload.single('file'), uploadWebStore);

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
