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
router.post("/upload-excel/canvas-rev"
, upload.fields([
  { name: "file", maxCount: 1 },
  { name: "costFile", maxCount: 1 },
])
, uploadCanvasRev);
router.post("/upload-excel/service-staff-1", upload.fields([ // thieu empty, buying label, scan label
  { name: "file1", maxCount: 1 }, // UCS_2025
  { name: "file2", maxCount: 1 }, // Daisy
  { name: "file3", maxCount: 1 }, // UCS Poster
]), uploadFulfillmentPosterCost);

router.post("/upload-excel/service-staff-2", upload.fields([
  { name: "file1", maxCount: 1 }, //UCS_2025
  { name: "file2", maxCount: 1 }, //UCS_seller management
]), uploadServiceStaff2);

router.post("/upload-excel/etsy-statement", upload.single("file"), uploadEtsyStatement);  //test
router.post("/upload-excel/etsy-cost", upload.single("file"), uploadEtsyFFCost);          //test
router.post("/upload-excel/etsy-order", upload.single("file"), uploadEtsyOrder);          //test
router.post("/upload-excel/etsy-store", upload.single("file"), uploadEtsyStore);          //test
router.post("/upload-excel/etsy-custom", upload.single("file"), uploadEtsyCustom);        //test
router.post("/upload-excel/etsy-profit",upload.single("file"), uploadEtsyProfit);         //E-commerce

router.post("/upload-excel/amz-transaction", upload.single("file"), uploadAmzTransaction);  //test
router.post("/upload-excel/amz-cost", upload.single("file"), uploadAmzFFCost);              //test
router.post("/upload-excel/amz-order", upload.single("file"), uploadAmzOrder);              //test
router.post("/upload-excel/amz-store", upload.single("file"), uploadAmzStore);              //test
router.post("/upload-excel/amz-custom", upload.single("file"), uploadAmzCustom);            //test
router.post("/upload-excel/amz-profit", upload.single("file"), uploadAmzProfit);            //Ecommerce

router.post("/upload-excel/web-order", upload.single("file"), uploadWebOrder);                //test
router.post('/upload-excel/web-cost1', upload.single('file'), uploadWebCost1);                //test
router.post('/upload-excel/web-cost2', upload.single('file'), uploadWebCost2);                //test
router.post('/upload-excel/web-total-cost', upload.single('file'), uploadWebTotalCost);       //test
router.post('/upload-excel/web-store', upload.single('file'), uploadWebStore);                //test
router.post('/upload-excel/web-profit', upload.single('file'), uploadWebProfit);              //Ecommerce

router.post('/upload-excel/merch-order', upload.single('file'), uploadMerchOrder);                  //test
router.post('/upload-excel/merch-sku', upload.single('file'), uploadMerchSku);                      //test
router.post('/upload-excel/merch-profit', upload.single('file'), uploadMerchProfitByDesignerAndRD); //Ecommerce

router.post('/export-all', upload.single('file'), exportAllProfit);
router.post('/KPI/upload-kpi-target', upload.single('file'), uploadKpiTargetFile);
router.post(
  '/kpi-combined',
  upload.fields([
    { name: 'profit_file', maxCount: 1 }, //E-comerce
    { name: 'target_file', maxCount: 1 } //KPI Target
  ]),
  calculateCombinedKPI
);

module.exports = router;
