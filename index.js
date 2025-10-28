const express = require("express");
const excelRoutes = require("./routes/excelRoutes");
const { aggregateProfit } = require("./services/aggregateService")

const app = express();
const PORT = 3000;

app.use('/', excelRoutes); // ← BẮT BUỘC!

app.get("/", (req, res) => {
  res.send(`
    <h2>Upload file Excel để đọc và tính Profit</h2>
    <p>Truyền tháng và năm: <code>?month=5&year=2025</code></p>
    <form action="/upload-excel?month=5&year=2025" method="post" enctype="multipart/form-data">
      <input type="file" name="file" />
      <button type="submit">Tải lên</button>
    </form>
  `);
});

app.listen(PORT, () => console.log(`🚀 Server chạy tại http://localhost:${PORT}`));
