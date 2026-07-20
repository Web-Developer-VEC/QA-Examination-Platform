const ExcelJS = require("exceljs");
const { getDb } = require("../../config/db");

async function downloadQuestionBank(req, res) {
  try {
    const { subject_name } = req.body;
    const db = getDb();

    const data = await db.collection("qa_question").findOne({ subject_name });

    if (!data) {
      return res.status(404).json({
        success: false,
        message: "Subject not found",
      });
    }

    function convertValue(value) {
      if (
        typeof value === "string" &&
        /^\d+(\.\d+)?$/.test(value)
      ) {
        return Number(value);
      }

      return value;
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(subject_name);

    // Define columns first
    worksheet.columns = [
      { header: "S.No", key: "sno", width: 8 },
      { header: "Difficulty", key: "difficulty_level", width: 15 },
      { header: "Topic", key: "topic", width: 25 },
      { header: "Question", key: "question", width: 60 },
      { header: "Option A", key: "A", width: 20 },
      { header: "Option B", key: "B", width: 20 },
      { header: "Option C", key: "C", width: 20 },
      { header: "Option D", key: "D", width: 20 },
      { header: "Correct Option", key: "correct_option", width: 20 },
    ];

    // Insert a new first row
    worksheet.insertRow(1, []);

    // Merge the first row
    worksheet.mergeCells("A1:I1");
    worksheet.getCell("A1").value = `${subject_name} QUESTION BANK`;
    worksheet.getCell("A1").font = {
      bold: true,
      size: 16,
    };
    worksheet.getCell("A1").alignment = {
      horizontal: "center",
      vertical: "middle",
    };


    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        cell.alignment = {
          ...cell.alignment,
          vertical: "middle",
        };
      });
    });

    let row = 3;
    let sno = 1;

    for (const exam of data.exam) {
      for (const q of exam.topic_question) {
        worksheet.addRow({
          sno: sno++,
          difficulty_level: q.difficulty_level,
          topic: exam.topic,
          question: q.question,
          A: convertValue(q.A),
          B: convertValue(q.B),
          C: convertValue(q.C),
          D: convertValue(q.D),
          correct_option: convertValue(q.correct_option),
        });
      }
    }
    worksheet.getRow(1).font = {
      bold: true,
      size: 16,
    };

    worksheet.getRow(1).alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    worksheet.getRow(2).font = {
      bold: true,
    };

    worksheet.getRow(2).alignment = {
      horizontal: "center",
      vertical: "middle",
    };

    worksheet.getRow(1).height = 25;

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber <= 2) return;


      row.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
      row.getCell(2).alignment = { horizontal: "center", vertical: "middle" };
      row.getCell(3).alignment = { horizontal: "left", vertical: "middle" };
      row.getCell(4).alignment = {
        horizontal: "left",
        vertical: "middle",
        wrapText: true,
      };

      for (let i = 5; i <= 8; i++) {
        row.getCell(i).alignment = {
          horizontal: "center",
          vertical: "middle",
          wrapText: true,
        };
      }

      row.getCell(9).alignment = {
        horizontal: "center",
        vertical: "middle",
        wrapText: true,
      };
    });

    worksheet.autoFilter = {
      from: "A2",
      to: "I2",
    };

    worksheet.views = [
      {
        state: "frozen",
        ySplit: 2,
      },
    ];

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    res.setHeader(
      "Content-Disposition",
      `attachment; filename=${subject_name}_Question_Bank.xlsx`
    );

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).json({
      success: false,
      message: err.message,
    });
  }
}

module.exports = { downloadQuestionBank };