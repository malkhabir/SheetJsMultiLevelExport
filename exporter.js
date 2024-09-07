const data = [
    {
        question: "Question 1",
        style: "Style A",
        thead: "Thead A",
        season: "Season 1",
        group: 1,
        detail: [
            {
                answer: "Answer 1",
                percent: "50%",
                response: "Response 1",
                subdetail: [
                    { subanswer: "Sub Answer 1", subpercent: "10%", subresponse: "Sub Response 1" },
                    { subanswer: "Sub Answer 2", subpercent: "15%", subresponse: "Sub Response 2" },
                ],
            },
            {
                answer: "Answer 2",
                percent: "30%",
                response: "Response 2",
                subdetail: [
                    { subanswer: "Sub Answer 3", subpercent: "20%", subresponse: "Sub Response 3" },
                    { subanswer: "Sub Answer 4", subpercent: "25%", subresponse: "Sub Response 4" },
                ],
            },
        ],
    },
    {
        question: "Question 2",
        style: "Style B",
        thead: "Thead B",
        season: "Season 2",
        group: 2,
        detail: [
            {
                answer: "Answer 3",
                percent: "20%",
                response: "Response 3",
                subdetail: [
                    { subanswer: "Sub Answer 5", subpercent: "5%", subresponse: "Sub Response 5" },
                    { subanswer: "Sub Answer 6", subpercent: "10%", subresponse: "Sub Response 6" },
                ],
            },
            {
                answer: "Answer 4",
                percent: "40%",
                response: "Response 4",
                subdetail: [
                    { subanswer: "Sub Answer 7", subpercent: "15%", subresponse: "Sub Response 7" },
                    { subanswer: "Sub Answer 8", subpercent: "10%", subresponse: "Sub Response 8" },
                ],
            },
        ],
    },
    {
        question: "Question 3",
        style: "Style C",
        thead: "Thead C",
        season: "Season 3",
        group: 3,
        detail: [
            {
                answer: "Answer 5",
                percent: "25%",
                response: "Response 5",
                subdetail: [
                    { subanswer: "Sub Answer 9", subpercent: "12%", subresponse: "Sub Response 9" },
                    { subanswer: "Sub Answer 10", subpercent: "8%", subresponse: "Sub Response 10" },
                ],
            },
            {
                answer: "Answer 6",
                percent: "35%",
                response: "Response 6",
                subdetail: [
                    { subanswer: "Sub Answer 11", subpercent: "18%", subresponse: "Sub Response 11" },
                    { subanswer: "Sub Answer 12", subpercent: "10%", subresponse: "Sub Response 12" },
                ],
            },
        ],
    },
    {
        question: "Question 4",
        style: "Style D",
        thead: "Thead D",
        season: "Season 4",
        group: 4,
        detail: [
            {
                answer: "Answer 7",
                percent: "40%",
                response: "Response 7",
                subdetail: [
                    { subanswer: "Sub Answer 13", subpercent: "20%", subresponse: "Sub Response 13" },
                    { subanswer: "Sub Answer 14", subpercent: "15%", subresponse: "Sub Response 14" },
                ],
            },
            {
                answer: "Answer 8",
                percent: "30%",
                response: "Response 8",
                subdetail: [
                    { subanswer: "Sub Answer 15", subpercent: "12%", subresponse: "Sub Response 15" },
                    { subanswer: "Sub Answer 16", subpercent: "18%", subresponse: "Sub Response 16" },
                ],
            },
        ],
    },
    {
        question: "Question 5",
        style: "Style E",
        thead: "Thead E",
        season: "Season 5",
        group: 5,
        detail: [
            {
                answer: "Answer 9",
                percent: "55%",
                response: "Response 9",
                subdetail: [
                    { subanswer: "Sub Answer 17", subpercent: "25%", subresponse: "Sub Response 17" },
                    { subanswer: "Sub Answer 18", subpercent: "10%", subresponse: "Sub Response 18" },
                ],
            },
            {
                answer: "Answer 10",
                percent: "45%",
                response: "Response 10",
                subdetail: [
                    { subanswer: "Sub Answer 19", subpercent: "15%", subresponse: "Sub Response 19" },
                    { subanswer: "Sub Answer 20", subpercent: "10%", subresponse: "Sub Response 20" },
                ],
            },
        ],
    },
    {
        question: "Question 6",
        style: "Style F",
        thead: "Thead F",
        season: "Season 6",
        group: 6,
        detail: [
            {
                answer: "Answer 11",
                percent: "60%",
                response: "Response 11",
                subdetail: [
                    { subanswer: "Sub Answer 21", subpercent: "30%", subresponse: "Sub Response 21" },
                    { subanswer: "Sub Answer 22", subpercent: "10%", subresponse: "Sub Response 22" },
                ],
            },
            {
                answer: "Answer 12",
                percent: "25%",
                response: "Response 12",
                subdetail: [
                    { subanswer: "Sub Answer 23", subpercent: "10%", subresponse: "Sub Response 23" },
                    { subanswer: "Sub Answer 24", subpercent: "15%", subresponse: "Sub Response 24" },
                ],
            },
        ],
    },
];


document.addEventListener('DOMContentLoaded', () => {
    // Function to export data
    function exportData() {
        // Convert data to a worksheet with proper grouping
        const wsData = [];
        data.forEach((item) => {
            // Add question row (Level 1)
            wsData.push([item.question, item.style, item.thead, item.season, ""]);

            // Add details (Level 2) and subdetails (Level 3)
            item.detail.forEach((detail) => {
                wsData.push(["", detail.answer, detail.percent, detail.response, ""]);

                // Add subdetails rows (Level 3) if they exist
                if (detail.subdetail && detail.subdetail.length > 0) {
                    detail.subdetail.forEach((sub) => {
                        wsData.push(["", "", sub.subanswer, sub.subpercent, sub.subresponse]);
                    });
                }


            });

            // Add an empty row to separate different groups
            wsData.push([]);
        });

        // Create a new workbook and add the worksheet
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(wsData);

        // Function to apply grouping
        function groupRows(ws, start_row, end_row, level = 1) {
            if (!ws["!rows"]) ws["!rows"] = [];
            for (let i = start_row; i <= end_row; ++i) {
                if (!ws["!rows"][i]) ws["!rows"][i] = { hpx: 20 };
                ws["!rows"][i].level = level ;
            }
        }

        // Grouping rows for three levels
        let startRow = 0;
        data.forEach((item) => {
            const questionRow = startRow;
            var detailRow = startRow + 1;

            // Group details under the question (Level 1) - all answers grouped under the question
            const endDetailRow =
                detailRow +
                item.detail.length +
                item.detail.reduce((acc, d) => acc + (d.subdetail ? d.subdetail.length : 0), 0);
            groupRows(ws, questionRow + 1, endDetailRow, 1); // Group all answers under the question

            // Group subdetails under each detail (Level 2)
            item.detail.forEach((detail) => {
                const detailStartRow = detailRow;
                const detailHasSubdetails = detail.subdetail && detail.subdetail.length > 0;

                if (detailHasSubdetails) {
                    const subdetailStartRow = detailStartRow + 1;
                    const subdetailEndRow = subdetailStartRow + detail.subdetail.length - 1;

                    // Group subdetail rows (Level 2) under the detail row
                    groupRows(ws, subdetailStartRow, subdetailEndRow, 2);

                    detailRow = subdetailEndRow + 1; // Move past the subdetail rows
                } else {
                    detailRow += 1; // If no subdetails, move to the next detail
                }
            });

            // Move to the next question group
            startRow = detailRow + 1; // Add a row for separation
        });

        // Ensure the outline settings are correct
        if (!ws["outline"]) ws["outline"] = {};
        ws["outline"].above = false; // Grouping buttons below the rows

        // Add the worksheet to the workbook
        XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

        // Write the file to disk
        XLSX.writeFile(wb, "nested_data_with_question_answer_grouping.xlsx");

        console.log("Excel file created with correct grouping.");
    }

    exportData();
});
