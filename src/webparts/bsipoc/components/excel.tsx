import * as React from "react";
import * as XLSX from 'xlsx';

function ExportExcel() {
    const data = [
        { name: "John", city: "New York", email: "john@example.com" },
        { name: "Jane", city: "Paris", email: "jane@example.com" },
        { name: "Dan", city: "London", email: "dan@example.com" }
    ];

    const handleExport = () => {
        const worksheet = XLSX.utils.json_to_sheet(data);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "People");
        XLSX.writeFile(workbook, "peopleData.xlsx");
    };

    return (
        <button onClick={handleExport}>Export to Excel</button>
    );
}

export default ExportExcel;
