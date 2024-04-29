import * as React from "react";
import { useEffect, useState } from 'react';
import { spfi } from "@pnp/sp";
import { getSP } from "../../../common/pnpjsConfig";
import * as Excel from 'exceljs';
// import { saveAs } from 'file-saver';

const SharePointExcelEditor: React.FC = () => {
    const sp = spfi(getSP());
    const [downloadUrl, setDownloadUrl] = useState<string | null>(null);

    useEffect(() => {
        const fetchAndModifyExcel = async ():Promise<void> => {
            try {
                const buffer = await sp.web.getFileByServerRelativePath("/sites/proj-testspfeatures/Shared Documents/UD BSI_Output Template.xlsx").getBuffer();
                const workbook = new Excel.Workbook();
                await workbook.xlsx.load(buffer); // 加载Excel文件
                const worksheet = workbook.getWorksheet(2); // 获取第一个工作表
                
                // 修改A2单元格并保留样式
                const cell = worksheet.getCell('A2');
                cell.value = "1558"; // 修改单元格的值
                workbook.eachSheet((worksheet, sheetId) => {
                  worksheet.eachRow((row, rowNumber) => {
                    row.eachCell((cell, colNumber) => {
                      if (cell.style.font) {
                        cell.style.font.color = { argb: 'FF000000' }; // 设置字体颜色为黑色
                      } else {
                        cell.style.font = {
                          color: { argb: 'FF000000' }
                        };
                      }
                    });
                  });
                });
                // 将修改后的工作簿写回Blob
                const updatedBuffer = await workbook.xlsx.writeBuffer();
                const blob = new Blob([updatedBuffer], { type: 'application/octet-stream' });
                setDownloadUrl(URL.createObjectURL(blob));
            } catch (error) {
                console.error('Error loading or modifying the file:', error);
            }
        };

        fetchAndModifyExcel().catch(error => {
          console.error('Failed to fetch or modify the Excel file:', error);
      });
    }, []);

    return (
        <div>
            {downloadUrl ? (
                <a href={downloadUrl} download="Modified_UD_BSI_Output_Template.xlsx">下载修改后的Excel文件</a>
            ) : (
                <p>Loading and processing Excel file...</p>
            )}
        </div>
    );
};

export default SharePointExcelEditor;
