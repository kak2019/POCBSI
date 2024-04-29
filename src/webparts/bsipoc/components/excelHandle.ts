import * as Excel from 'exceljs';
import { spfi } from "@pnp/sp";
import { getSP } from "../../../common/pnpjsConfig";

// 导出一个函数用于修改Excel文件
export const modifyExcelFile= async (setDownloadUrl: React.Dispatch<React.SetStateAction<string | null>>,values: string[]): Promise<void> => {
    const sp = spfi(getSP());
    try {
        const buffer = await sp.web.getFileByServerRelativePath("/sites/proj-testspfeatures/Shared Documents/UD BSI_Output Template.xlsx").getBuffer();
        const workbook = new Excel.Workbook();
        await workbook.xlsx.load(buffer); // 加载Excel文件
        const worksheet = workbook.getWorksheet(2); // 获取第一个工作表
                
        // 修改A2单元格并保留样式
        /* 
            B2 => Country
            
        
        */ 
        const cell = worksheet.getCell('B2');
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
