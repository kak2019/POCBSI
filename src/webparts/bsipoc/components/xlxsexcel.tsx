import * as React from 'react';
import  { useState, useEffect } from 'react';
import { spfi } from "@pnp/sp";
import { getSP } from "../../../common/pnpjsConfig";
import * as XLSX from 'xlsx';



function s2ab(s: string) :ArrayBuffer {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) {
        view[i] = s.charCodeAt(i) & 0xFF;
    }
    return buf;
}

const SharePointExcelEditor = (): React.ReactElement => {
    const sp = spfi(getSP());
    const [downloadUrl, setDownloadUrl] = useState("");

    useEffect(() => {
        const fetchAndModifyExcel = async ():Promise<void> => {
            try {
                const buffer = await sp.web.getFileByServerRelativePath("/sites/proj-testspfeatures/Shared Documents/UD BSI_Output Template.xlsx").getBuffer();
                const workbook = XLSX.read(buffer, {type: 'buffer'});

                // const firstSheetName = workbook.SheetNames[0];
                //const worksheet = workbook.Sheets[firstSheetName];
                //console.log("fir",firstSheetName,"worksheet",worksheet)
                
                const worksheet = workbook.Sheets[workbook.SheetNames[1]];
                console.log("worksheet",worksheet)
                if (worksheet.B3) {
                    console.log("Before modification:", worksheet.B3.v); // 打印修改前的值
                    worksheet.B3.v = '新的值';  // 尝试修改
                    console.log("After modification:", worksheet.B3.v); // 打印修改后的值
                } else {
                    console.log("B3 cell does not exist, creating one.");
                    worksheet.B3 = { t: 's', v: '新的值' };
                }
                const wbout = XLSX.write(workbook, {bookType: 'xlsx', type: 'binary'});
                const blob = new Blob([s2ab(wbout)], {type: 'application/octet-stream'});
                // 创建下载链接
                const url = URL.createObjectURL(blob);
                setDownloadUrl(url);
            } catch (error) {
                console.error('Error loading or modifying the file:', error);
            }
        };

        fetchAndModifyExcel().catch(error => {
            console.error('Failed to fetch or modify the Excel file:', error)});
    }, []);

    return (
        <div>
            {downloadUrl ? (
                <a href={downloadUrl} download="Modified_Excel1.xlsx">下载修改后的Excel文件</a>
            ) : (
                <p>Loading and processing Excel file...</p>
            )}
        </div>
    );
};

export default SharePointExcelEditor;
