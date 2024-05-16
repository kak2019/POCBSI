/* eslint-disable */
import * as React from "react";
import { memo, useEffect } from "react";
import { spfi } from "@pnp/sp";
import { getSP } from "../../../common/pnpjsConfig";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import { Label, PrimaryButton, Stack } from "office-ui-fabric-react";
import * as XLSX from 'xlsx';
import { Dropdown, IDropdownStyles, IDropdownOption } from '@fluentui/react/lib/Dropdown';

import YearPicker from "./control/YearSelect";
import createrFolder from "./control/CreateFolder"
import * as Excel from 'exceljs';
import AppContext from "../../../common/AppContext";
// import XlxsExcelFromSP from "./xlxsexcel"


const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 150 ,marginRight:10},
  
};

const options: IDropdownOption[] = [

  { key: 'Q1', text: 'Q1' },
  { key: 'Q2', text: 'Q2' },
  { key: 'Q3', text: 'Q3' },
  { key: 'Q4', text: 'Q4' },
];


export default memo(function App() {
  const sp = spfi(getSP());
  const ctx = React.useContext(AppContext);
  const webURL = ctx.context?._pageContext?._web?.absoluteUrl;
  const userEmail = ctx.context?._pageContext?._user?.email;
  const Site_Relative_Links = webURL.slice(webURL.indexOf('/sites'))
  console.log("site",Site_Relative_Links)
  console.log("weburl",webURL,userEmail)
  const [excel, setExcel] = React.useState([]);
  const [allCountry, setAllCountry] = React.useState([])
  const [priceTable, setPrice] = React.useState({})
  const [selectedKey, setSelectedKey] = React.useState<string>('');
  const [selectedYear, setSelectedYear] = React.useState<number | undefined>(undefined);

  const handleYearChange = (year: number) => {
      setSelectedYear(year);  // 更新状态以保存选中的年份
      console.log(`Selected year in App component: ${year}`);
  };

  const onChange = (event: React.FormEvent<HTMLDivElement>, item?: IDropdownOption): void => {
    if (item) {
      setSelectedKey(item.key as string);
      console.log('Selected:', item.key);
    }
  };

  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  function calcToExcel(orders: any, priceMap: any) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return orders.map((val: any) => {
      const obj: any = {
        Market: val.Market,
        Dealer: val.PartnerName,
        "Dealer ID": val.PartnerID,
        "Dealer Type": val.DealerType,
        "Dealer Category": val.DealerCategory,
        "Basic Package": val.BasicPackage,
        "Sales Package": val.SalesPackage,
        "Hub Package":val.HubPackage,
        "CPQ": val.CPQ,
        "UD CM": val.UDCM,
        "Argus 365": val.Argus365,
        "UDCP": val.UDCP,
        "SeMA": val.SeMA,
        "LDS": val.LDS,
        "LSS": val.LSS,
        "Pardot": val.Pardot,
        
      }
      // console.log("obj",obj)
      const keys = ['Hub Package','CPQ', 'UD CM', 'Argus 365', 'UDCP', 'SeMA', 'LDS', 'LSS', 'Pardot']
      obj['Total(Per Month)'] = keys.reduce((sum, key) => sum + (priceMap[key] || 0) * obj[key], 0)
      obj['Total(Per Month)'] += (priceMap['Basic Package;' + val.DealerCategory] || 0) * obj['Basic Package']
      obj['Total(Per Month)'] += (priceMap['Sales Package;' + (val.DealerCategory || 'NA')] || 0) * obj['Sales Package']
      return obj
    })
  }

  function calcToSummary(details: any, priceMap: any) {
    const resObj: any = {
      Country: details[0].Market,
      data: []
    }
    console.log("res",resObj)
    const p = {...priceMap}
    for(let key in p) {
      const value = p[key]
      p[key] = {
        price: value,
        count: 0,
        total: 0
      }
    }
    details.map((val:any) => ({...val})).forEach((val: any) => {
      val['Basic Package;' + val['Dealer Category']] = val['Basic Package']
      val['Sales Package;' + (val['Dealer Category'] || 'NA')] = val['Basic Package']
      for(let key in p) {
        p[key].count += Number(val[key] || 0)
      }
    })

    for(let key in p) {
      if(p[key].count === 0) continue;
      resObj.data.push({
        A: key.split(';')[0],
        B: key.split(';').length > 1 ? '- ' + key.split(';')[1] : '',
        C: p[key].count,
        D: p[key].price,
        E: p[key].count * p[key].price
      })
    }
    resObj.data = resObj.data.filter((item: any) => item.C > 0)
    const total = resObj.data.reduce((t: number, e: any) => t + Number(e.E), 0)
    resObj.data.push({
      A: 'Total',
      E: total
    })
    return resObj
  }

  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  const initData = async () => {
    // 拿应用价格表
    const appObj = await sp.web.lists.getByTitle("ApplicationPriceMaster").renderListDataAsStream({
      /* 字段关系如下
      Title ===Application Name
      field_1 = Price Type
      field_2 = Price (USD)
      
      */

      ViewXml: `<View>
                      <ViewFields>
                        <FieldRef Name="ApplicationName"/>
                        <FieldRef Name="PriceType"/>
                        <FieldRef Name="Price_x0028_USD_x0029_"/>
                      </ViewFields>
                   
                    </View>`,
      // <RowLimit>400</RowLimit>
    }).then((response) => {
      console.log("应用单价/人", response.Row)
      // console.log("resAPP", response.Row.filter((item)=>item.field_2))
      if (response.Row.length > 0) {
        const resObj: any = {}
        response.Row.forEach(val => {
          resObj[val.ApplicationName] = val.Price_x0028_USD_x0029_ * 1
        })
        return resObj
      }

      return {}
    })


    // 拿包的单价表
    const packageObj = await sp.web.lists.getByTitle("PackageMaster").renderListDataAsStream({
      /* 字段关系如下
      Title ===Package Name
      field_1 = PartnerType
      field_2 = Dealer Category
      field_3 = Monthly Price (USD)
      Comment = Comment
      */

      ViewXml: `<View>
                      <ViewFields>
                        <FieldRef Name="PackageCategory"/>
                        <FieldRef Name="PartnerType"/>
                        <FieldRef Name="DealerCategory"/>
                        <FieldRef Name="MonthlyPrice_x0028_USD_x0029_"/>
                        <FieldRef Name="Comment"/>
                      </ViewFields>
                   
                    </View>`,
      // <RowLimit>400</RowLimit>
    }).then((response) => {
      console.log("包单价", response.Row)
      // console.log("respackage", response.Row.filter((item)=>item.field_2))
      if (response.Row.length > 0) {
        const resObj: any = {}
        response.Row.forEach(val => {
          resObj[`${val.PackageCategory};${val.DealerCategory}`] = val.MonthlyPrice_x0028_USD_x0029_ * 1
        })
        console.log("resobji",resObj)
        return resObj
      }

      return {}
    })

    // 拿period 周期对应关系
    const period = await sp.web.lists.getByTitle("Cost Calculation Period").renderListDataAsStream({

      ViewXml: `<View>
                      <ViewFields>
                        <FieldRef Name="PeriodName"/>
                        <FieldRef Name="PeriodDetails"/>
                        <FieldRef Name="NumberofMonths"/>
                        <FieldRef Name="AvailableforSelection"/>
                        <FieldRef Name="Generated"/>
                      </ViewFields>
                   
                    </View>`,
      // <RowLimit>400</RowLimit>
    }).then((response) => {
      console.log("period", response.Row)
      // console.log("respackage", response.Row.filter((item)=>item.field_2))
      // if (response.Row.length > 0) {
      //   const resObj: any = {}
      //   response.Row.forEach(val => {
      //     resObj[`${val.Title};${val.field_2}`] = val.field_3 * 1
      //   })
      //   return resObj
      // }

      return {}
    })

    // 拿主表订单
    const order = await sp.web.lists.getByTitle("Partner Master").renderListDataAsStream({
      /* 字段关系如下
      Title ===Hub
      field_1 = PartnerType
      field_2 = country
      field_3 = Partner Name
      field_4 = Partner ID
      field_5 = Exclusive or Multi brand
      field_6 = Dealer Type
      field_7 = No of Bays
      field_8 = Dealer Category
      field_9 = Basic Package
      field_10 = Sales Package
      field_11 = CPQ
      field_12 = UD CM
      field_13 = Argus 365
      field_14 = UDCP
      field_15 = SeMA
      field_16 = LDS
      field_17 = LSS
      field_18 = Pardot
      field_19 = Hub Package
      */

      ViewXml: `<View>
                      <ViewFields>
                        <FieldRef Name="Hub"/>
                        <FieldRef Name="PartnerType"/>
                        <FieldRef Name="Market"/>
                        <FieldRef Name="PartnerName"/>
                        <FieldRef Name="PartnerID"/>
                        <FieldRef Name="ExclusiveorMultibrand"/>
                        <FieldRef Name="DealerType"/>
                        <FieldRef Name="NoofBays"/>
                        <FieldRef Name="DealerCategory"/>
                        <FieldRef Name="BasicPackage"/>
                        <FieldRef Name="SalesPackage"/>
                        <FieldRef Name="HubPackage"/>
                        <FieldRef Name="CPQ"/>
                        <FieldRef Name="UDCM"/>
                        <FieldRef Name="Argus365"/>
                        <FieldRef Name="UDCP"/>
                        <FieldRef Name="SeMA"/>
                        <FieldRef Name="LDS"/>
                        <FieldRef Name="LSS"/>
                        <FieldRef Name="Pardot"/>
                        
                        
                      </ViewFields>
                   
                    </View>`,
      // <RowLimit>400</RowLimit>
    }).then((response) => {
      console.log("主表订单", response.Row)
      // console.log("res", response.Row.filter((item)=>item.field_2==="NZ"))
      if (response.Row.length > 0) {
        const uniqueList = Array.from(new Set(response.Row.map(item => item.Market)))
        console.log("country", uniqueList)
        setAllCountry(uniqueList)
        //setAllCountry
        return response.Row
      }
      return []
    })

    const price = {
      ...appObj,
      ...packageObj
    }
    setPrice(price)
    const finalExcelData = calcToExcel(order, price)
    console.log(calcToSummary(finalExcelData, price))
    console.log("excel121",finalExcelData)
    setExcel(finalExcelData)
  }

  const changeStyle = async (buffer: Excel.Buffer) => {
    const workbook = new Excel.Workbook();
    await workbook.xlsx.load(buffer); // 加载Excel文件
    const worksheet = workbook.getWorksheet(1); // 获取第一个工作表

    ['A4', 'B4', 'C4', 'D4', 'E4'].forEach(cell => {
      worksheet.getCell(cell).fill= {
        type: 'pattern',
        pattern:'solid',
        bgColor:{argb:'FFc0d6ed'}
      }
    })

    const worksheetDetail = workbook.getWorksheet(2)
    const zimu = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'N', 'M', 'O', 'P',"Q"]
    zimu.forEach(z => {
      worksheetDetail.getCell(z + '2').fill= {
        type: 'pattern',
        pattern:'solid',
        bgColor:{argb:'FFb2d7e5'}
      }
      worksheetDetail.getCell(z + '3').fill= {
        type: 'pattern',
        pattern:'solid',
        bgColor:{argb:'FFb2d7e5'}
      }
    })
    
    // 将修改后的工作簿写回Blob
    const updatedBuffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([updatedBuffer], { type: 'application/octet-stream' });
    return blob
  }

  // 定义上传文件的函数
  async function uploadFileToSP(libraryUrl: string, fileName: string, blob: Blob): Promise<void> {
    try {
      const folder = sp.web.getFolderByServerRelativePath(libraryUrl);

      const file = await folder.files.addUsingPath(fileName, blob, { Overwrite: true });
      console.log(`File uploaded successfully! File URL: ${file.data.ServerRelativeUrl}`);
    } catch (error) {
      console.error('Error uploading file:', error);
    }
  }


  const handleExport = async (): Promise<void> => {
    // 创建一个数组来存储所有的上传Promise
    const uploadPromises: any[] = [];

    const buffer = await sp.web.getFileByServerRelativePath(Site_Relative_Links+"/BSITemplate/UD BSI_Output Template.xlsx").getBuffer();

    // 遍历所有国家
    allCountry.forEach(Market => {
      // 筛选出该国家的订单
      // console.log("2exe",excel)
      const countryOrders = excel.filter(order => order.Market === Market);
      if (countryOrders.length === 0) {
        console.log(`No data for ${Market}`);
        return;
      }

      /* summary */
      const workbookTemplate = XLSX.read(buffer, { type: 'buffer' });
      const summaryTemplateName = workbookTemplate.SheetNames[1]
      const workSheetSummaryTpt = workbookTemplate.Sheets[summaryTemplateName]
      // const arrTpt = XLSX.utils.sheet_to_json(workSheetSummaryTpt)
      
      const tongji = calcToSummary(countryOrders, priceTable)
      // \console.log("tongji",tongji)
      workSheetSummaryTpt['B2'] = { v: tongji.Market }
      workSheetSummaryTpt['E2'] = { v: `${selectedYear}/${(Number(selectedKey.replace('Q', '')) -1 )*3+1} - ${selectedYear}/${(Number(selectedKey.replace('Q', '')))*3}` }
      for(let i = 5; i<tongji.data.length+5; i++) {
        workSheetSummaryTpt['A'+ i] = { v: tongji.data[i-5].A}
        workSheetSummaryTpt['B'+ i] = { v: tongji.data[i-5].B}
        workSheetSummaryTpt['C'+ i] = { v: tongji.data[i-5].C}
        workSheetSummaryTpt['D'+ i] ={ v:  tongji.data[i-5].D}
        workSheetSummaryTpt['E'+ i] = { v: tongji.data[i-5].E}
      }
      workSheetSummaryTpt['!ref'] = 'A1:G20'
      /* summary */

      // 创建工作表
      // const worksheet = XLSX.utils.json_to_sheet(countryOrders);
      // ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1', 'J1', 'K1', 'L1', 'N1', 'M1', 'O1', 'P1'].forEach(key => {
      //   if (worksheet[key]) {
      //     worksheet[key].s = {
      //       fill: {
      //         fgColor: { rgb: "add8e6" }
      //       }
      //     };
      //   }
      // });
      const summaryTemplateName2 = workbookTemplate.SheetNames[3]
      const workSheetDetails = workbookTemplate.Sheets[summaryTemplateName2]
      const zimu = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P','Q']
      for(let i = 3; i<countryOrders.length+3; i++) {
        let j = 0;
        for(let key in countryOrders[i-3]) {
          workSheetDetails[zimu[j] +(i+1)] = {v: countryOrders[i-3][key]} 
          j++
        }
      }
      workSheetDetails['!ref'] = 'A1:R100'
      
      // 创建工作簿
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, workSheetSummaryTpt, "Market Summary");
      // XLSX.utils.book_append_sheet(workbook, worksheet, "Package Details");
      XLSX.utils.book_append_sheet(workbook, workSheetDetails, "Package Details");

      // 将工作簿转换为Blob
      const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      // const blob = new Blob([wbout], { type: "application/octet-stream" });
      
      changeStyle(wbout).then(blob => {
        // 添加上传任务到数组
        const uploadPromise = uploadFileToSP(
          `${Site_Relative_Links}/Shared Documents/${selectedYear}${selectedKey}`,
          `UD ${Market} ${selectedYear}${selectedKey}.xlsx`,
          blob
        );
        uploadPromises.push(uploadPromise);
      })
    });
    // await createrFolder(Site_Relative_Links+"/Shared Documents/", selectedYear+selectedKey)
    // 等待所有文件上传完成
    Promise.all(uploadPromises).then(() => {
      alert("All cost summaries are generated and uploaded successfully.");
      setSelectedKey("")
    }).catch(err => {
      console.log("An error occurred during uploading:", err);
    });
  };


  // 应用单价（每人）表
  useEffect(() => {
    initData().then(res => res).catch(err => err)
  }, [])

  const handleCreateFolder = async () => {
    if (!selectedYear) {
        alert("请选择年份名称");
        return;
    }
    if (!selectedKey) {
      alert("请选择月份");
      return;
  }
    try {
        await createrFolder(Site_Relative_Links+"/Shared Documents", selectedYear+selectedKey);
        alert("文件夹创建成功!");
        await handleExport();
    } catch (error) {
        alert("创建文件夹失败: " + error.message);
    }
};

  return (
    <>
      <h1>Business System Cost Calculation</h1>


        {/* <YearPicker startYear={2023} endYear={2030}  onYearChange={handleYearChange}/> */}
        <Stack horizontal style={{width:600}}>
          <Label style={{marginRight:10,width:100,whiteSpace:'nowarp'}}>Select Period</Label> 
          <Dropdown 
           options={options}
           styles={dropdownStyles}
          />
          <Label>Period Details: {selectedKey}</Label>
        </Stack>
        <Stack horizontal style={{width:600 ,marginTop:10}}>
          <Label style={{marginRight:10,width:100,whiteSpace:'nowarp'}}>Select Market</Label>
          <Dropdown
          options={options}
          styles={dropdownStyles}
          />
        </Stack>
        <Dropdown
          placeholder="Select an option"
          label="Choose a quarter:"
          ariaLabel="Choose a quarter"
          options={options}
          styles={dropdownStyles}
          onChange={onChange} // 绑定onChange事件处理器
          selectedKey={selectedKey} // 设置选中项
        />
        <Stack style={{width:230}}>
        <PrimaryButton style={{ marginTop: 10 }} disabled={excel.length === 0 || selectedKey === ''} onClick={handleCreateFolder}>Generate Summary file</PrimaryButton>
        </Stack>
        <Stack style={{width:230}}>
        <PrimaryButton style={{ marginTop: 10 }} disabled={excel.length === 0 || selectedKey === ''} onClick={()=>alert("正在开发")}>View Summary file</PrimaryButton>
        </Stack>
        <Stack style={{width:230}}>
        <PrimaryButton style={{ marginTop: 10 }} disabled={excel.length === 0 || selectedKey === ''} onClick={()=>alert("正在开发")}>Notify Hub Representative</PrimaryButton>
        </Stack>


    </>
  )









})