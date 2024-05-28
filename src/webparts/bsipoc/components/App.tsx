/* eslint-disable */
import * as React from "react";
import { memo, useEffect } from "react";
import { spfi } from "@pnp/sp";
import { getSP } from "../../../common/pnpjsConfig";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import { DefaultButton, Label, mergeStyleSets, PrimaryButton, Stack } from "office-ui-fabric-react";
import * as XLSX from 'xlsx';
import { Dropdown, IDropdownStyles, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { ProgressIndicator } from '@fluentui/react/lib/ProgressIndicator';
// import YearPicker from "./control/YearSelect";
import createrFolder from "./control/CreateFolder"
import * as Excel from 'exceljs';
import AppContext from "../../../common/AppContext";
// import XlxsExcelFromSP from "./xlxsexcel"

import { Modal, Toggle } from '@fluentui/react';
import { mergeStyles } from '@fluentui/react/lib/Styling';

import { useBoolean } from "@fluentui/react-hooks";
import { addRequest, fetchUserGroups } from '../assets/request'
const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 150, margin: 10 },

};

// const options: IDropdownOption[] = [

//   { key: 'Q1', text: 'Q1' },
//   { key: 'Q2', text: 'Q2' },
//   { key: 'Q3', text: 'Q3' },
//   { key: 'Q4', text: 'Q4' },
// ];


export default memo(function App() {
  const [submiting, setSubmiting] = React.useState<boolean>(false)
  //modal
  const [isModalOpen, { setTrue: showModal, setFalse: hideModal }] = useBoolean(false);
  const [isModalOpenhub, { setTrue: showModalhub, setFalse: hideModalhub }] = useBoolean(false);
  const [isModalOpenConfirm, { setTrue: showModalConfirm, setFalse: hideModalconfirm }] = useBoolean(false);
  const [isModalOpenConfirmGenerate, { setTrue: showModalConfirmGenerate, setFalse: hideModalconfirmGenerate }] = useBoolean(false);
  const [isDraggable, { toggle: toggleIsDraggable }] = useBoolean(false);
  const [keepInBounds, { toggle: toggleKeepInBounds }] = useBoolean(false);
  const classNames = mergeStyleSets({
    modal: {
      width: '90%',
      maxwidth: 400,
      margin: 'auto',
      padding: 20, // 增加内边距
      boxSizing: 'border-box', // 确保padding包含在宽度内
    },
    container: {
      display: 'flex',
      justifyContent: 'center',
      alignItems: 'center',

    },
    header: {
      textAlign: 'center',
    },
    paragraph: {
      textAlign: 'left',
      wordWrap: 'break-word',
      whiteSpace: 'normal',
    }, buttonContainer: {
      display: 'flex',
      justifyContent: 'flex-end',
      marginTop: 10,
    },
    button: {
      marginLeft: 10,
      marginRight: 10
    }
  });
  //
  const sp = spfi(getSP());
  const ctx = React.useContext(AppContext);
  const webURL = ctx.context?._pageContext?._web?.absoluteUrl;
  const userEmail = ctx.context?._pageContext?._user?.email;
  const Site_Relative_Links = webURL.slice(webURL.indexOf('/sites'))
  const [nameObj, setNameObj] = React.useState<any>({})
  // console.log("site",Site_Relative_Links)
  // console.log("weburl",webURL,userEmail)
  const [excel, setExcel] = React.useState([]);
  const [allCountry, setAllCountry] = React.useState([])
  const [allCountryandHub, setallCountryandHub] = React.useState([])
  const [priceTable, setPrice] = React.useState({})
  const [selectedKey, setSelectedKey] = React.useState<string>('');
  // 判断是否已经生成过文件
  const [fileExistState, setfileExistState] = React.useState(false);
  // 生成打开文件链接
  const [filelink, setfilelink] = React.useState("")
  // 仍然生成文件
  // const [generareFileAgain, setGenerareFileAgain] = React.useState(false);
  // Period 选项
  const [selectedKeyPeriod, setSelectedKeyPeriod] = React.useState<string>("");
  const [periodNameOption, setPeriodNameOption] = React.useState<IDropdownOption[]>()
  const [NumberofMonths, setNumberOfMonth] = React.useState(1)
  const [periodDetails, setperiodDetails] = React.useState([])
  //HUb info
  const [hubinfosp, sethubinfosp] = React.useState([])
  const handleDropdownChange_Period = (event: React.FormEvent<HTMLDivElement>, item?: IDropdownOption): void => {
    if (item) {
      setSelectedKeyPeriod(item.key as string);
      setSelectedKey(item.text as string);
      const value = doesFolderExist("Shared Documents", item.text).then(exsit => { console.log("value", exsit); setfileExistState(exsit) })

    }

  };
  // Market 选项
  const [selectedKeyMarket, setSelectedKeyMarket] = React.useState<string>('');
  const [marketNameOption, setMarketNameOption] = React.useState<IDropdownOption[]>()
  const handleDropdownChange_Market = (event: React.FormEvent<HTMLDivElement>, item?: IDropdownOption): void => {
    if (item) {
      setSelectedKeyMarket(item.key as string);
    }
  };


  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  function calcToExcel(orders: any, priceMap: any, map: any) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return orders.map((val: any) => {
      const obj: any = {
        Market: val.Market,
        Dealer: val.PartnerName,
        "Dealer ID": val.PartnerID,
        // "Dealer Type": val.DealerType,
        "PartnerType":val.PartnerType,
        "Dealer Category": val.DealerCategory,
        "Basic Package": val.BasicPackage,
        "Sales Package": val.SalesPackage,
        "Hub Package": val.HubPackage,
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
      const keys = ['Hub Package', 'CPQ', 'UD CM', 'Argus 365', 'UDCP', 'SeMA', 'LDS', 'LSS', 'Pardot']
      obj['Total(Per Month)'] = keys.reduce((sum, key) => sum + (priceMap[key] || 0) * obj[key], 0)
      obj['Total(Per Month)'] -= 100 * Math.min(obj['LDS'], obj['LSS'])
      // obj['Total(Per Month)'] += (priceMap['Basic Package;' + val.DealerCategory] || 0) * obj['Basic Package']
      // obj['Total(Per Month)'] += (priceMap['Sales Package;' + (val.DealerCategory || 'NA')] || 0) * obj['Sales Package']
      obj['Total(Per Month)'] += (priceMap['Basic Package;' + val.DealerCategory] || 0) * obj['Basic Package']
      obj['Total(Per Month)'] += (priceMap['Sales Package;' + (val.PartnerType)] || 0) * obj['Sales Package']
      obj['Total(Per Month)'] = obj['Total(Per Month)'].toFixed(2)
      obj['Hub'] = val.Hub
      // console.log("obj3232",obj)
      // console.log("price,ap",priceMap)
      return obj
    })
  }

  function calcToSummary(details: any, priceMap: any, map: any, period: any) {
    // 添加了从period 获得的月份 在计算数量和总价格的时候用到了
    let numMonth = 1
    if (selectedKey !== "") {
      console.log(period, selectedKeyPeriod)
      numMonth = period.filter((per: any) => per.text === selectedKey)[0].nummonth
      console.log("nummonth1", numMonth)
    }
    console.log("nummonth2", numMonth)

    const resObj: any = {
      Country: details[0].Market,
      data: []
    }
    // console.log("res",resObj)
    const p = { ...priceMap }
    for (let key in p) {
      const value = p[key]
      p[key] = {
        price: value,
        count: 0,
        total: 0
      }
    }
    details.map((val: any) => ({ ...val })).forEach((val: any) => {
      val['Basic Package;' + val['Dealer Category']] = val['Basic Package']
      val['Sales Package;' + (val['PartnerType'])] = val['Sales Package']
      val['LDS+LSS'] = Math.min(val['LDS'], val['LSS'])
      for (let key in p) {
        p[key].count += Number(val[key] || 0)
      }
    })
    // console.log("selectedKeyPeriod", periodDetails, ["3434"])
    //const numMonth = periodDetails.filter(item => item.periodDetails === selectedKey)
    //console.log("nm,omth", numMonth)
    for (let key in p) {
      const isLDSorLSS = key === 'LDS' || key === 'LSS';

      // 检查 LDS 和 LSS 是否都有值
      const bothLDSandLSSHaveValues = p['LDS'].count > 0 && p['LSS'].count > 0;
      if (p[key].count === 0 || (isLDSorLSS && bothLDSandLSSHaveValues)) continue;

      resObj.data.push({
        // A: key.split(';')[0],
        // B: key.split(';').length > 1 ? '- ' + key.split(';')[1] : '',
        A: map[key],
        B: null,
        C: p[key].count * numMonth,
        D: (p[key].price).toFixed(2),
        E: (p[key].count * numMonth * p[key].price).toFixed(2)
      })
    }
    resObj.data = resObj.data.filter((item: any) => item.C > 0)
    //console.log("data",resObj.data)
    //

    const total = resObj.data.reduce((t: number, e: any) => t + Number(e.E), 0)
    resObj.data.push({
      A: 'Total',
      E: total.toFixed(2)
    })

    // 分离出包含"Package"的项和其他项
    const packages = resObj.data.filter((item: any) => item.A.toLowerCase().includes("package"));
    const otherItems = resObj.data.filter((item: any) => !item.A.toLowerCase().includes("package") && item.A !== 'Total');

    // 对packages和其他项进行排序
    packages.sort((a: any, b: any) => a.A.localeCompare(b.A));
    otherItems.sort((a: any, b: any) => a.A.localeCompare(b.A));

    // 合并结果并将Total项放在最后
    resObj.data = [...packages, ...otherItems, { A: 'Total', E: total.toFixed(2) }];
    return resObj
  }



  const changeStyle = async (buffer: Excel.Buffer) => {
    const workbook = new Excel.Workbook();
    await workbook.xlsx.load(buffer); // 加载Excel文件
    const worksheet = workbook.getWorksheet(1); // 获取第一个工作表

    ['A4', 'B4', 'C4', 'D4', 'E4'].forEach(cell => {
      worksheet.getCell(cell).fill = {
        type: 'pattern',
        pattern: 'solid',
        bgColor: { argb: 'FFA1C1E5' }
      }
    })

    const worksheetDetail = workbook.getWorksheet(2)
    const zimu = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'N', 'M', 'O', 'P', "Q"]
    zimu.forEach(z => {
      worksheetDetail.getCell(z + '2').fill = {
        type: 'pattern',
        pattern: 'solid',
        bgColor: { argb: 'FFA1C1E5' }
      }
      worksheetDetail.getCell(z + '3').fill = {
        type: 'pattern',
        pattern: 'solid',
        bgColor: { argb: 'FFA1C1E5' }
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

    const buffer = await sp.web.getFileByServerRelativePath(Site_Relative_Links + "/BSITemplate/UD BSI_Output Template.xlsx").getBuffer();

    let selectCountry = allCountry.slice(0)
    if (selectedKeyMarket !== "" && selectedKeyMarket !== "ALL") {
      console.log(selectedKeyMarket, "erer")
      selectCountry = allCountry.filter(country => country === selectedKeyMarket)
    }

    // 遍历所有国家
    selectCountry.forEach(Market => {
      // 筛选出该国家的订单
      // console.log("2exe",excel)
      const countryOrders = excel.filter(order => order.Market === Market);
      if (countryOrders.length === 0) {
        console.log(`No data for ${Market}`);
        return;
      }
      console.log("country", countryOrders)
      /* summary */
      const workbookTemplate = XLSX.read(buffer, { type: 'buffer' });
      const summaryTemplateName = workbookTemplate.SheetNames[1]
      const workSheetSummaryTpt = workbookTemplate.Sheets[summaryTemplateName]
      // const arrTpt = XLSX.utils.sheet_to_json(workSheetSummaryTpt)

      const tongji = calcToSummary(countryOrders, priceTable, nameObj, periodDetails)
      // \console.log("tongji",tongji)
      workSheetSummaryTpt['B2'] = { v: tongji.Country }
      // workSheetSummaryTpt['E2'] = { v: `${selectedYear}/${(Number(selectedKey.replace('Q', '')) -1 )*3+1} - ${selectedYear}/${(Number(selectedKey.replace('Q', '')))*3}` }
      workSheetSummaryTpt['E2'] = { v: selectedKeyPeriod }
      for (let i = 5; i < tongji.data.length + 5; i++) {
        workSheetSummaryTpt['A' + i] = { v: tongji.data[i - 5].A }
        workSheetSummaryTpt['B' + i] = { v: tongji.data[i - 5].B }
        workSheetSummaryTpt['C' + i] = { v: tongji.data[i - 5].C }
        workSheetSummaryTpt['D' + i] = { v: tongji.data[i - 5].D }
        workSheetSummaryTpt['E' + i] = { v: tongji.data[i - 5].E }
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
      const zimu = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']
      for (let i = 3; i < countryOrders.length + 3; i++) {
        let j = 0;
        for (let key in countryOrders[i - 3]) {
          workSheetDetails[zimu[j] + (i + 1)] = { v: countryOrders[i - 3][key] }
          j++
        }
      }


      const total = countryOrders.reduce((t: number, item: any) => t + Number(item['Total(Per Month)']), 0)
      workSheetDetails['A' + (countryOrders.length + 4)] = { v: 'Total' }
      workSheetDetails['Q' + (countryOrders.length + 4)] = { v: total.toFixed(2) }

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
        doesFolderExist("Shared Documents", `${selectedKey}/${countryOrders[0].Hub}`).then(async exists => {
          if (!exists) {
            createrFolder(Site_Relative_Links + `/Shared Documents/${selectedKey}`, countryOrders[0].Hub)
          }
        }).then(() => {
          return new Promise<void>((resolve) => {
            // setTimeout(() => {
            //   resolve();
            // }, 3000);
            resolve();
          });
        })
          .then(() => {
            // 添加上传任务到数组
            const uploadPromise = uploadFileToSP(
              `${Site_Relative_Links}/Shared Documents/${selectedKey}/${countryOrders[0].Hub}`,
              `UD ${Market} ${selectedKey}.xlsx`,
              blob
            );
            uploadPromises.push(uploadPromise);
          })

      })
    });
    // 等待所有文件上传完成
    Promise.all(uploadPromises).then(() => {
      setTimeout(() => {
        // alert("All cost summaries are generated and uploaded successfully.");
        showModalConfirmGenerate()
      }, 500);
      const value = doesFolderExist("Shared Documents", selectedKey).then(exsit => { console.log("value", exsit); setfileExistState(exsit) })
      // setSelectedKey("")
    }).catch(err => {
      console.log("An error occurred during uploading:", err);
    });
  };
  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  const initData = async () => {
    let obj: any = {}
    let perioddetails_init = {}
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
                        <FieldRef Name="Description"/>
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
          obj[val.ApplicationName] = val.Description
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
                        <FieldRef Name="Description"/>
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
          
          if(val.PackageCategory === "Sales Package"){ 
            resObj[`${val.PackageCategory};${val.PartnerType}`] = val.MonthlyPrice_x0028_USD_x0029_ * 1
            obj[`${val.PackageCategory};${val.PartnerType}`] = val.Description
          }else{
            obj[`${val.PackageCategory};${val.DealerCategory}`] = val.Description
            resObj[`${val.PackageCategory};${val.DealerCategory}`] = val.MonthlyPrice_x0028_USD_x0029_ * 1
          }
          
        })
        // console.log("resobji", resObj)
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
      if (response.Row?.length > 0) {

        setPeriodNameOption(
          response.Row.filter(item => item.AvailableforSelection === "Yes").map(period => ({
            key: period.PeriodDetails,
            text: period.PeriodName
          })))
        const details = response.Row.filter(item => item.AvailableforSelection === "Yes").map(period => ({
          key: period.PeriodDetails,
          text: period.PeriodName,
          nummonth: period.NumberofMonths
        }))
        perioddetails_init = details
        console.log("erioddetails", perioddetails_init)
        setperiodDetails(details)


      }
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


    // 拿 Hub Representative 周期对应关系
    const hub = await sp.web.lists.getByTitle("Hub Representative").renderListDataAsStream({

      ViewXml: `<View>
                      <ViewFields>
                        <FieldRef Name="Hub"/>
                        <FieldRef Name="HubReprensentative"/>
                        <FieldRef Name="Copyto"/>
                      </ViewFields>
                   
                    </View>`,
      // <RowLimit>400</RowLimit>
    }).then((response) => {
      console.log("hub", response.Row)
      if (response.Row?.length > 0) {
        const hubinfo = response.Row.map(item => ({
          "Hub": item.Hub, HubReprensentative: item.HubReprensentative.map((rep: any) => ({
            id: rep.id,
            value: rep.value,
            title: rep.title,
            email: rep.email,
            sip: rep.sip,
            picture: rep.picture,
            jobTitle: rep.jobTitle,
            department: rep.department,
          }))
        }))
        sethubinfosp(hubinfo)
      }
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
        // console.log("country", uniqueList)
        setAllCountry(uniqueList)
        const uniqueListAndHub = Array.from(new Set(response.Row.map(item => JSON.stringify({ "market": item.Market, "Hub": item.Hub })))).map(item => JSON.parse(item));
        // console.log("origin",uniqueListAndHub)
        setallCountryandHub(uniqueListAndHub)

        const uniqueMarket = Array.from(new Set(uniqueList))

        uniqueMarket.push("ALL");
        // console.log("hhh",uniqueMarket);

        setMarketNameOption(uniqueMarket.map(market => ({
          key: market,
          text: market
        })).sort((a, b) => a.text.localeCompare(b.text)))
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
    setNameObj(obj)
    console.log('price', price)
    const finalExcelData = calcToExcel(order, price, obj)

    console.log(calcToSummary(finalExcelData, price, obj, period))
    console.log("excel121", finalExcelData)
    setExcel(finalExcelData)
  }
  async function doesFolderExist(libraryName: string, folderName: string): Promise<boolean> {
    try {
      // 尝试获取目标文件夹的属性
      const folder = await sp.web.getFolderByServerRelativePath(`${Site_Relative_Links}/${libraryName}/${folderName}`).select("Exists")();
      return folder.Exists;
    } catch (error) {
      // 如果抛出404错误，表示文件夹不存在
      if (error.status === 404) {
        return false;
      } else {
        throw error; // 处理其他可能的错误
      }
    }
  }
  // 应用单价（每人）表
  useEffect(() => {
    initData().then(res => res).catch(err => err)
  }, [])
  //   useEffect(() => {
  //     handleDropdownChange_Market(null, marketNameOption[marketNameOption?.length -1]);
  // }, []);
  const handleCreateFolder = async (genarateAgain: boolean = false) => {
    if (!selectedKeyPeriod) {
      alert("Please choose period");
      return;
    }
    doesFolderExist("Shared Documents", selectedKey).then(async exists => {

      console.log(`Folder exists: ${exists}`);
      if (!exists || genarateAgain) {
        try {
          await createrFolder(Site_Relative_Links + "/Shared Documents", selectedKey);
          // alert("Folder created successfully");
          // setTimeout(async () => {
          //   await handleExport();
          // }, 3000);
          await handleExport();
        } catch (error) {
          alert("Failed to create folder: " + error.message);
        }
        hideModal()
      }
      else {
        showModal()
      }
    }).catch(error => {
      console.error(`Error: ${error}`);
    });

  };
  const submitform = () => {
    if (submiting) return
    setSubmiting(true)
    let a = hubinfosp.filter((item) => {
      if (selectedKeyMarket === "ALL") {
        return true;
      } else {
        return item.Hub === allCountryandHub.find((hub) => hub.market === selectedKeyMarket)?.Hub;
      }
    })
    // console.log("aaaaa", a)
    const request = {
      Hub: selectedKeyMarket === "ALL" ? "All Hub" : allCountryandHub.find(hub => hub.market === selectedKeyMarket)?.Hub,
      FileMainLink: filelink,

    }

    const sp = spfi(getSP());
    // let promiss
    addRequest({ request }).then(async promises => {
      console.log("promiss", promises, typeof (promises));
      hideModalhub()
      showModalConfirm()
      setSubmiting(false)
    })



  }
  useEffect(() => {
    // 模拟组件加载后触发 onChange 事件
    if (marketNameOption && marketNameOption.length > 0) {
      handleDropdownChange_Market(null, marketNameOption[0]);
    }
  }, [marketNameOption]);

  useEffect(() => {
    let filelink;
    if (selectedKeyMarket === "ALL") {
      filelink = `${Site_Relative_Links}/Shared Documents/${selectedKey}/`
    } else {
      filelink = `${Site_Relative_Links}/Shared Documents/${selectedKey}/${selectedKeyMarket === "ALL" ? "All Hub" : allCountryandHub.find(hub => hub.market === selectedKeyMarket)?.Hub}`
    }
    setfilelink(filelink)
  }, [selectedKey, selectedKeyMarket])
  return (
    <>
      <h1 style={{ margin: 10 }}>Business System Cost Calculation</h1>



      <Stack horizontal style={{ width: 600, marginLeft: 10 }}>
        <Label style={{ marginTop: 10, width: 100, whiteSpace: 'nowarp' }}>Select Period</Label>
        <Dropdown
          options={periodNameOption}
          styles={dropdownStyles}
          onChange={handleDropdownChange_Period}
        />
        {selectedKeyPeriod && <Label style={{ marginTop: 10 }}>Period Details: {selectedKeyPeriod}</Label>}
      </Stack>
      <Stack horizontal style={{ width: 600, marginLeft: 10 }}>
        <Label style={{ marginTop: 10, width: 100, whiteSpace: 'nowarp' }}>Select Market</Label>
        <Dropdown
          options={marketNameOption}
          styles={dropdownStyles}
          onChange={handleDropdownChange_Market}
          // selectedKey={selectedKeyMarket}
          defaultSelectedKey={"ALL"}
        />

        {selectedKeyMarket && <Label style={{ marginTop: 10 }}>Hub Details: {selectedKeyMarket === "ALL" ? "All Hub" : allCountryandHub.find(hub => hub.market === selectedKeyMarket)?.Hub}</Label>}
      </Stack>
      {/* <Stack style={{marginLeft:30}}>
          This is for Test
        <YearPicker startYear={2023} endYear={2030}  onYearChange={handleYearChange}/>
        <Dropdown
          placeholder="Select an option"
          label="Choose a quarter:"
          ariaLabel="Choose a quarter"
          options={options}
          styles={dropdownStyles}
          onChange={onChange} // 绑定onChange事件处理器
          selectedKey={selectedKey} // 设置选中项
        />

</Stack> */}
      <Stack style={{ margin: 10, width: 230 }}>
        {/* //{selectedKeyPeriod} */}
        <PrimaryButton style={{ marginTop: 10 }} disabled={excel.length === 0 || selectedKeyPeriod === null || selectedKeyPeriod === ""} onClick={() => handleCreateFolder()}>Generate Summary File </PrimaryButton>
      </Stack>
      <Stack style={{ margin: 10, width: 230 }}>
        <PrimaryButton style={{ marginTop: 10 }} disabled={excel.length === 0 || selectedKey === '' || !fileExistState}
          onClick={() => window.open(filelink, "_blank")}>View Summary File
        </PrimaryButton>
      </Stack>
      <Stack style={{ margin: 10, width: 230 }}>
        <PrimaryButton style={{ marginTop: 10 }} disabled={excel.length === 0 || selectedKey === '' || !fileExistState} onClick={showModalhub}>Notify Hub Representative</PrimaryButton>
      </Stack>
      <Stack>
        {/* <ProgressIndicator label="Uploading files now" description="Example description" /> */}
        {/* <Toggle label="Is draggable" inlineLabel onChange={toggleIsDraggable} checked={isDraggable} /> */}
        {/* <DefaultButton onClick={showModalhub} text="Open Modal" /> */}
        <Modal
          titleAriaId={"title"}
          isOpen={isModalOpen}
          // onDismiss={hideModal}
          isBlocking={false}
          containerClassName={classNames.container}
        // dragOptions={isDraggable ? dragOptions : undefined}
        >
          {/* <Stack horizontalAlign="center" > */}
          <h2 className={classNames.header}>Warning</h2>
          {/* </Stack> */}
          <p className={classNames.paragraph}>
            Summary file for selected Period or Market has been </p> <p>generated.
              Choosing to re-generate will overwrite the existing files.</p>
          <p className={classNames.paragraph}>Please confirm if you wish to proceed.</p>
          <div className={classNames.buttonContainer}>
            <PrimaryButton className={classNames.button} onClick={() => handleCreateFolder(true)}>Yes</PrimaryButton>
            <DefaultButton className={classNames.button} onClick={hideModal}>No</DefaultButton>
          </div>
        </Modal>
        <Modal
          titleAriaId={"hub"}
          isOpen={isModalOpenhub}
          // onDismiss={hideModal}
          isBlocking={false}
          containerClassName={classNames.container}
        // dragOptions={isDraggable ? dragOptions : undefined}
        >
          {/* <Stack horizontalAlign="center" > */}
          <h2 className={classNames.header}>Hub Notify</h2>
          <p>You are going to send email notification to cantacts below.Please confirm if you wish to proceed</p>
          {/* </Stack> */}
          <Label>Summary File:    {selectedKey} </Label>

          <ul>


            {
              hubinfosp
                .filter((item) => {
                  if (selectedKeyMarket === "ALL") {
                    return true;
                  } else {
                    return item.Hub === allCountryandHub.find((hub) => hub.market === selectedKeyMarket)?.Hub;
                  }
                }).map((hubInfo, index) => (
                  //  hubinfosp.filter(item=>{if(selectedKeyMarket!=="All"){console.log("item.hub",item.Hub,allCountryandHub.find(hub => hub.market === selectedKeyMarket)?.Hub);item.Hub === allCountryandHub.find(hub => hub.market === selectedKeyMarket)?.Hub}}).map((hubInfo, index) => (
                  <li key={index}>
                    Hub: {hubInfo.Hub}
                    <ul>
                      {hubInfo.HubReprensentative.map((rep: any, repIndex: number) => (
                        <li key={repIndex}>

                          <p>Email: {rep.email}</p>

                        </li>
                      ))}
                    </ul>
                  </li>
                ))}
          </ul>
          <div className={classNames.buttonContainer}>
            <PrimaryButton className={classNames.button} onClick={submitform}>Yes</PrimaryButton>
            <DefaultButton className={classNames.button} onClick={hideModalhub}>No</DefaultButton>
          </div>
        </Modal>
        <Modal
          titleAriaId={"confirm"}
          isOpen={isModalOpenConfirm}
          // onDismiss={hideModal}
          isBlocking={false}
          containerClassName={classNames.container}
        // dragOptions={isDraggable ? dragOptions : undefined}
        >
          {/* <Stack horizontalAlign="center" > */}
          <h2 className={classNames.header}>Notice</h2>
          {/* </Stack> */}
          <p className={classNames.paragraph}>
            A record has been created and the message will be sent in a few minutes</p>
          <div className={classNames.buttonContainer}>
            {/* <PrimaryButton className={classNames.button} onClick={() => handleCreateFolder(true)}>Yes</PrimaryButton> */}
            <DefaultButton className={classNames.button} onClick={hideModalconfirm}>OK</DefaultButton>
          </div>
        </Modal>
        <Modal
          titleAriaId={"notify generate"}
          isOpen={isModalOpenConfirmGenerate}
          // onDismiss={hideModal}
          isBlocking={false}
          containerClassName={classNames.container}
        // dragOptions={isDraggable ? dragOptions : undefined}
        >
          {/* <Stack horizontalAlign="center" > */}
          <h2 className={classNames.header}>Notice</h2>
          {/* </Stack> */}
          <p className={classNames.paragraph}>
          Summary files are generated successfully. Please click "View Summary File" button to check the files</p>
          <div className={classNames.buttonContainer}>
            {/* <PrimaryButton className={classNames.button} onClick={() => handleCreateFolder(true)}>Yes</PrimaryButton> */}
            <DefaultButton className={classNames.button} onClick={hideModalconfirmGenerate}>OK</DefaultButton>
          </div>
        </Modal>
      </Stack>
      {/* {periodDetails} */}

    </>
  )









})