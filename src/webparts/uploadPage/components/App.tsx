/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable dot-notation */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable */
import * as React from "react";
import { memo, useContext } from "react";
//import { read, utils, SSF }  as XLSX from 'xlsx';
import * as XLSX from 'xlsx';
import { useState } from "react";
import { Button } from "antd";
import { addRequest, fetchUserGroups } from "./utils/request";
import 'antd/dist/antd.css';
import { Stack } from '@fluentui/react/lib/Stack';
// import { Label } from '@fluentui/react/lib/Label';
import { Dropdown, Icon, IDropdownOption, IDropdownStyles, Label } from "office-ui-fabric-react";
// import { Icon as IconBase } from '@fluentui/react/lib/Icon';
import { Upload, Modal } from 'antd';
// import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { useEffect } from "react";

import { spfi } from "@pnp/sp";
import { getSP } from "../../../common/pnpjsConfig";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/batching";
import styles from './UploadPage.module.scss'
import FileSvg from '../assets/file'
import Del from '../assets/delete'
import Error from '../assets/error'
import AppContext from "../../../common/AppContext";
import Viewhistory from '../assets/submit';
// import HintIcon from "../assets/hinticon"
import "./App.css";


import { countByRules } from "./utils/handledata"




interface Address {
  addressType: string;
  addressLine: string;
  inCareOf: string;
  street: string;
  houseNr: string;
  poBox: string;
  city: string;
  postalCode: string;
  district: string;
  poBoxCity: string;
  poBoxPostalCode: string;
  countryCode: string;
  countryName: string;
  regionCode: string;
  regionName: string;
  phoneNumber: string;
  faxNumber: string;
  email: string;
}
interface InternationalVersion {
  address: Address[];
}
interface JsonData {
  id: string;
  parentParmaId: string;
  parmaID: string;
  status: string;
  creationDate: string;
  updatedDate: string;
  communicationLanguage: string;
  internationalVersion: InternationalVersion;
  // ... 其他属性
}
interface DataRecord {
  PartnerID: number;
  Application: string;
  MachineNo: string;
  Description?: string;
  AmountInUSD: number;
  Market?: string;
  Period?:string;
  "HWI" :number,
   "S410 W VOCOM": number;
  "S410 W/O VOCOM": number;
  "V110 W VOCOM": number;
  "V110 W/O VOCOM": number;
  // PartnerID: string;
}
function sanitize(input: string) {
  // based on https://support.microsoft.com/en-us/help/905231/information-about-the-characters-that-you-cannot-use-in-site-names--fo
  // replace invalid characters
  let sanitizedInput = input.replace(/['~"#%&*:<>?/{|}]/g, "_");
  // replace consecutive periods
  sanitizedInput = sanitizedInput.replace(/\.+/g, ".");
  // replace leading period
  sanitizedInput = sanitizedInput.replace(/^\./, "");
  // replace leading underscore
  sanitizedInput = sanitizedInput.replace(/^_/, "");
  return sanitizedInput;
}

const dropdownStyles: Partial<IDropdownStyles> = {
  root: { background: '#fff', display: 'flex', flexShrink: 0, alignItems: 'center', width: 150, marginRight: 60, fontSize: '14px', height: 30, color: '#191919', border: '1px solid #454545', borderRadius: '10px' },
  dropdown: { ':focus::after': { border: 'none' }, width: 230 },
  title: { border: 'none', background: 'none' }
};

export default memo(function App() {
  const sp = spfi(getSP());
  const [items, setItems] = useState([]);
  const [data, setData] = useState([]);
  const [error, setError] = useState('');
  const [showBtn, setShowBtn] = useState(false)
  const [uploadFile, setFile] = useState<any>()
  const [isShowModal, setIsShowModal] = useState(false)
  // const [apiResponse, setApiResponse] = useState<any>(null);
  const [submiting, setSubmiting] = React.useState<boolean>(false)

  const [fileWarning, setfileWarning] = React.useState("")
  const [buttonvisible, setbuttonVisible] = React.useState<boolean>(true)
  const ctx = useContext(AppContext);
  const userEmail = ctx.context?._pageContext?._user?.email;
  const webURL = ctx.context?._pageContext?._web?.absoluteUrl;

  const [selectedKeyPeriod, setSelectedKeyPeriod] = React.useState<string>("");
  const [periodNameOption, setPeriodNameOption] = React.useState<IDropdownOption[]>()
  const [selectedKey, setSelectedKey] = React.useState<string>('');
  const [periodDetails, setperiodDetails] = React.useState([])

  const Site_Relative_Links = webURL.slice(webURL.indexOf('/sites'))

  const [fileExistFlag, setfileExistFlag] = React.useState<Boolean>(false)

  const [groupedByMarketlist,setgroupedByMarketlist] = React.useState()
  const handleDropdownChange_Period = (event: React.FormEvent<HTMLDivElement>, item?: IDropdownOption): void => {
    if (item) {
      setSelectedKeyPeriod(item.key as string);
      setSelectedKey(item.text as string);
      //   const value = doesFolderExist("Shared Documents", item.text).then(exsit => { console.log("value", exsit); setfileExistState(exsit) })


    }

  };
  // 测试读取特定的excel 文件
  async function readExcelFromLibrary(fileName: string) {
    try {
      const file = await sp.web.getFileByServerRelativePath(Site_Relative_Links + "/VCAD Documents/2024Q1.xlsx").getBuffer();

      const workbook = XLSX.read(file, { type: "buffer" });

      // 假设我们读取第一个工作表
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      // 将工作表转换为 JSON 数据
      const jsonData = XLSX.utils.sheet_to_json(worksheet);

      console.log(jsonData);
      return jsonData;
    } catch (error) {
      console.error("Error reading Excel file from library", error);
    }
  }
  interface VcadItem {
    Title: string;
    // 添加其他字段，例如：
    Market: string;
    PartnerID: string;
    'S410 W VOCOM': number;
    'S410 W/O VOCOM': number;
    'V110 W VOCOM': number;
    'V110 W/O VOCOM': number;
    HWI:number;
    Period:string
    // 根据你的需求添加更多字段
  }
  // 多个项目提交
  async function addMultipleItems(Vcadlist:VcadItem[]) {
    try {
      const [batchedSP, execute] = sp.batched();
      const list = batchedSP.web.lists.getByTitle("VCAD Summary");

      let res: any[] = [];

      for (let i = 0; i < Vcadlist.length; i++) {
        const item = Vcadlist[i];
        list.items.add({
          Title: item.Title,
          Market:item.Market,
          Partner_x0020_ID:item.PartnerID,
          S410WVOCOM:item["S410 W VOCOM"],
          S410W_x002f_OVOCOM:item["S410 W/O VOCOM"],
          V110WVOCOM:item["V110 W VOCOM"],
          V110W_x002f_OVOCOM:item["V110 W/O VOCOM"],
          HWI:item.HWI,
          Period:item.Period


          // Add other fields as needed

        }).then(r => res.push(r));
      }

      // Executes the batched calls
      await execute();

      // Results for all batched calls are available
      for (let i = 0; i < res.length; i++) {
        console.log(res[i]); // or do something with the results
      }

      console.log(Vcadlist.length+ " items added successfully");
    } catch (error) {
      console.error("Batch execution failed", error);
    }
  }


  interface IItem {
    ID: number;
    Period: string;
    // 其他字段
  }
  
  async function getAllItemsByPeriod(listTitle: string, periodValue: string): Promise<IItem[]> {
    const sp = spfi(getSP());
    try {
      let items: IItem[] = [];
      let pager = await sp.web.lists
        .getByTitle(listTitle)
        .items
        .filter(`Period eq '${periodValue}'`)  // 根据 Period 字段进行过滤
        .select("ID", "Period")  // 选择必要的字段
        .top(5000)  // 设置分页大小
        .orderBy("ID", true)  // 确保分页排序一致性
        .getPaged();
  
      items = items.concat(pager.results);
  
      while (pager.hasNext) {
        const response = await pager.getNext();
        items = items.concat(response.results);
        pager = response;
      }
  
      console.log("Fetched Items:", items);
  
      return items;
    } catch (err) {
      console.error("Error when fetching items by period:", err);
      return Promise.reject("Error when fetching items by period");
    }
  }
  
  async function deleteItemsByPeriod(periodValue: string) {
    try {
      const listTitle = "VCAD Summary";
      const itemsToDelete = await getAllItemsByPeriod(listTitle, periodValue);
  
      if (itemsToDelete.length === 0) {
        console.log(`No items found with Period = ${periodValue}`);
        return;
      }
  
      // 使用批处理删除
      const sp = spfi(getSP());
      const [batchedSP, execute] = sp.batched();
      const batchedList = batchedSP.web.lists.getByTitle(listTitle);
  
      let res: any[] = [];
  
      for (let i = 0; i < itemsToDelete.length; i++) {
        const itemId = itemsToDelete[i].ID;
        batchedList.items.getById(itemId).delete().then(r => res.push(r));
      }
  
      // 执行批处理
      await execute();
  
      // 处理结果
      for (let i = 0; i < res.length; i++) {
        console.log(res[i]);
      }
  
      console.log(itemsToDelete.length + " items deleted successfully");
    } catch (error) {
      console.error("Batch execution failed", error);
    }
  }
  

  


  // 校验方法
  const validateHeaders = (headers: string[]): boolean => {
    const requiredHeaders = ['Dealer ID', 'Application', 'Machine No', 'Description', 'Amount in USD'];
    return requiredHeaders.every(header => headers.indexOf(header) !== -1);;
  };

  async function doesFolderExist(folderName: string): Promise<boolean> {
    try {
      // 尝试获取目标文件夹的属性
      const folder = await sp.web.getFileByServerRelativePath(`${Site_Relative_Links}/VCAD Documents/${folderName}`).select("Exists")();
      console.log("filename", folderName, folder, folder.Exists)
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

  const handleFileUpload = (info: any) => {
    // 重置状态
    setItems([]);
    setData([]);
    setError(null);
    setFile(null); // 确保每次上传前文件状态都被重置
    setShowBtn(false); // 根据需要可能还需要重置其他UI状态

    if (info.file) {
      const file = info.file;
      if (!file) return;

      const reader = new FileReader();
      reader.onload = (e: ProgressEvent<FileReader>) => {
        const binaryStr = e.target?.result;
        if (typeof binaryStr === 'string') {
          const workbook = XLSX.read(binaryStr, { type: 'binary' });
          const worksheet = workbook.Sheets[workbook.SheetNames[0]];
          const jsonData: unknown[] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

          // 调用校验方法
          const headers = jsonData[0] as string[];
          const isValid = validateHeaders(headers);
          if (!isValid) {
            setError('The first line of the file must contain Dealer ID, Application, Machine No, Description, Amount in USD');
            return;
          }
          setFile(file)
          setShowBtn(true)
          // 将数据转换为 DataRecord 格式并设置到状态中
          const exceldata = (XLSX.utils.sheet_to_json(worksheet) as any[]).map(row => ({
            DealerID: row['Dealer ID'],
            Application: row['Application'],
            MachineNo: row['Machine No'],
            Description: row['Description'],
            AmountInUSD: row['Amount in USD'],
            Market: row['Market'],
            // PartnerID: row['Partner ID'],
        }));
        console.log("sek",selectedKey)
        const result = countByRules(exceldata,selectedKey);

        console.log(result);
        setData(result);
    }
};
      reader.onerror = (error) => {
        setError('文件读取失败'); // 处理文件读取错误
      };
      reader.readAsBinaryString(file);
    }
  };


  // 提交方法
  const submitFunction = async (): Promise<void> => {
    if (submiting) return;

    setSubmiting(true);

    try {
      // 上传文件
      if (uploadFile) {
        await sp.web.getFolderByServerRelativePath('VCAD Documents').files.addUsingPath(selectedKey + uploadFile.name.substring(uploadFile.name.lastIndexOf('.')), uploadFile, { Overwrite: true });
      }
      const a = updateResultsWithMarkets(data,groupedByMarketlist).filter(item => 
        (item['S410 W VOCOM']>0 || item['HWI']>0 || item['S410 W/O VOCOM']>0 || item['V110 W VOCOM']>0 || item['V110 W/O VOCOM']>0 ))
      console.log(a)
      if(fileExistFlag){deleteItemsByPeriod(selectedKey)}
      addMultipleItems(a)
      // 操作成功，更新UI状态
      setbuttonVisible(false);
    } catch (error) {
      setError('文件上传失败');
    } finally {
      setSubmiting(false);
    }
  };

  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  const initData = async () => {
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
          response.Row.filter(item => item.AvailableforSelection === "是" || item.AvailableforSelection === "Yes").map(period => ({
            key: period.PeriodDetails,
            text: period.PeriodName
          })))
        const details = response.Row.filter(item => item.AvailableforSelection === "是" || item.AvailableforSelection === "Yes").map(period => ({
          key: period.PeriodDetails,
          text: period.PeriodName,
          nummonth: period.NumberofMonths
        }))
        const perioddetails_init = details
        console.log("erioddetails", perioddetails_init)
        setperiodDetails(details)


      }
    

      return {}
    });
    // 拿主表订单
    const order = await sp.web.lists.getByTitle("Partner Master").renderListDataAsStream({


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
        const uniqueListAndHub = Array.from(new Set(response.Row.map(item => JSON.stringify({ "market": item.Market, "Hub": item.Hub })))).map(item => JSON.parse(item));
        // console.log("origin",uniqueListAndHub)
        

        const uniqueMarket = Array.from(new Set(uniqueList))

        

        
        //setAllCountry
        return response.Row
      }
      return []
    })
    findmarket(order)
  }
// 定义更新方法
const updateResultsWithMarkets = (results: any[], marketDealerMapping: { [market: string]: string[] }) => {
  results.forEach(result => {
    for (const market in marketDealerMapping) {
      if (marketDealerMapping[market].includes(result.PartnerID)) {
        result.Market = market;
        result.Period = selectedKey;
        break;
      }
    }
  });
  console.log("res",results)
  return results;
};
function findmarket (detail:any){
    console.log("Main", detail)
    // 按 Market 分组，并提取 PartnerID 列表
    const groupedByMarket = detail.reduce((acc: any, item: any) => {
      if (!acc[item.Market]) {
        acc[item.Market] = [];
      }
      // 确保 PartnerID 不重复
      if (!acc[item.Market].includes(item.PartnerID)) {
        acc[item.Market].push(item.PartnerID);
      }
      return acc;
    }, {});
    console.log('group', groupedByMarket);
    setgroupedByMarketlist(groupedByMarket)
    // updateResultsWithMarkets(data,groupedByMarket)
  }
  useEffect(() => {
    initData().then(res => res).catch(err => err)
  }, [])

  return (
    <div className={styles.uploadPage}>
{/* <Button onClick = {()=>deleteItemsByPeriod(selectedKey)}> 按钮</Button> */}
      {/* <Button onClick={()=> addMultipleItems().catch(console.error)}> 提交测试</Button>
            <Button onClick={()=> readExcelFromLibrary("222")}> 提交测试</Button> */}
      {
        buttonvisible ? <div className={styles.content}>
          <Stack horizontal>
            <Icon style={{ fontSize: "14px", color: '#00829B' }} iconName="Back" />
            <span style={{ marginLeft: '8px', color: '#00829B' }} ><a href={webURL + "/sitepages/CollabHome.aspx"} style={{ color: '#00829B', fontSize: "12px" }}>Return to home</a></span>
          </Stack>
          <div className={styles.title}>Upload VCAD File</div>
          <Stack horizontal horizontalAlign="space-between" style={{ marginBottom: '8px' }}>
            {/* <div className={styles.subTitle}>Upload an excel document</div> */}
            {/* <div className={styles.subTitle}>*Invalid file case</div> */}
          </Stack>
          <Stack horizontal horizontalAlign="start" style={{ marginLeft: 10, marginBottom: 10 }}>
            <Label style={{ width: 100, whiteSpace: 'nowrap', flexShrink: 0 }}>Select Period</Label>
            <Dropdown
              options={periodNameOption}
              styles={dropdownStyles}
              onChange={handleDropdownChange_Period}
            // required
            />
            {selectedKeyPeriod && <Label>Period Details: {selectedKeyPeriod}</Label>}
          </Stack>
          {
            uploadFile
              ? <Stack className={styles.uploadBox} verticalAlign="center" style={{ alignItems: 'flex-start' }}>
                <Stack horizontal style={{ alignItems: 'center' }}>
                  <div className={styles.subTitle}>{uploadFile.name}</div>
                  {/* <div>Parma: {String(items[3]?.__EMPTY_1)}</div> */}
                  <div onClick={() => {
                    setFile(null)
                    setData([])
                    setError('')
                    setfileWarning('')
                    setShowBtn(false)
                  }} style={{
                    marginLeft: '16px',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    borderRadius: '6px',
                    border: '1px solid #D6D3D0',
                    background: '#FFF',
                    padding: '13px',
                    cursor: 'pointer'
                  }}><Del /></div>

                </Stack>

              </Stack>
              : <Stack className={styles.uploadBox} verticalAlign="center">
                {
                  error
                    ? <div style={{ display: 'flex', alignItems: 'center' }}><Error /> <div className={styles.subTitle} style={{ color: '#E0402E', marginLeft: '8px' }}>{error}</div></div>
                    : <div className={styles.subTitle}>*Please contain VCAD Data</div>
                }
                <Upload
                  beforeUpload={() => false}
                  accept=".xlsx, .xls"
                  onChange={handleFileUpload}
                  maxCount={1}
                  showUploadList={false}
                >
                  <Button style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: '12px',
                    padding: '13px 34px',
                    fontSize: '16px',
                    borderRadius: '6px',
                    border: '1px solid #D6D3D0',
                    background: '#FFF'
                  }} icon={<FileSvg />}

                  >Select files</Button>
                </Upload>
                {
                  fileWarning && <div style={{ display: 'flex', alignItems: 'center' }}><Error /> <div className={styles.subTitle} style={{ color: '#E0402E', marginLeft: '8px' }}>{fileWarning}</div></div>
                }
              </Stack>

          }

          {
            selectedKey !== "" && showBtn && !error ? <> {
              fileWarning && <div style={{ display: 'flex', alignItems: 'center', color: 'green' }}><Error /> <div className={styles.subTitle} style={{ color: 'rgb(219 155 22)', marginLeft: '8px' }}>{fileWarning}</div></div>
            }
              <Button style={{
                width: 140, marginTop: '32px', borderRadius: '6px', color: '#fff',
                background: '#00829B'
              }} onClick={async () => {
                try {
                  doesFolderExist(selectedKey + uploadFile.name.substring(uploadFile.name.lastIndexOf('.'))).then(exsit => {
                    console.log(exsit)
                    setfileExistFlag(exsit)
                  });

                  setIsShowModal(true);
                } catch (error) {
                  console.error('Error checking if folder exists:', error);
                  setfileExistFlag(false);
                  setIsShowModal(true);
                }
              }}
              // disabled={selectedKey===""}
              >
                Upload</Button></>
              :
              <Button style={{
                width: 140, marginTop: '32px', borderRadius: '6px', color: '#fff',
                background: '#C4C4C4'
              }}>Upload</Button>
          }

        </div> : <div style={{ height: '100px', paddingTop: '64px' }}>
          <p style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '24px', fontWeight: 'bold' }}>
            <div style={{ marginRight: '10px', display: 'flex', alignItems: 'center' }}>
              <Viewhistory />
            </div>
            Submitted!
          </p>
          <p style={{ fontSize: '14px', textAlign: 'center' }}>Submitted successfully! </p>
          <Stack style={{ alignItems: 'center' }}>
            <Button style={{
              width: 80, height: 42, marginTop: '2px', borderRadius: '6px', color: '#fff',
              background: '#00829B', alignItems: 'center'
            }} onClick={() => window.location.href = webURL + "/sitepages/CollabHome.aspx"}>OK </Button></Stack>
        </div>}
      <Modal open={isShowModal} closable={false} footer={null} width={500} style={{ borderRadius: '6px', overflow: 'hidden', paddingBottom: 0 }}>
        <Stack verticalAlign="center" style={{ alignItems: 'center', paddingTop: '64px', paddingBottom: '54px' }}>
          {fileExistFlag ? <p>File already exists. Do you want to overwrite it?</p> : <p>Are you sure you want to upload this file?</p>}
          <div style={{ paddingLeft: 98, display: 'flex', alignItems: 'flex-start', flexDirection: 'column' }}>

          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '36px' }}>
            <Button onClick={() => setIsShowModal(false)} style={{ width: 120, height: 42, marginTop: '32px', borderRadius: '6px' }}>Cancel</Button>
            <Button style={{
              width: 120, height: 42, marginTop: '32px', borderRadius: '6px', color: '#fff',
              background: '#00829B'
            }} onClick={() => {
              setIsShowModal(false)
              submitFunction()
            }}>Yes</Button>
          </div>
        </Stack>
      </Modal>

    </div>
  )





})