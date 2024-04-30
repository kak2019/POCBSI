/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import { memo, useEffect } from "react";
import { spfi } from "@pnp/sp";
import { getSP } from "../../../common/pnpjsConfig";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import { PrimaryButton, Stack } from "office-ui-fabric-react";
import * as XLSX from 'xlsx';
import { Dropdown, IDropdownStyles, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import ReadExcelFromSP from "./ReadExcelFromSp";
import YearPicker from "./control/YearSelect";
import createrFolder from "./control/CreateFolder"
// import XlxsExcelFromSP from "./xlxsexcel"


const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300 },
};

const options: IDropdownOption[] = [

  { key: 'Q1', text: 'Q1' },
  { key: 'Q2', text: 'Q2' },
  { key: 'Q3', text: 'Q3' },
  { key: 'Q4', text: 'Q4' },
];


export default memo(function App() {
  const sp = spfi(getSP());



  const [excel, setExcel] = React.useState([]);
  const [allCountry, setAllCountry] = React.useState([])
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
        Country: val.field_2,
        Dealer: val.field_3,
        "Dealer ID": val.field_4,
        "Dealer Type": val.field_6,
        "Dealer Category": val.field_8,
        "Basic Package": val.field_9,
        "Sales Package": val.field_10,
        "CPQ": val.field_11,
        "UD CM": val.field_12,
        "Argus 365": val.field_13,
        "UDCP": val.field_14,
        "SeMA": val.field_15,
        "LDS": val.field_16,
        "LSS": val.field_17,
        "Pardot": val.field_18
      }

      const keys = ['CPQ', 'UD CM', 'Argus 365', 'UDCP', 'SeMA', 'LDS', 'LSS', 'Pardot']
      obj['Total(Per Month)'] = keys.reduce((sum, key) => sum + (priceMap[key] || 0) * obj[key], 0)
      obj['Total(Per Month)'] += (priceMap['Basic Package;' + val.field_8] || 0) * obj['Basic Package']
      obj['Total(Per Month)'] += (priceMap['Sales Package;' + (val.field_8 || 'NA')] || 0) * obj['Sales Package']
      return obj
    })
  }

  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  const initData = async () => {
    // 拿应用价格表
    const appObj = await sp.web.lists.getByTitle("UD BSI_AppPriceMaster").renderListDataAsStream({
      /* 字段关系如下
      Title ===Application Name
      field_1 = Price Type
      field_2 = Price (USD)
      
      */

      ViewXml: `<View>
                      <ViewFields>
                        <FieldRef Name="Title"/>
                        <FieldRef Name="field_1"/>
                        <FieldRef Name="field_2"/>
                      </ViewFields>
                   
                    </View>`,
      // <RowLimit>400</RowLimit>
    }).then((response) => {
      console.log("应用单价/人", response.Row)
      // console.log("resAPP", response.Row.filter((item)=>item.field_2))
      if (response.Row.length > 0) {
        const resObj: any = {}
        response.Row.forEach(val => {
          resObj[val.Title] = val.field_2 * 1
        })
        return resObj
      }

      return {}
    })


    // 拿包的单价表
    const packageObj = await sp.web.lists.getByTitle("UD BSI_PackageMaster").renderListDataAsStream({
      /* 字段关系如下
      Title ===Package Name
      field_1 = PartnerType
      field_2 = Dealer Category
      field_3 = Monthly Price (USD)
      Comment = Comment
      */

      ViewXml: `<View>
                      <ViewFields>
                        <FieldRef Name="Title"/>
                        <FieldRef Name="field_1"/>
                        <FieldRef Name="field_2"/>
                        <FieldRef Name="field_3"/>
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
          resObj[`${val.Title};${val.field_2}`] = val.field_3 * 1
        })
        return resObj
      }

      return {}
    })

    // 拿period 周期对应关系
    const period = await sp.web.lists.getByTitle("BSI_Period").renderListDataAsStream({
      /* 字段关系如下
      Title ===Package Name
      field_1 = PartnerType
      field_2 = Dealer Category
      field_3 = Monthly Price (USD)
      Comment = Comment
      */

      ViewXml: `<View>
                      <ViewFields>
                        <FieldRef Name="Title"/>
                        <FieldRef Name="Month"/>
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
    const order = await sp.web.lists.getByTitle("UD BSI_PartnerConfig").renderListDataAsStream({
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
                        <FieldRef Name="Title"/>
                        <FieldRef Name="field_1"/>
                        <FieldRef Name="field_2"/>
                        <FieldRef Name="field_3"/>
                        <FieldRef Name="field_4"/>
                        <FieldRef Name="field_5"/>
                        <FieldRef Name="field_6"/>
                        <FieldRef Name="field_7"/>
                        <FieldRef Name="field_8"/>
                        <FieldRef Name="field_9"/>
                        <FieldRef Name="field_10"/>
                        <FieldRef Name="field_11"/>
                        <FieldRef Name="field_12"/>
                        <FieldRef Name="field_13"/>
                        <FieldRef Name="field_14"/>
                        <FieldRef Name="field_15"/>
                        <FieldRef Name="field_16"/>
                        <FieldRef Name="field_17"/>
                        <FieldRef Name="field_18"/>
                        <FieldRef Name="field_19"/>
                        
                      </ViewFields>
                   
                    </View>`,
      // <RowLimit>400</RowLimit>
    }).then((response) => {
      console.log("主表订单", response.Row)
      // console.log("res", response.Row.filter((item)=>item.field_2==="NZ"))
      if (response.Row.length > 0) {
        const uniqueList = Array.from(new Set(response.Row.map(item => item.field_2)))
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

    const finalExcelData = calcToExcel(order, price)
    console.log(finalExcelData)
    setExcel(finalExcelData)
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

    // 遍历所有国家
    allCountry.forEach(country => {
      // 筛选出该国家的订单
      const countryOrders = excel.filter(order => order.Country === country);
      console.log("countryOrders", countryOrders);
      if (countryOrders.length === 0) {
        console.log(`No data for ${country}`);
        return;
      }

      // 创建工作表
      const worksheet = XLSX.utils.json_to_sheet(countryOrders);
      ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1', 'J1', 'K1', 'L1', 'N1', 'M1', 'O1', 'P1'].forEach(key => {
        if (worksheet[key]) {
          worksheet[key].s = {
            fill: {
              fgColor: { rgb: "add8e6" }
            }
          };
        }
      });

      // 创建工作簿
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Package Details");

      // 将工作簿转换为Blob
      const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbout], { type: "application/octet-stream" });

      // 添加上传任务到数组
      const uploadPromise = uploadFileToSP(
        "/sites/proj-testspfeatures/Shared Documents/Cost Summary/2024Q1",
        `UD ${country} 2024Q1.xlsx`,
        blob
      );
      uploadPromises.push(uploadPromise);
    });

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
        await createrFolder("/sites/proj-testspfeatures/Shared Documents", selectedYear+selectedKey);
        alert("文件夹创建成功!");
    } catch (error) {
        alert("创建文件夹失败: " + error.message);
    }
};

  return (
    <>
      <h2>BSI POC</h2>
      <Stack horizontal>

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
        <PrimaryButton style={{ marginTop: 10 }} disabled={excel.length === 0 || selectedKey === ''} onClick={handleExport}>generate excel file</PrimaryButton>
      </Stack>
<Stack>
      <ReadExcelFromSP />
      <div>分割</div>
      <PrimaryButton onClick={handleCreateFolder}/>
      </Stack>
      {/* <XlxsExcelFromSP/> 这个组件 无法保留格式*/}

    </>
  )









})