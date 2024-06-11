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
import { Icon, Label } from "office-ui-fabric-react";
// import { Icon as IconBase } from '@fluentui/react/lib/Icon';
import { Upload, Modal } from 'antd';
// import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { useEffect } from "react";

import { spfi } from "@pnp/sp";
import { getSP } from "../../../common/pnpjsConfig";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import styles from './UploadPage.module.scss'
import FileSvg from '../assets/file'
import Del from '../assets/delete'
import Error from '../assets/error'
import AppContext from "../../../common/AppContext";
import Viewhistory from '../assets/submit';
import HintIcon from "../assets/hinticon"
import "./App.css";
// import * as moment from "moment";
// import helpers from "../../../config/Helpers";
// 定义 Excel 文件中数据的类型
// interface IexcelData {
//     PARMANo: string;
//     CompanyName: string;
//     ASNStreet: string;
//     ASNPostCode: string;
//     ASNCountryCode: string;
//     ASNPhone: string
// }



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
function extractAddresses(jsonObject: JsonData): Address[] {
    // 提取地址信息
    const addresses = jsonObject.internationalVersion.address;
    return addresses;
}
// eslint-disable-next-line @typescript-eslint/explicit-function-return-type
const getArrayKey = (arr: Array<{ [key in string]: any }>, key: string) => {
    for (let i = 0; i < arr.length; i++) {
        if (arr[i]['UD-KMP'] === key) {
            return arr[i]
        }
    }
}


// eslint-disable-next-line @typescript-eslint/explicit-function-return-type
const getSubTableData = (arr: Array<{ [key in string]: any }>) => {
    let i = 30
    const res = []
    while (arr[i]['UD-KMP'] !== "CONSEQUENSES FOR OTHER SUPPLIERS?:") {
        // res.push(arr[i])
        res.push({ "Packaging account no": arr[i]['UD-KMP']??"", "company name": arr[i]?.__EMPTY??"", "City": arr[i]?.__EMPTY_1??"", 'Country Code': arr[i]?.__EMPTY_2 ? String(arr[i]?.__EMPTY_2) : "" })
        i++
    }
    return res
}

// 获取61-72
const getPackageData = (arr: Array<{ [key in string]: any }>) => {
    // @ts-ignore 
    const start = arr.findIndex(val => val.__rowNum__ === 59)
    // @ts-ignore eslint-disable-next-line
    const end = arr.findIndex(val => val.__rowNum__ === 72)

    return arr.slice(start + 1, end).map(val => {
        return { "Packaging": val['__EMPTY_1']??'', "Packaging Name": val.__EMPTY_2??'', "Yearly need": val.__EMPTY_3??'' }
    }).filter(val => val.Packaging !== 0 && val.Packaging !== undefined && val.Packaging !== '')
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

// 获取79-101
const getData2 = (arr: Array<{ [key in string]: any }>) => {
    // @ts-ignore 
    const start = arr.findIndex(val => val.__rowNum__ === 79)
    // @ts-ignore eslint-disable-next-line
    const end = arr.findIndex(val => val.__rowNum__ === 101)

    const table1 = arr.slice(start + 1, end).map(val => {
        return {
            "Packaging": val.__EMPTY_2??'',
            "weekly need": val.__EMPTY??'',
            "Packaging Name": val.__EMPTY_3??'',
            "Yearly need": val.__EMPTY_4??'',
        }
    }).filter(val => val.Packaging !== 0 && val.Packaging !== undefined && val.Packaging !== '')

    // const table2 = arr.slice(start + 1, end).map(val => {
    //     return { 
    //         "Packaging": val['__EMPTY_3'],
    //         "Packaging Name": val['__EMPTY_4'],
    //         "Yearly need":val['__EMPTY_5'],
    //     }
    // })
    // return [table1, table2]
    return table1

}

// eslint-disable-next-line @typescript-eslint/explicit-function-return-type


export default memo(function App() {
    const sp = spfi(getSP());
    const [items, setItems] = useState([]);
    const [data, setData] = useState({});
    const [error, setError] = useState('');
    const [showBtn, setShowBtn] = useState(false)
    const [uploadFile, setFile] = useState<any>()
    const [isShowModal, setIsShowModal] = useState(false)
    // const [apiResponse, setApiResponse] = useState<any>(null);
    const [submiting, setSubmiting] = React.useState<boolean>(false)
    const [spParmaList, setspParmaList] = React.useState([])
    const [fileWarning, setfileWarning] = React.useState("")
    const [buttonvisible, setbuttonVisible] = React.useState<boolean>(true)
    const ctx = useContext(AppContext);
    const userEmail = ctx.context?._pageContext?._user?.email;
    const webURL = ctx.context?._pageContext?._web?.absoluteUrl;
    const [isLoading, setisloading] = React.useState<boolean>(false)
    const [isPLTeam, setisPLTeam] = React.useState<boolean>(false)
  

   



    // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
    const handleFileUpload = (info: any) => {
        // 重置状态
        setItems([]);
        setData({});
        setError(null);
        setFile(null); // 确保每次上传前文件状态都被重置
        setShowBtn(false); // 根据需要可能还需要重置其他UI状态
        if (info.file) {
            const file = info.file
            if (!file) return;

            const reader = new FileReader();
            reader.onload = (e: ProgressEvent<FileReader>) => {
                const binaryStr = e.target?.result;
                if (typeof binaryStr === 'string') {
                    const workbook = XLSX.read(binaryStr, { type: 'binary' });
                    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);
                   
            };
            reader.readAsBinaryString(file);
        }
    };
   

    const submitFunction = async (): Promise<void> => {
        if (submiting) return
        setSubmiting(true)
        let index = 0
        for (let i = 0; i < items.length; i++) {
            if (items[i]['UD-KMP'] === 'CONSEQUENSES FOR OTHER SUPPLIERS?:') {
                index = i
                break
            }
        }
      

        }
        const sp = spfi(getSP());
        // let promiss
        const request = {}
        addRequest({ request }).then(async promises => {
            console.log("promiss", promises, typeof (promises));
            const responseData = (promises as Record<string, any>).data;
            const id = responseData.ID;
            console.log('ID:', id);
            // console.log(promises.indexOf('ID'))
            const folderName = id;
            await sp.web.folders.addUsingPath(`Nii Case Library/${folderName}`)
            const res = await sp.web.getFolderByServerRelativePath(`Nii Case Library/${folderName}`).files.addUsingPath(sanitize(uploadFile.name), uploadFile)
            const item = await res.file.getItem()
            // const contentTypes = await sp.web.lists.getByTitle('Nii Case Library').contentTypes.getContextInfo()
            // console.log(contentTypes)
            // .get()
            // .then(result => {
            //   const FinalFileContentTypeId = result.filter((contenType) => {
            //     return contenType.Name === CONST.FinalFileCT;
            //   })[0].StringId;
            // @ts-ignore
            sp.web.lists.getByTitle('Nii Case Library').contentTypes()
                .then(async (result: any[]) => {
                    const UploadFileContentTypeId = result.filter((contenType) => {
                        return contenType.Name === 'uploadFile';
                    })[0].StringId;
                    //@ts-ignore
                    const finalRes = await sp.web.lists.getByTitle('Nii Case Library').items.getById(item.ID).update({
                        ContentTypeId: UploadFileContentTypeId
                    })
                }).then(() =>
                    // window.location.href = webURL + "/sitepages/CollabHome.aspx"
                    setbuttonVisible(false)
                );
            //sp.web.lists.getByTitle("Nii Case Library").rootFolder.folders.add(folderName.toString());
        }).catch(err => console.log("err", err));
    }


    return (
       <div className={styles.uploadPage}>
            {/* {error} */}
            {/* <div className={styles.header}>
                <Stack horizontal>
                    <Label style={{ width: "70%", fontSize: 20 }}>Create New Case</Label> 
                    <Icon style={{ fontSize: "25px" }} iconName="HomeSolid" />
                    <Link rel="www.baidu.com" style={{ textAlign: "right" }}>GO to Home Page</Link>
                </Stack>
            </div> */}
            {
                buttonvisible ? <div className={styles.content}>
                    <Stack horizontal>
                        <Icon style={{ fontSize: "14px", color: '#00829B' }} iconName="Back" />
                        <span style={{ marginLeft: '8px', color: '#00829B' }} ><a href={webURL + "/sitepages/CollabHome.aspx"} style={{ color: '#00829B', fontSize: "12px" }}>Return to home</a></span>
                    </Stack>
                    <div className={styles.title}>Create New Case</div>
                    <Stack horizontal horizontalAlign="space-between" style={{ marginBottom: '8px' }}>
                        <div className={styles.subTitle}>Upload an excel document</div>
                        {/* <div className={styles.subTitle}>*Invalid file case</div> */}
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
                                <div style={{paddingLeft:10}}> 
                                    <Label>Parma : {String(items[3]?.__EMPTY_1)}</Label>
                                    <Label>Company : {items[4]?.__EMPTY_1&&String(items[4]?.__EMPTY_1)}</Label>
                                    </div>
                            </Stack>
                            : <Stack className={styles.uploadBox} verticalAlign="center">
                                {
                                    error
                                        ? <div style={{ display: 'flex', alignItems: 'center' }}><Error /> <div className={styles.subTitle} style={{ color: '#E0402E', marginLeft: '8px' }}>{error}</div></div>
                                        : <div className={styles.subTitle}>*Please contain supplier company name</div>
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
                                    disabled={isPLTeam}
                                    >Select files</Button>
                                </Upload>
                                {
                                    fileWarning && <div style={{ display: 'flex', alignItems: 'center' }}><Error /> <div className={styles.subTitle} style={{ color: '#E0402E', marginLeft: '8px' }}>{fileWarning}</div></div>
                                }
                            </Stack>

                    }

                    {
                        showBtn && !error ? <> {
                            fileWarning && <div style={{ display: 'flex', alignItems: 'center', color: 'green' }}><Error /> <div className={styles.subTitle} style={{ color: 'rgb(219 155 22)', marginLeft: '8px' }}>{fileWarning}</div></div>
                        }
                            <Button style={{
                                width: 140, marginTop: '32px', borderRadius: '6px', color: '#fff',
                                background: '#00829B'
                            }} onClick={() => setIsShowModal(true)}>Upload</Button></>
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
                    <p style={{ fontSize: '14px', textAlign: 'center' }}>Submitted successfully! The request will be listed in some minutes.</p>
                    <Stack style={{ alignItems: 'center' }}>
                        <Button style={{
                            width: 80, height: 42, marginTop: '2px', borderRadius: '6px', color: '#fff',
                            background: '#00829B', alignItems: 'center'
                        }} onClick={() => window.location.href = webURL + "/sitepages/CollabHome.aspx"}>OK </Button></Stack>
                </div>}
            <Modal open={isShowModal} closable={false} footer={null} width={500} style={{ borderRadius: '6px', overflow: 'hidden', paddingBottom: 0 }}>
                <Stack verticalAlign="center" style={{ alignItems: 'center', paddingTop: '64px', paddingBottom: '54px' }}>
                    <p>Are you sure you want to upload this file?</p>
                    <div style={{paddingLeft:98 ,display: 'flex', alignItems: 'flex-start', flexDirection: 'column'}}> 
                                    <Label>Parma : {String(items[3]?.__EMPTY_1)}</Label>
                                    <Label >Company : {items[4]?.__EMPTY_1&&String(items[4]?.__EMPTY_1)}</Label>
                                    </div>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '36px' }}>
                        <Button onClick={() => setIsShowModal(false)} style={{ width: 120, height: 42, marginTop: '32px', borderRadius: '6px' }}>Cancel</Button>
                        <Button style={{
                            width: 120, height: 42, marginTop: '32px', borderRadius: '6px', color: '#fff',
                            background: '#00829B'
                        }} onClick={() => {
                            setIsShowModal(false)
                           
                        }}>Yes</Button>
                    </div>
                </Stack>
            </Modal>

        </div>
    )





})