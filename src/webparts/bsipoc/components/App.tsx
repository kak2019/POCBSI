import * as React from "react";
import { memo, useContext, useEffect } from "react";
import AppContext from "../../../common/AppContext";
import { DefaultButton } from "@fluentui/react/lib/Button";
import { spfi } from "@pnp/sp";
import { getSP } from "../../../common/pnpjsConfig";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import { ChoiceGroup, PrimaryButton } from "office-ui-fabric-react";


import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from '@fluentui/react/lib/Dropdown';

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300 },
};

const options: IDropdownOption[] = [
  
  { key: 'NZ', text: 'NZ' },
  { key: 'banana', text: 'Banana' },
  { key: 'grape', text: 'Grape' },
  { key: 'broccoli', text: 'Broccoli' },
  { key: 'carrot', text: 'Carrot' },
  { key: 'lettuce', text: 'Lettuce' },
];



export default memo(function App() {
    const sp = spfi(getSP());
    useEffect(() => {
        // 拿主表订单
        const itemoption = sp.web.lists.getByTitle("UD BSI_PartnerConfig").renderListDataAsStream({
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
            field_11 = Sales Package
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
            console.log("resoriginal", response.Row)
            console.log("res", response.Row.filter((item)=>item.field_2==="NZ"))
            if (response.Row.length > 0) {
                //const parmaNoList = response.Row.map((item: { PARMANo: string; }) => item.PARMANo).filter(parmaNo => parmaNo !== undefined && parmaNo !== "undefined");
                // const parmaNoList = response.Row.filter(item => item.PARMANo && item.CaseID).map(item => ({
                //     PARMANo: item.PARMANo,
                //     CaseID: item.CaseID
                // }));
                
        }})
    
        
    }, [])


    // 包的单价表
    useEffect(() => {
        // 拿包的单价表
        const itemoption = sp.web.lists.getByTitle("UD BSI_PackageMaster").renderListDataAsStream({
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
            console.log("resoriginalPackage", response.Row)
            console.log("respackage", response.Row.filter((item)=>item.field_2))
            if (response.Row.length > 0) {
                //const parmaNoList = response.Row.map((item: { PARMANo: string; }) => item.PARMANo).filter(parmaNo => parmaNo !== undefined && parmaNo !== "undefined");
                // const parmaNoList = response.Row.filter(item => item.PARMANo && item.CaseID).map(item => ({
                //     PARMANo: item.PARMANo,
                //     CaseID: item.CaseID
                // }));
                
        }})
    
        
    }, [])

    // 应用单价（每人）表
    useEffect(() => {
        // 拿应用价格表
        const itemoption = sp.web.lists.getByTitle("UD BSI_AppPriceMaster").renderListDataAsStream({
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
            console.log("resoriginalAPP", response.Row)
            console.log("resAPP", response.Row.filter((item)=>item.field_2))
            if (response.Row.length > 0) {
                //const parmaNoList = response.Row.map((item: { PARMANo: string; }) => item.PARMANo).filter(parmaNo => parmaNo !== undefined && parmaNo !== "undefined");
                // const parmaNoList = response.Row.filter(item => item.PARMANo && item.CaseID).map(item => ({
                //     PARMANo: item.PARMANo,
                //     CaseID: item.CaseID
                // }));
                
        }})
    
        
    }, [])


    return (
        <>
        <div>BSI POC</div>

        <ChoiceGroup></ChoiceGroup>
        <PrimaryButton>Click</PrimaryButton>
        </>
    )









})