import * as React from "react";
import { memo, useContext, useEffect } from "react";
import AppContext from "../../../common/AppContext";
import { DefaultButton } from "@fluentui/react/lib/Button";
import { spfi } from "@pnp/sp";
import { getSP } from "../../../common/pnpjsConfig";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import { PrimaryButton } from "office-ui-fabric-react";






export default memo(function App() {
    const sp = spfi(getSP());
    useEffect(() => {

        const itemoption = sp.web.lists.getByTitle("UD BSI_PartnerConfig").renderListDataAsStream({
            /*title ===Hub
            field_1 = PartnerType
            field_2 = country
            field_3 = Partner Name
            */

            ViewXml: `<View>
                          <ViewFields>
                            <FieldRef Name="Title"/>
                            <FieldRef Name="field_1"/>
                            <FieldRef Name="field_2"/>
                            <FieldRef Name="field_3"/>
                            
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


    return (
        <>
        <div>BSI POC</div>


        <PrimaryButton>Click</PrimaryButton>
        </>
    )









})