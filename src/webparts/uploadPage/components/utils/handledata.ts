interface DataRecord {
    Period?: string;
    DealerID: number;
    Application: string;
    MachineNo: string;
    Description?: string;
    AmountInUSD: number;
    Market?: string;
    "S410 W VOCOM"?: number;
    "S410 W/O VOCOM"?: number;
    "V110 W VOCOM"?: number;
    "V110 W/O VOCOM"?: number;
    "HWI"?:number;
    // PartnerID?: string;
}
const rules = [
    { key: 'HWI', condition: (record: DataRecord) => record.Description === 'HWI' },
    { key: 'S410 W VOCOM', condition: (record: DataRecord) => record.MachineNo.includes('S410') && record.MachineNo.includes('VOCOM') },
    { key: 'S410 W/O VOCOM', condition: (record: DataRecord) => record.MachineNo.includes('S410') && !record.MachineNo.includes('VOCOM') },
    { key: 'V110 W VOCOM', condition: (record: DataRecord) => record.MachineNo.includes('V110') && record.MachineNo.includes('VOCOM') },
    { key: 'V110 W/O VOCOM', condition: (record: DataRecord) => record.MachineNo.includes('V110') && !record.MachineNo.includes('VOCOM') },
];

// eslint-disable-next-line @typescript-eslint/explicit-function-return-type
const countByRules = (data: DataRecord[],selectkey:string) => {
    const groupedData: { [key: string]: DataRecord[] } = {};
    data.forEach(record => {
        if (!groupedData[record.DealerID]) {
            groupedData[record.DealerID] = [];
        }
        groupedData[record.DealerID].push(record);
    });

    const results = Object.keys(groupedData).map(dealerID => {
        const records = groupedData[dealerID];
        const counts: { [key: string]: number } = {};

        rules.forEach(rule => {
            counts[rule.key] = records.filter(rule.condition).length;
        });

        const market = records[0].Market;

        return {
            Market: market,
            PartnerID: dealerID,
            ...counts,
            Period: selectkey
        };
    });

    return results;
};


export {countByRules}