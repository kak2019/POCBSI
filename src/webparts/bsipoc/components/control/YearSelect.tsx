import * as React from 'react';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';

interface YearPickerProps {
    startYear: number;
    endYear: number;
    onYearChange: (year: number) => void;  // 添加一个新的prop
}

const YearPicker: React.FC<YearPickerProps> = ({ startYear, endYear ,onYearChange}) => {
    const years: IDropdownOption[] = Array.from({ length: endYear - startYear + 1 }, (_, index) => {
        const year = startYear + index;
        return { key: year, text: year.toString() };
    });
    const onChange = (event: React.FormEvent<HTMLDivElement>, item?: IDropdownOption): void => {
        if (item) {
            onYearChange(parseInt(item.text));  // 使用onYearChange函数更新父组件的状态
        }
    };
    return (
        <Dropdown
            placeholder="Select Year"
            label="Year"
            options={years}
            styles={{ dropdown: { width: 200 } }}
            onChange= {onChange}
        />
    );
};

export default YearPicker;
