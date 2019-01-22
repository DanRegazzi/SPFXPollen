import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';

export interface IAsyncDropdownProps {
    label: string;
    loadOptions: () => Promise<IDropdownOption[]>;
    onChanged: (options: IDropdownOption, index?: number) => void;
    selectedKey: string | number;
    disabled: boolean;
    stateKey: string;
}