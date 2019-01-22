import { IWebPartContext } from "@microsoft/sp-webpart-base";

export interface IPollenCommandBarProps {
    context: IWebPartContext;
    initialized: boolean;
    onChangeViewMode: () => any;
}