import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneCheckbox, PropertyPaneLabel, PropertyPaneHorizontalRule } from "@microsoft/sp-webpart-base";
import { IDropdownOption } from "office-ui-fabric-react/lib/components/Dropdown";
import { update, get, each } from "@microsoft/sp-lodash-subset";
import { Environment, EnvironmentType, DisplayMode } from "@microsoft/sp-core-library";
import { autobind } from "office-ui-fabric-react/lib/Utilities";

import * as strings from "PollenWebPartStrings";
import { PropertyPaneAsyncDropdown } from "../../controls/PropertyPaneAsyncDropdown/PropertyPaneAsyncDropdown";
import { Pollen } from "./components";
import { IPollenProps } from "./components/pollen/IPollenProps";
import { IPollListItem } from "./components/types";
import PollenSPHttpClient from "./PollenSPHttpClient";
import { PnPClientStorage } from "@pnp/common";
import { sp } from "@pnp/sp";

export interface IPollenWebPartProps {
    pollQuestion: number;
    poll: IPollListItem;
    scheduler: boolean;
    displayMode: number;
    contextUrl: string;
}

export default class PollenWebPart extends BaseClientSideWebPart<IPollenWebPartProps> {
    public onInit(): Promise<void> {
        return super.onInit().then(_ => {
      
            sp.setup({
                spfxContext: this.context
            });
          
        });
    }

    public render(): void {
        const element: React.ReactElement<IPollenProps> = React.createElement(Pollen, {
            pollQuestion: this.properties.pollQuestion,
            scheduler: this.properties.scheduler,
            displayMode: this.getDisplayMode(),
            context: this.context
        });

        ReactDom.render(element, this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse("1.0");
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneCheckbox("scheduler", {                                    
                                    text: strings.SchedulerFieldLabel
                                }),
                                PropertyPaneHorizontalRule(),
                                new PropertyPaneAsyncDropdown("pollQuestion", {
                                    label: strings.PollQuestionFieldLabel,
                                    loadOptions: this.loadPolls,
                                    onPropertyChange: this.onPollChange,
                                    selectedKey: this.properties.pollQuestion,
                                    disabled: this.properties.scheduler
                                })                                
                            ]
                        }
                    ]
                }
            ]
        };
    }

    private loadPolls(): Promise<IDropdownOption[]> {
        return PollenSPHttpClient.GetPolls().then((polls => {
            var options: IDropdownOption[] = [];

            each(polls, (poll: IPollListItem) => {
                options.push({key: poll.Id, text: poll.Title});
            });

            return options;
        }));
    }

    @autobind
    private onPollChange(propertyPath: string, newValue: any): void {
        const oldValue: any = get(this.properties, propertyPath);
        // store new value in web part properties
        update(this.properties, propertyPath, (): any => {
            console.log(newValue);
            return newValue;
        });
        // refresh webPart
        this.render();
    }

    private getDisplayMode() {
        //Detect display mode on classic and modern pages pages
        if (Environment.type == EnvironmentType.ClassicSharePoint) {
            let isInEditMode: boolean;
            let interval: any;
            interval = setInterval(() => {
                if (typeof (<any>window).SP.Ribbon !== "undefined") {
                    isInEditMode = (<any>window).SP.Ribbon.PageState.Handlers.isInEditMode();
                    if (isInEditMode) {
                        //Classic SharePoint in Edit Mode
                        clearInterval(interval);
                        return DisplayMode.Edit;
                    } else {
                        //Classic SharePoint in Read Mode
                        clearInterval(interval);
                        return DisplayMode.Read;
                    }
                }
            }, 100);
        } else if (Environment.type == EnvironmentType.SharePoint || Environment.type === EnvironmentType.Local) {
            if (this.displayMode == DisplayMode.Edit) {
                //Modern SharePoint in Edit Mode'
                return DisplayMode.Edit;
            } else if (this.displayMode == DisplayMode.Read) {
                //Modern SharePoint in Read Mode
                return DisplayMode.Read;
            }
        }
    }
}
