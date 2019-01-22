import * as React from 'react';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { CommandButton } from 'office-ui-fabric-react/lib/Button';
import { SPComponentLoader } from '@microsoft/sp-loader';

import IConfigureListsProps from './IConfigureListsProps';
import PSPClient from '../../PollenSPHttpClient';

export interface ICreatePollState {
    isLoadingScripts: boolean;
}

export default class CreatePoll extends React.Component<IConfigureListsProps, ICreatePollState> {    
    constructor(props: IConfigureListsProps) {
        super(props);

        this.state = {isLoadingScripts: true};
    }   

    public componentDidMount() {
        SPComponentLoader.loadScript('/_layouts/15/init.js', {
            globalExportsName: '$_global_init'
          })
          .then((): Promise<{}> => {
            return SPComponentLoader.loadScript('/_layouts/15/MicrosoftAjax.js', {
              globalExportsName: 'Sys'
            });
          })
          .then((): Promise<{}> => {
            return SPComponentLoader.loadScript('/_layouts/15/SP.Runtime.js', {
              globalExportsName: 'SP'
            });
          })
          .then((): Promise<{}> => {
            return SPComponentLoader.loadScript('/_layouts/15/SP.js', {
              globalExportsName: 'SP'
            });
          })
          .then((): void => {
            this.setState({isLoadingScripts: false});
          });
    }

    public render() {

        return (
            <div>
                <CommandButton
                    description='Configure Lists'
                    onClick={this._configureLists}
                    iconProps={ { iconName: 'Add' } }
                    disabled={this.state.isLoadingScripts}
                    text='Configure Lists' />                
            </div>
        );
    }

    @autobind
    private _configureLists(): void{
        PSPClient.ConfigureLists();
    }
}