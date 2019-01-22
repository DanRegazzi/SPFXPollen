import * as React from 'react';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import { CommandButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';

import { IPollenCommandBarProps } from './IPollenCommandBarProps';
import { ConfigureLists, CreatePoll } from '../index';

export default class PollenCommandBar extends React.Component<IPollenCommandBarProps, any> {
    public render() {
        if(this.props.initialized){
            return (
                <CommandBar items={[
                    {
                        key: 'new',
                        name: 'Add Poll',
                        onRender: this._onRenderCreatePollButton,
                        className: 'ms-CommandBarItem'
                    },
                    {
                        key: 'view',
                        name: 'Switch View',
                        onRender: this._onRenderSwitchViewButton,
                        className: 'ms-CommandBarItem'
                    }
                ]} />
            );
        } else {
            return (
                <CommandBar items={[
                    {
                        key: 'configure',
                        name: 'Configure Lists',
                        onRender: this._onRenderConfigureListsButton,
                        className: 'ms-CommandBarItem'
                    }
                ]} />
            );
        }
        
    }

    @autobind
    private _onRenderCreatePollButton(){
        return (
            <CreatePoll context={this.props.context}/>
        );
    }

    @autobind
    private _onRenderSwitchViewButton(){
        return (
            <CommandButton
                description='Switch the poll view'
                onClick={this._switchView}
                iconProps={ { iconName: 'Switch' } }
                text='Switch View' />
        );
    }

    @autobind
    private _onRenderConfigureListsButton() {
        return (
            <ConfigureLists context={this.props.context} />
        );
    }

    @autobind
    private _switchView(){
        this.props.onChangeViewMode();
    }
}