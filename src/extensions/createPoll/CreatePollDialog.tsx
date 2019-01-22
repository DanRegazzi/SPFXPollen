import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { autobind } from 'office-ui-fabric-react';

import CreatePollDialogContent from './CreatePollDialogContent';

export default class CreatePollDialog extends BaseDialog {
    public render(): void {
        ReactDOM.render(<CreatePollDialogContent
            close={this.close}
            submit={this._onSubmit}/>,
            this.domElement
        );
    }

    public getConfig(): IDialogConfiguration {
        return {
            isBlocking: false
        };
    }

    @autobind
    private _onSubmit(): void {
        this.close();
        location.reload();
    }
}