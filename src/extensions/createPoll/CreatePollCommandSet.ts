import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
    BaseListViewCommandSet,
    Command,
    IListViewCommandSetListViewUpdatedParameters,
    IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'CreatePollCommandSetStrings';
import CreatePollDialog from './CreatePollDialog';

const CREATE_POLL = 'CREATE_POLL';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICreatePollCommandSetProperties {}

const LOG_SOURCE: string = 'CreatePollCommandSet';

export default class CreatePollCommandSet extends BaseListViewCommandSet<ICreatePollCommandSetProperties> {

@override
public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized CreatePollCommandSet');
    return Promise.resolve();
}

@override
public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    console.log('list view updated');
}

@override
public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
        case CREATE_POLL:
            const dialog: CreatePollDialog = new CreatePollDialog();
            dialog.show(); 
            break;
        default:
            throw new Error('Unknown command');
        }
    }
}
