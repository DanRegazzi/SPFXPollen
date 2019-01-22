import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import {
    BaseFieldCustomizer,
    IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';

import * as strings from 'PollenViewResultsFieldCustomizerStrings';
import PollenViewResults, { IPollenViewResultsProps } from './components/PollenViewResults';

/**
 * If your field customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IPollenViewResultsFieldCustomizerProperties { }

const LOG_SOURCE: string = 'PollenViewResultsFieldCustomizer';

export default class PollenViewResultsFieldCustomizer
    extends BaseFieldCustomizer<IPollenViewResultsFieldCustomizerProperties> {

    @override
    public onInit(): Promise<void> {
        // Add your custom initialization to this method.  The framework will wait
        // for the returned promise to resolve before firing any BaseFieldCustomizer events.
        Log.info(LOG_SOURCE, 'Activated PollenViewResultsFieldCustomizer with properties:');
        Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
        Log.info(LOG_SOURCE, `The following string should be equal: "PollenViewResultsFieldCustomizer" and "${strings.Title}"`);
        return Promise.resolve();
    }

    @override
    public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
        // Use this method to perform your custom cell rendering.
        const choices: string[] = JSON.parse(event.fieldValue);
        const id: string = event.listItem.getValueByName('ID');

        const pollenViewResults: React.ReactElement<{}> =
        React.createElement(PollenViewResults, { choices, id } as IPollenViewResultsProps);

        ReactDOM.render(pollenViewResults, event.domElement);
    }

    @override
    public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
        // This method should be used to free any resources that were allocated during rendering.
        // For example, if your onRenderCell() called ReactDOM.render(), then you should
        // call ReactDOM.unmountComponentAtNode() here.
        ReactDOM.unmountComponentAtNode(event.domElement);
        super.onDisposeCell(event);
    }
}
