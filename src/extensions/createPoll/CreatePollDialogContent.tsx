import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { 
    autobind, 
    PrimaryButton, 
    DefaultButton, 
    IconButton, 
    DialogFooter, 
    DialogContent, 
    DialogType,
    TextField,
    DatePicker,
    DayOfWeek } from 'office-ui-fabric-react';
import { clone } from '@microsoft/sp-lodash-subset';

import * as strings from 'CreatePollCommandSetStrings';
import styles from './CreatePollDialogContent.module.scss';
import {IPollListItem, IPollResponse} from '../../webparts/pollen/components/types';
import PollenSPHttpClient from '../../webparts/pollen/PollenSPHttpClient';

export interface ICreatePollDialogContentProps {
    close: () => void;
    submit: () => void;
}

export interface ICreatePollDialogContentState {
    poll: IPollListItem;
}

export default class CreatePollDialogContent extends React.Component<ICreatePollDialogContentProps, ICreatePollDialogContentState> {
    constructor(props){
        console.log('constructor');
        super(props);

        let newPoll = {
            Title: "",
            PollenChoices: ["",""],
            PollenStartDate: null,
            PollenEndDate: null
        } as IPollListItem;
        
        this.state = { poll: newPoll };
        console.log('state initted');
    }

    public render(): JSX.Element {
        return (
            <DialogContent
                title="Create Poll"
                subText="Create a new poll, set answer options, and configure start and end dates for the poll."
                onDismiss={this.props.close}
                showCloseButton={true}
                type={DialogType.largeHeader}>
                
                <div className="ms-Grid">
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-sm12">
                            <TextField
                                label='Poll Question'
                                multiline
                                rows={4}
                                //errorMessage='Poll question is required.'
                                value={this.state.poll.PollenQuestion}
                                onChanged={ e => this._onChange('PollenQuestion', e)} />
                        </div>                            
                    </div>
                                        
                    {this.state.poll.PollenChoices.map( (choice, index) => {
                        return (
                            <div className="ms-Grid-row">                                          
                                <div className="ms-Grid-col ms-sm10">
                                    <TextField key={index+1}
                                        label={`Choice ${index+1}`}
                                        value={choice}
                                        validateOnLoad={false}
                                        validateOnFocusOut={true}
                                        onChanged={ value => this._onChangeChoice(index, value)} />
                                </div>
                                <div className="ms-Grid-col ms-sm2">
                                    {index >= 2 &&
                                        <IconButton
                                            className={styles.removeChoice}
                                            description="Remove Choice"
                                            onClick={e => this._onRemoveChoice(index, e)}
                                            iconProps={{iconName: "ErrorBadge"}} />
                                    }                                        
                                </div>
                            </div>
                        );
                    })}
                    
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-sm12">
                            <DialogFooter>
                                <PrimaryButton 
                                    description="Add Choice"
                                    onClick={this._onAddChoice}
                                    iconProps={{iconName: "Add"}}
                                    text="Add Choice"
                                    />                                        
                            </DialogFooter>                                
                        </div>
                    </div>

                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-sm12">
                            <DatePicker
                                label='Start Date'
                                firstDayOfWeek={DayOfWeek.Sunday}
                                strings={strings.DayPickerStrings.DayPickerStrings}
                                placeholder='Select a start date' 
                                value={this.state.poll.PollenStartDate}
                                onSelectDate={ e => this._onChange('PollenStartDate', e)}
                                />
                            <DatePicker
                                label='End Date'
                                firstDayOfWeek={DayOfWeek.Sunday}
                                strings={strings.DayPickerStrings}
                                placeholder='Select an end date' 
                                value={this.state.poll.PollenEndDate}
                                onSelectDate={ e => this._onChange('PollenEndDate', e)}
                                />
                        </div>
                    </div>
                </div>

                <DialogFooter>
                    <PrimaryButton onClick={ this._createPoll } text='Save' />
                    <DefaultButton onClick={ this.props.close } text='Cancel' />
                </DialogFooter>                
            </DialogContent>
        );            
    }

    @autobind
    private _onChange(key:string, value: any): void{
        let poll = clone(this.state.poll);
        
        poll[key] = value;

        if(key === "PollQuestion") {
            poll["Title"] = value;
        }

        this.setState({poll: poll});
    }

    @autobind
    private _onChangeChoice(id:number, value:any): void{
        let poll = clone(this.state.poll);
        poll.PollenChoices[id] = value;

        this.setState({poll: poll});
    }

    @autobind
    private _onAddChoice(): void {
        let poll = clone(this.state.poll);
        poll.PollenChoices.push("");

        this.setState({poll: poll});
    }

    @autobind
    private _onRemoveChoice(index: number, value: any): void {
        let poll = clone(this.state.poll);

        poll.PollenChoices.splice(index, 1);
        
        // remove poll at index id
        this.setState({poll: poll});
    }

    @autobind
    private _createPoll(): void{
        PollenSPHttpClient.CreatePoll(this.state.poll).then(poll => {
            this.props.submit();    
        });
    }
}