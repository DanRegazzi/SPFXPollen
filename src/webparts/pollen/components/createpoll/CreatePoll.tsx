import * as React from 'react';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton, CommandButton, IconButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DatePicker, DayOfWeek, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { clone } from '@microsoft/sp-lodash-subset';

import * as strings from 'PollenWebPartStrings';
import styles from './CreatePoll.module.scss';
import { IPollListItem } from '../types';
import ICreatePollProps from './ICreatePollProps';
import MockHttpClient from '../../MockHttpClient';
import PollenSPHttpClient from '../../PollenSPHttpClient';

export interface ICreatePollState {
    isDialogOpen: boolean;
    poll: IPollListItem;
}

export default class CreatePoll extends React.Component<ICreatePollProps, ICreatePollState> {
    
    private newPoll: IPollListItem = {
        Title: "",
        PollenQuestion: "",
        PollenChoices: ["", ""],
        PollenStartDate: new Date(),
        PollenEndDate: new Date()
    };

    constructor(props: ICreatePollProps) {
        super(props);

        this.state = {
            isDialogOpen: false,
            poll: clone(this.newPoll)
        };
    }

    public render() {

        return (
            <div>
                <CommandButton
                    description='Create a new Poll'
                    onClick={this._openDialog}
                    iconProps={ { iconName: 'Add' } }
                    text='Create Poll' />

                <Dialog
                    isOpen={this.state.isDialogOpen}
                    onDismiss={this._closeDialog}
                    containerClassName={styles["modal-container"]}
                    type={DialogType.largeHeader}
                    title={'Create new poll'}
                    subText={'Create a new poll, set answer options, and configure start and end dates when the poll will be available.'}>                    
                    
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
                                    strings={strings.DayPickerStrings}
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
                        <DefaultButton onClick={ this._closeDialog } text='Cancel' />
                    </DialogFooter>
                </Dialog>
            </div>
        );
    }

    @autobind
    private _openDialog(): void{
        this.setState({isDialogOpen: true, poll: clone(this.newPoll)});
    }

    @autobind
    private _closeDialog(): void{
        this.setState({isDialogOpen: false, poll: this.state.poll});
    }

    @autobind
    private _onChange(key:string, value: any): void{
        var state = clone(this.state);
        
        state.poll[key] = value;

        if(key === "PollQuestion") {
            state.poll["Title"] = value;
        }

        this.setState(state);
    }

    @autobind
    private _onChangeChoice(id:number, value:any): void{
        var state = clone(this.state);
        state.poll.PollenChoices[id] = value;

        this.setState(state);
    }

    @autobind
    private _onAddChoice(): void {
        let state = clone(this.state);
        state.poll.PollenChoices.push("");

        this.setState(state);
    }

    @autobind
    private _onRemoveChoice(index: number, value: any): void {
        let state = clone(this.state);

        state.poll.PollenChoices.splice(index, 1);
        
        // remove poll at index id
        this.setState(state);
    }

    @autobind
    private _createPoll(): void{
        if(Environment.type === EnvironmentType.Local){
            this._createMockPoll(this.state.poll);
        } else {
            this._createSPPoll(this.state.poll);
        }

        this.setState({isDialogOpen: false, poll: this.state.poll});
    }

    private _createMockPoll(poll: IPollListItem): Promise<any> {
        return MockHttpClient.create(poll).then((pollId) => {
            this._onChange('Id', pollId);
        });
    }

    private _createSPPoll(poll: IPollListItem): Promise<any> {
        return PollenSPHttpClient.CreatePoll(poll);
    }
}