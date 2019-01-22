import * as React from "react";
import { escape } from "@microsoft/sp-lodash-subset";
import { DisplayMode } from "@microsoft/sp-core-library";
import { ChoiceGroup, IChoiceGroupOption } from "office-ui-fabric-react/lib/ChoiceGroup";
import { PrimaryButton, DefaultButton } from "office-ui-fabric-react/lib/Button";
import { Icon } from "office-ui-fabric-react/lib/Icon";
import { autobind } from "office-ui-fabric-react/lib/Utilities";
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { clone, findIndex } from '@microsoft/sp-lodash-subset';

// Use -> PnPJS
import styles from "./Pollen.module.scss";
import MockHttpClient from "../../MockHttpClient";
import PSPClient from "../../PollenSPHttpClient";
import { IPollenProps } from "./IPollenProps";
import { IPollListItem, IPollResponse } from "../types";
import { PollenCommandBar } from "../index";

export interface IPollenState {
    initialized: boolean;
    author: string;
    currentPoll: IPollListItem;
    currentView: string;
    pollOptions: IChoiceGroupOption[];
    pollResponse: IPollResponse;
    pollResponses: IPollResponse[];
}

export default class Pollen extends React.Component<IPollenProps, IPollenState> {
    private pollListItemEntityTypeName: string = "SP.Data.PollenListItem";
    private responseListItemEntityTypeName: string = "SP.Data.PollResponsesListItem";

    constructor(props: IPollenProps) {
        super(props);

        this.state = { initialized: false, author: null, currentPoll: null, currentView: null, pollOptions: [], pollResponse: null, pollResponses: null };
    }

    public componentDidMount(): void {
        PSPClient.GetUserInfo().then(user => {            
            this.setState({author: user.Email});
            this.loadCurrentPoll();
        });
    }

    public render(): React.ReactElement<IPollenProps> {
        
        return (
            <div className={styles.pollen}>
                <div className={styles.container}>
                    {this.props.displayMode === DisplayMode.Edit && <PollenCommandBar context={this.props.context} onChangeViewMode={this.setViewMode} initialized={this.state.initialized} />}
                                    
                    <div className={styles.contentBox}>
                        <div className={styles.contentBoxHeader}>
                            <h2>
                                <Icon iconName='BarChart4' /> Poll
                            </h2>
                            <a href="mailto:{{pollSuggestEmail}}?subject=Suggest a question">suggest a question</a>
                        </div>
                        <div className={[styles.contentBoxUnscrollable, styles.poll].join(" ")}>
                            
                            {this.renderContent()}
                            
                        </div>
                    </div>
                
                </div>
            </div>
        );
    }

    // Render Methods
    private renderContent(){
        if(!this.state.initialized){
            // Install Pollen
            return this.renderNoPollList();
        } else if(!this.state.currentPoll){
            // No polls in list
            return this.renderNoPolls();
        } else {
            // Display Poll
            return this.renderPoll();
        }
    }

    private renderNoPollList(){
        return (
            <div>
                {this.props.displayMode === DisplayMode.Edit && 
                    <div>
                        <h2>Missing Pollen lists</h2>

                        <div>
                            Create the required lists using the 'Create Lists' button.
                        </div>
                    </div>
                }
            </div>
        );
    }

    private renderNoPolls() {
        return (
            <div>
                <h2>There are no polls to display.</h2>

                {this.props.displayMode === DisplayMode.Edit &&
                    <div>
                        Please create a poll using the menu.
                    </div>
                }
            </div>
        );
    }
    
    private renderPoll(){
        return (
            <div>
                <h2>{this.state.currentPoll.PollenQuestion}</h2>

                {this.state.currentView === 'poll' &&
                    <div>
                        <ChoiceGroup
                            options={this.state.pollOptions}
                            onChange={this._onSelectChoice}
                            />

                        <PrimaryButton
                            onClick={this._onSubmitPoll}
                            iconProps={{iconName: 'Send'}}
                            text="Submit"
                            />                                            
                    </div>
                }
                
                {this.state.currentView === 'responses' &&
                    <div>
                        {this.state.currentPoll.PollenChoices.length > 0 &&
                        this.state.currentPoll.PollenChoices.map((choice, index) => {
                            return (
                                <div key={index}>
                                    <div className={styles.pollAnswer}>
                                        <label>{choice}</label>
                                    </div>
                                    <div className={styles.pollResult}>
                                        <div id={`poll-result-${index}`} style={{width: this.getPercent(index) + '%'}}></div>
                                        <span className={styles.resultPercentage}>{this.getPercent(index)}%</span>
                                    </div>
                                </div>
                            );
                        })}
                    </div>
                }
            </div>
        );
    }

    private getPollOptions(poll: IPollListItem): IChoiceGroupOption[] {
        var options: IChoiceGroupOption[] = poll.PollenChoices.map((choice, index): IChoiceGroupOption => {
            return {key:(index+1).toString(), text:choice};
        });

        return options;
    }

    private getPercent(choice: number): number {
        if(this.state.pollResponses.length === 0) {
            return 0;
        } else {
            var responses = this.state.pollResponses.filter((_response: IPollResponse) => _response.PollenResponse == choice).length;
            return Math.round((responses / this.state.pollResponses.length) * 100);
        }
    }

    private getViewMode(poll: any, responses: IPollResponse[]): string {
        var subIndex = findIndex(responses, (response:IPollResponse) => {return response.Author.EMail === this.state.author;});
        
        if (this.props.scheduler && poll.PollenEndDate <= new Date().toISOString()) {
            return 'responses';
        }
        
        if(subIndex !== -1){
            return 'responses';
        } else {
            return 'poll';
        }
    }

    // Data Methods
    private loadCurrentPoll() {
        if(Environment.type === EnvironmentType.Local){
            return this._loadMockPoll();
        } else {
            return this._loadSPPoll();
        }        
    }

    private _loadMockPoll(){
        MockHttpClient.getPoll(this.props.context.pageContext.web.absoluteUrl).then((poll: IPollListItem) => {
            MockHttpClient.GetPollResponses(poll.Id).then((responses: IPollResponse[]) => {
                this.setState({
                    initialized: true,
                    currentPoll: poll,
                    currentView: this.getViewMode(poll, responses),
                    pollOptions: this.getPollOptions(poll),
                    pollResponses: responses
                });
            });
        });
    }

    private _loadSPPoll(){
        PSPClient.IsConfigured().then((isConfigured) => {
            if(isConfigured) {
                this.setState({initialized: true}, () => {
                    if(this.props.scheduler) {
                        PSPClient.GetCurrentPoll().then(this.getPollResponses);
                    } else {
                        PSPClient.GetPoll(this.props.pollQuestion).then(this.getPollResponses);
                    }
                });
            }
        });
    }

    @autobind
    private getPollResponses(poll: IPollListItem){
        return PSPClient.GetPollResponses(poll.Id).then((responses: IPollResponse[]) => {
            this.setState({
                currentPoll: poll,
                currentView: this.getViewMode(poll, responses),
                pollOptions: this.getPollOptions(poll),
                pollResponses: responses
            });
        });
    }

    private _saveMockResponse(): Promise<IPollResponse>{
        return MockHttpClient.SaveResponse(this.state.pollResponse);
    }

    private _saveSPResponse(): Promise<IPollResponse> {
        var poll = this.state.currentPoll;
        var response = this.state.pollResponse;

        return PSPClient.SaveResponse(response, poll);
    }

    private refreshPollResponses(): Promise<any>{
        if(Environment.type === EnvironmentType.Local){
            return MockHttpClient.GetPollResponses(this.state.currentPoll.Id);
        } else {
            return PSPClient.GetPollResponses(this.state.currentPoll.Id);
        }
    }

    // Event Handlers
    @autobind
    private _onSelectChoice(ev: React.FormEvent<HTMLInputElement>, option: any) {
        var response: IPollResponse = {
            Title: this.state.currentPoll.Title,
            PollenPoll: { ID: this.state.currentPoll.Id },
            PollenResponse: option.key
        };

        this.setState({pollResponse: response});
    }

    @autobind
    private _onSubmitPoll(): void{
        if(Environment.type === EnvironmentType.Local){
            this._saveMockResponse().then(() => this.refreshPollResponses().then((responses) => {
                this.setState({pollResponses: responses});
                this.setViewMode('responses');}
            ));
        } else {
            this._saveSPResponse().then(() => this.refreshPollResponses().then((responses) => {
                this.setState({pollResponses: responses});
                this.setViewMode('responses');}
            ));
        }
    }

    @autobind
    private setViewMode(view?:string): void {
        if(view){
            this.setState({currentView: view});
        } else if(this.state.currentView === 'poll'){
           this.setState({currentView: 'responses'});
        } else {
            this.setState({currentView: 'poll'});
        }
    }
}
