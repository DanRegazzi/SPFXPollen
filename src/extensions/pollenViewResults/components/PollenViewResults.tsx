import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import * as React from 'react';

import styles from './PollenViewResults.module.scss';
import IPollResponse from './IPollResponse';
import PSPClient from '../../../webparts/pollen/PollenSPHttpClient';

export interface IPollenViewResultsProps {
    id: string;
    choices: string[];
}

export interface IPollenViewResultsState {
    pollResponses: IPollResponse[];
}

const LOG_SOURCE: string = 'PollenViewResults';
const RESULT_LIST: string= '';

export default class PollenViewResults extends React.Component<IPollenViewResultsProps, IPollenViewResultsState> {
    constructor(props: IPollenViewResultsProps){
        super(props);

        this.state = {pollResponses: null};
    }
    
    @override
    public componentDidMount(): void {
        Log.info(LOG_SOURCE, 'React Element: PollenViewResults mounted');
        PSPClient.GetPollResponses(this.props.id).then(responses => this.setState({pollResponses:responses}));
    }

    @override
    public componentWillUnmount(): void {
        Log.info(LOG_SOURCE, 'React Element: PollenViewResults unmounted');
    }

    @override
    public render(): React.ReactElement<{}> {
        return (
        <div className={styles.PollenViewResults}>
            <div className={styles.cell}>
                {this.state.pollResponses && this.props.choices.length > 0 &&
                this.props.choices.map((choice, index) => {
                    return (
                        <div key={index}>
                            <div>
                                <label>{choice} ({this.getPercent(index)}%)</label>
                            </div>
                            <div className={styles.pollResult}>
                                <div id={`poll-result-${index}`} style={{width: this.getPercent(index) + '%'}}></div>
                            </div>
                        </div>
                    );
                })}
            </div>
        </div>
        );
    }

    //@autobind
    private getPercent(choice: number): number {
        console.log(choice);
        if(this.state.pollResponses.length === 0) {
            return 0;
        } else {
            var responses = this.state.pollResponses.filter((_response: IPollResponse) => _response.PollenResponse == choice).length;
            return Math.round((responses / this.state.pollResponses.length) * 100);
        }
    }
}
