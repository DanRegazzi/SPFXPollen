import { IPollListItem, IPollResponse } from './components/types';
import { findIndex } from '@microsoft/sp-lodash-subset';

export default class MockHttpClient {
    public static getPoll(restUrl: string, options?: any): Promise<IPollListItem> {
        var polls = JSON.parse(localStorage.getItem('polls'));

        return new Promise<IPollListItem>((resolve) => {
            resolve(polls[polls.length-1]);
        });
    }

    public static get(restUrl: string, options?: any): Promise<IPollListItem[]> {
        var polls = JSON.parse(localStorage.getItem('polls'));
        
        if(!polls){
            polls = [];
        } 

        return new Promise<IPollListItem[]>((resolve) => {
            resolve(polls);
        });
    }

    public static create(pollListItem: IPollListItem): Promise<any> {
        var polls = JSON.parse(localStorage.getItem('polls'));
        
        if(!polls){
            polls = [];
            pollListItem.Id = polls.length;
            polls.push(pollListItem);
        } else {
            var pollIndex = findIndex(polls, (poll:IPollListItem)=>{return poll.Id === pollListItem.Id;});
            if(pollIndex === -1) {
                pollListItem.Id = polls.length;
                polls.push(pollListItem);
            } else {
                polls[pollIndex] = pollListItem;
            }
        }
        
        localStorage.setItem('polls', JSON.stringify(polls));

        return new Promise<any>((resolve) => {
            resolve(pollListItem.Id);
        });
    }

    public static SaveResponse(response: IPollResponse): Promise<IPollResponse> {
        var responses = JSON.parse(localStorage.getItem('pollResponses'));
        
        if(!responses){
            responses = [];            
        }
        
        response.Id = responses.length;
            responses.push(response);

        localStorage.setItem('pollResponses', JSON.stringify(responses));

        return new Promise<IPollResponse>((resolve) => {
            resolve(response);
        });
    }

    public static GetPollResponses(pollId: string): Promise<IPollResponse[]>{
        var responses = JSON.parse(localStorage.getItem('pollResponses'));
        
        if(!responses){
            responses = [];            
        }

        var pollResponses = responses.filter((response) => response.Poll === pollId);

        return new Promise<IPollResponse[]>((resolve) => {
            resolve(pollResponses);
        });
    }
}