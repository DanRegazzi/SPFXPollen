import { sp, ItemAddResult } from '@pnp/sp';
import { IWebPartContext } from '@microsoft/sp-webpart-base';

import { IPollListItem, IPollResponse } from './components/types';
import Provisioner from './provision';

export const POLL_LIST = 'Pollen Polls';
export const POLL_RESPONSE_LIST = 'Pollen Responses';

export default class PollenSPHttpClient {

    public static IsConfigured(): Promise<boolean>{
        return sp.web.lists.getByTitle(POLL_LIST).get().then((pollList) => {
            return sp.web.lists.getByTitle(POLL_RESPONSE_LIST).get().then((pollResponseList) => {
                return true;
            }, (error) => {
                return false;
            });            
        }, (error) => {
            return false;
        });
    }

    public static GetUserInfo(): Promise<any> {
        return sp.profiles.myProperties.get();
    }

    // Provisioning
    public static ConfigureLists(): Promise<any> {
        // TODO Make ProvisionLists() return a promise
        return Provisioner.ProvisionLists();
    }

    public static CreatePoll(poll: IPollListItem): Promise<IPollListItem>{
        return sp.web.lists.getByTitle(POLL_LIST).items.add({
            Title: poll.PollenQuestion,
            PollenQuestion: poll.PollenQuestion,
            PollenChoices: JSON.stringify(poll.PollenChoices),
            PollenStartDate: poll.PollenStartDate,
            PollenEndDate: poll.PollenEndDate
        }).then((pollResult: ItemAddResult) => {
            return pollResult.data;
        });
    }

    public static GetCurrentPoll(): Promise<IPollListItem> {
        return sp.web.lists.getByTitle(POLL_LIST).items.filter("PollenStartDate le datetime'" + new Date().toISOString() + "'").top(1).orderBy("PollenStartDate", false).get().then((polls) => {
            if(polls.length > 0){
                var poll = polls[0];
                poll.PollenChoices = JSON.parse(poll.PollenChoices);
                
                return poll;
            } else {
                return null;
            }
        });
    }

    public static GetPoll(pollId: number): Promise<IPollListItem> {
        return sp.web.lists.getByTitle(POLL_LIST).items.getById(pollId).get().then(poll => {
            if(poll){
                poll.PollenChoices = JSON.parse(poll.PollenChoices);
                return poll;
            } else {
                return null;
            }
        });
    }

    public static GetPolls(): Promise<IPollListItem[]> {
        return sp.web.lists.getByTitle(POLL_LIST).items.get();
    }

    public static GetPollResponses(pollId: string): Promise<IPollResponse[]> {
        return sp.web.lists.getByTitle(POLL_RESPONSE_LIST).items.filter(`PollenPoll eq ${pollId}`).select("Title", "PollenPoll/ID", "Author/EMail", "PollenResponse").expand("PollenPoll", "Author").get().then((responses: IPollResponse[]) => {
            return responses;
        });
    }

    public static SaveResponse(response: IPollResponse, poll: IPollListItem): Promise<any> {
        return sp.web.lists.getByTitle(POLL_RESPONSE_LIST).items.add({
            'Title': poll.PollenQuestion,
            'PollenPollId': poll.Id,
            'PollenResponse': response.PollenResponse - 1
        }).then((addResult: ItemAddResult) => {
            return addResult.data;
        });
    }
}