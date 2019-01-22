export default interface IPollResponse {
    Id?: string;
    Title: string;
    Author?: {
        EMail: string;
    };
    PollenPoll: {
        ID: string;
    };
    PollenResponse: number;
}