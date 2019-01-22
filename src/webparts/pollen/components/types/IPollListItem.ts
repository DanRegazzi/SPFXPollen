export default interface IPollListItem {
    Id?: string;
    Title: string;
    PollenQuestion: string;
    PollenChoices: string[];
    PollenStartDate: Date;
    PollenEndDate: Date;
}