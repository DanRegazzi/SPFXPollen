export default class ContentTypeSchema {
    
    public static PollContentType = {
        Id: '0x01002DE89328A3F8407DA5ED68C87A8BCFEE',
        Name: 'Poll Item',
        Description: 'Represents a Poll',
        Group: "Pollen Content Types",
        Fields: [
            "PollenQuestion",
            "PollenChoices",
            "PollenStartDate",
            "PollenEndDate"
        ]
    };

    public static ResponseContentType = {
        Id: '0x0100D2E698AF3AD549E8AD5825B981BD2037',
        Name: 'Poll Response',
        Description: 'Represents a Poll Response',
        Group: "Pollen Content Types",
        Fields: [
            "PollenPoll",
            "PollenResponse"
        ]
    };    
}