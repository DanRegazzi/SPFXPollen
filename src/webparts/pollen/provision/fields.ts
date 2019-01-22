export default class FieldSchema {
    public static PollFields = [
        {
            title: "PollenQuestion",
            fieldType: "SP.FieldMultiLineText",
            properties: {
                FieldTypeKind: 3,
                NumberOfLines: 4,
                Description: "This is the poll question.",
                Required: "TRUE",
                RichText: "TRUE",
                AllowHyperlink: "TRUE",
                Group: "Pollen Columns"
            },
            updateProperties: {
                Title: "Poll Question"
            }
        },
        {
            title: "PollenChoices",
            fieldType: "SP.FieldMultiLineText",
            properties: {
                FieldTypeKind: 3,
                NumberOfLines: 1,
                Description: "This stores the choices for the poll.",
                Required: "TRUE",
                RichText: "FALSE",
                AllowHyperlink: "FALSE",
                Group: "Pollen Columns"
            },
            updateProperties: {
                Title: "Poll Choices"
            }
        },
        {
            title: "PollenStartDate",
            fieldType: "SP.FieldDateTime",
            properties: {
                FieldTypeKind: 4,
                DisplayFormat: 0,
                Description: "If using scheduling, this is the date the poll will begin to display in the webpart.",
                Required: "FALSE",
                Group: "Pollen Columns"
            },
            updateProperties: {
                Title: "Poll Start Date"
            }
        },
        {
            title: "PollenEndDate",
            fieldType: "SP.FieldDateTime",
            properties: {
                FieldTypeKind: 4,
                DisplayFormat: 0,
                Description: "If using scheduling, this is the date the poll will stop being displayed in the webpart.",
                Required: "FALSE",
                Group: "Pollen Columns"
            },
            updateProperties: {
                Title: "Poll End Date"
            }
        }
    ];

    public static ResponseFields = [
        {   // Lookup fields must use the addLookup method
            title: "PollenPoll",
            fieldType: "SP.FieldLookup",
            listName: "Pollen Polls",
            properties: {
                Required: "TRUE"
            },
            updateProperties: {
                Title: "Pollen Poll",
                Description: "The poll that this is in response to.",
                Group: "Pollen Columns"
            }
        },
        {
            title: "PollenResponse",
            fieldType: "SP.FieldNumber",
            properties: {
                FieldTypeKind: 9,
                Description: "Index of poll response.",
                Required: "TRUE",                
                Group: "Pollen Columns"
            },
            updateProperties: {
                Title: "Pollen Response"
            }
        }
    ];
}