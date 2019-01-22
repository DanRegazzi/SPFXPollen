declare interface IPollenWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  PollQuestionFieldLabel: string;
  SchedulerFieldLabel: string;
  DayPickerStrings: IDatePickerStrings;
}

declare module 'PollenWebPartStrings' {
  const strings: IPollenWebPartStrings;
  export = strings;
}