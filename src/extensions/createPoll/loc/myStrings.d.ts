declare interface ICreatePollCommandSetStrings {
  CreatePoll: string;
  DayPickerStrings: IDatePickerStrings;
}

declare module 'CreatePollCommandSetStrings' {
  const strings: ICreatePollCommandSetStrings;
  export = strings;
}
