declare interface IMyMailsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  TrackInsightsLabel: string;
  AllEmailsLabel: string;
  UnreadEmailsLabel: string;
}

declare module 'MyMailsWebPartStrings' {
  const strings: IMyMailsWebPartStrings;
  export = strings;
}
