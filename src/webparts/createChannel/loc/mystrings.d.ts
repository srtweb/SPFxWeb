declare interface ICreateChannelWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;

  TeamNameLabel: string;
  TeamDescriptionLabel: string;
  Owners: string;
  Members: string;
  CreateChannel: string;
  ChannelName: string;
  ChannelDescription: string;
  Welcome: string;
  Create: string;
  Clear: string;
  CreatingGroup: string;
  CreatingTeam: string;
  CreatingChannel: string;
  Error: string;
  Success: string;
  StartOver: string;
  OpenTeams: string;
}

declare module 'CreateChannelWebPartStrings' {
  const strings: ICreateChannelWebPartStrings;
  export = strings;
}
