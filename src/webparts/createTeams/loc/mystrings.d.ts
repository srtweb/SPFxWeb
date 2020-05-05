declare interface ICreateTeamsWebPartStrings {
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
  AddingMembers: string;
  Error: string;
  Success: string;
  cSuccess: string;
  StartOver: string;
  OpenTeams: string;
}

declare module 'CreateTeamsWebPartStrings' {
  const strings: ICreateTeamsWebPartStrings;
  export = strings;
}
