export interface ICreateTeamsState {
   teamName: string;
   teamDescription?: string;
   owners?: string[];
   members?: string[];
   createChannel?: boolean;
   channelName?: string;
   channelDescription?: string;
   spinnerText?: string;
   creationState?: CreationState;
}

export enum CreationState {
    /**
     * Initial state - user input
     */
    notStarted = 0,
    /**
     * creating all selected elements (group, team, channel, tab)
     */
    creating = 1,
    /**
     * everything has been created
     */
    created = 2,
    /**
     * error during creation
     */
    error = 4
  }