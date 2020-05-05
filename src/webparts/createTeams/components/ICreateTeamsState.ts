import { IMyTeams } from './ICreateTeamsProps';

export interface ICreateTeamsState {
   teamName?: string;
   teamDescription?: string;
   cowners?: string[];
   cmembers?: string[];
   towners?: string[];
   tmembers?: string[];
   createChannel?: boolean;
   cchannelName?: string;
   cchannelDescription?: string;
   spinnerText?: string;
   creationState?: CreationState;
   channelTeam: any;
   cmyTeams?: IMyTeams[];
   cselectedTeam?: any;
   cchannelType?: any;
   cchannelUrl?: string;
   Success?: string;
   buttonText?: string;
   messageToDisplay?: string;
   channelTeamUrl?: string;
   teamMembers?: string[];
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