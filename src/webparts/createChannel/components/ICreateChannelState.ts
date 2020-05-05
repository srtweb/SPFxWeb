import { IMyTeams } from './ICreateChannelProps';

export interface ICreateChannelState {
    teamName: any;
    owners?: string[];
    members?: string[];
    channelName?: string;
    channelDescription?: string;
    channelType?: any;
    spinnerText?: string;
    creationState?: CreationState;
    myTeams: IMyTeams[];
    channelUrl?: string;
    messageToDisplay?: string;
    
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