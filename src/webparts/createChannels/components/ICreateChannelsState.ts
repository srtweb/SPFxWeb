import { IMyTeams } from './ICreateChannelsProps';

export interface ICreateChannelsState {
    messageToDisplay: string;
    myTeams: IMyTeams[];
    selectedTeam: any;
    existingChannels: IMyTeams[];
    newChannel: string;
    createNewChannel: boolean;
}

