import { ISPMessage } from "./ISPMessage";

export interface IMyMailsState {
    selectedTab: string;
    allMailsCount: number;
    unreadMailsCount: number;
    allMails: any[];
    unreadMails: any[];
    readyToLoadUnread: boolean;
    readyToLoadAllMails: boolean;
    showUserPanel: boolean;
    userInfoForPanel: any;
    readyToLoadPanelData: boolean;
}