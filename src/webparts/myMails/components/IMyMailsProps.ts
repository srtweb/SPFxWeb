import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as microsoftTeams from '@microsoft/teams-js';

export interface IMyMailsProps {
  description: string;
  trackInsights: boolean;
  context: WebPartContext;
  teamsContext: microsoftTeams.Context;
}

export interface DisplayMailsProps {
  mailsToDisplay: any[];
  facePileClick: any;
  readyToLoad: boolean;
}