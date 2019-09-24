import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMyMailsProps {
  description: string;
  context: WebPartContext;
}

export interface DisplayMailsProps {
  mailsToDisplay: any[];
  facePileClick: any;
  readyToLoad: boolean;
}