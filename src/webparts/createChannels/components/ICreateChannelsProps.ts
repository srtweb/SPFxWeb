import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ICreateChannelsProps {
  context: WebPartContext;
}

export interface IMyTeams {
  text: string;
  key: string;
}
