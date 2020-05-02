import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ICreateChannelProps {
  context: WebPartContext;
}

export interface IMyTeams {
  text: string;
  key: string;
}
