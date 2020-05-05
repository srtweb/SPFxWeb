import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ICreateTeamsProps {
  context: WebPartContext;
}

export interface IMyTeams {
  text: string;
  key: string;
}
