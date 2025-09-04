import {MSGraphClientFactory} from "@microsoft/sp-http";
export interface IGetAllusersProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  graphClient:MSGraphClientFactory;
}
