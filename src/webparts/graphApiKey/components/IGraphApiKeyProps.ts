import { INasaItem } from "../../Models/INasaImageSearchResponse";
export interface IGraphApiKeyProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
apolloMissionImage:INasaItem;
}
