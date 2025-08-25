import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IFunctionalFormProps {
 
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  siteurl:string;
  ListName:string;
  context:WebPartContext;
  genderOptions:any;
  departmentOptions:any;
  skillsOptions:any;
  cityOptions:any;
}
