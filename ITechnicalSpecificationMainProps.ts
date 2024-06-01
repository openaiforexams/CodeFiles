import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface ITechnicalSpecificationMainProps {  
  userDisplayName: string;
  context:WebPartContext;
  Title:string;
  CurrentUserAccessLevel:string;
}
