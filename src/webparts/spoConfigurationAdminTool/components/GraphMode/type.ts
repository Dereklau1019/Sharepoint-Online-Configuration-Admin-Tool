import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ISpoConfigurationAdminToolWebPartProps } from "../../SpoConfigurationAdminToolWebPart";

export interface IGraphApiTesterProps {
  context: WebPartContext;
  properties: ISpoConfigurationAdminToolWebPartProps;
}
