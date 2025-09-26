import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ISpoConfigurationAdminToolWebPartProps } from "../../SpoConfigurationAdminToolWebPart";

export interface IWebpartSettingProps {
  context: WebPartContext;
  properties: ISpoConfigurationAdminToolWebPartProps;
}

export enum EWebPartPropertyRow {
  siteId = "siteId",
  pageUrl = "pageUrl",
  pageId = "pageId",
  pageName = "pageName",
  webpartDetails = "webpartDetails",
  webpartId = "webpartId",
  webPartType = "webPartType",
  webpartName = "webpartName",
  innerHtml = "innerHtml",
  properties = "properties",
  serverProcessedContent = "serverProcessedContent",
}

export interface IWebPartPropertyRow {
  [EWebPartPropertyRow.siteId]?: string;
  [EWebPartPropertyRow.pageUrl]?: string;
  [EWebPartPropertyRow.pageId]?: string;
  [EWebPartPropertyRow.pageName]?: string;
  [EWebPartPropertyRow.webpartId]?: string;
  [EWebPartPropertyRow.webPartType]?: string;
  [EWebPartPropertyRow.webpartDetails]?: any;
  [EWebPartPropertyRow.webpartName]?: string;
  [EWebPartPropertyRow.innerHtml]?: string;
  [EWebPartPropertyRow.properties]?: string;
  [EWebPartPropertyRow.serverProcessedContent]?: string;
}
