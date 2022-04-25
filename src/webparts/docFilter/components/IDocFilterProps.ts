import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDocFilterProps {
  description: string;
  webpartName: string;
  sharePointList: string;
  sharePointView: string;
  sharePointLink: string;
  sharePointColumn: string;
  context: WebPartContext
}
