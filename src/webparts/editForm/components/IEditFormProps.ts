import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IEditFormProps {
  description: string;
  context: WebPartContext;
  webURL: string;
}
