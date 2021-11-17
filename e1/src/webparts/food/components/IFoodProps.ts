import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IFoodProps {
  description: string;
  context: WebPartContext;
  webUrl: string;
}
