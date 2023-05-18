import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface INewsWebpartProps {
  context: WebPartContext;
  listName: string;
  emptyMessage: string;
  componentTitle: string;
}
