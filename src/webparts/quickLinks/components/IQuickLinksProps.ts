import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IQuickLinksProps {
  context: WebPartContext;
  listName: string;
  emptyMessage: string;
  componentTitle: string;
}
