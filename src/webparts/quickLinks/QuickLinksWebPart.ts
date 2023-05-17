import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "QuickLinksWebPartStrings";
import QuickLinks from "./components/QuickLinks";
import { IQuickLinksProps } from "./components/IQuickLinksProps";

export interface IQuickLinksWebPartProps {
  listName: string;
  emptyMessage: string;
  componentTitle: string;
  numberOfColumsToShow: any;
}

export default class QuickLinksWebPart extends BaseClientSideWebPart<IQuickLinksWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IQuickLinksProps> = React.createElement(
      QuickLinks,
      {
        context: this.context,
        listName: this.properties.listName,
        emptyMessage: this.properties.emptyMessage,
        componentTitle: this.properties.componentTitle,
        numberOfColumsToShow: this.properties.numberOfColumsToShow,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then((message) => {});
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app
        .getContext()
        .then((context) => {
          let environmentMessage: string = "";
          switch (context.app.host.name) {
            case "Office": // running in Office
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;
            case "Outlook": // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;
            case "Teams": // running in Teams
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error("Unknown host");
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }
  protected onAfterPropertyPaneChangesApplied(): any {
    ReactDom.unmountComponentAtNode(this.domElement);
    this.render();
  }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField("componentTitle", {
                  label: "Component title",
                }),

                PropertyPaneTextField("listName", {
                  label: "Enter a list name",
                }),

                PropertyPaneTextField("emptyMessage", {
                  label: "Text, If no data available",
                }),

                PropertyPaneDropdown("numberOfColumsToShow", {
                  label: "Number of colums to show",

                  options: [
                    {
                      key: "1",

                      text: "1 Columns",
                    },

                    {
                      key: "2",

                      text: "2 Columns",
                    },

                    {
                      key: "3",

                      text: "3 Columns",
                    },
                    {
                      key: "4",

                      text: "4 Columns",
                    },
                  ],

                  selectedKey: "4",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
