import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, PropertyPaneTextField } from "@microsoft/sp-property-pane";
import * as strings from 'CreateTeamsWebPartStrings';
import CreateTeams from './components/CreateTeams';
import { ICreateTeamsProps } from './components/ICreateTeamsProps';
import {AppInsights} from "applicationinsights-js";

export interface ICreateTeamsWebPartProps {
  description: string;
}

export default class CreateTeamsWebPart extends BaseClientSideWebPart<ICreateTeamsWebPartProps> {

  public onInit(): Promise<any> {
    /* App Insights key: */
    let appInsightsKey: string = "f98c94fb-b07a-485b-b434-078fbda560dd";

    if(!AppInsights.config) {
      AppInsights.downloadAndSetup({ instrumentationKey: appInsightsKey });
    }

    return Promise.resolve<void>();
  }
  
  public render(): void {
    const element: React.ReactElement<ICreateTeamsProps > = React.createElement(
      CreateTeams,
      {
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
