import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox
} from '@microsoft/sp-webpart-base';

import * as strings from 'MyMailsWebPartStrings';
import MyMails from './components/MyMails';
import { IMyMailsProps } from './components/IMyMailsProps';
import {AppInsights} from "applicationinsights-js";
import * as microsoftTeams from '@microsoft/teams-js';
import { string } from 'prop-types';

export interface IMyMailsWebPartProps {
  description: string;
  trackInsights: boolean;
  msTeamsContext: microsoftTeams.Context;
}

export default class MyMailsWebPart extends BaseClientSideWebPart<IMyMailsWebPartProps> {
  //private _teamsContext: microsoftTeams.Context;

  public onInit(): Promise<any> {
    /* App Insights key: */
    let appInsightsKey: string = "f98c94fb-b07a-485b-b434-078fbda560dd";

    if(!AppInsights.config) {
      AppInsights.downloadAndSetup({ instrumentationKey: appInsightsKey });
    }

    let retVal: Promise<any> = Promise.resolve();
    if (this.context.microsoftTeams) {
      retVal = new Promise((resolve, reject) => {
        this.context.microsoftTeams.getContext(context => {
          this.properties.msTeamsContext = context;
          resolve();
        });
      });
    }
    return retVal;

    //return Promise.resolve<void>();
  }


  public render(): void {
    const element: React.ReactElement<IMyMailsProps > = React.createElement(
      MyMails,
      {
        description: this.properties.description,
        trackInsights: this.properties.trackInsights,
        context: this.context,
        teamsContext: this.properties.msTeamsContext
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
                }),
                PropertyPaneCheckbox('trackInsights', {
                  checked: false,
                  disabled: false,
                  text: strings.TrackInsightsLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
