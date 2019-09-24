import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'MyMailsWebPartStrings';
import MyMails from './components/MyMails';
import { IMyMailsProps } from './components/IMyMailsProps';
import {AppInsights} from "applicationinsights-js";

export interface IMyMailsWebPartProps {
  description: string;
}

export default class MyMailsWebPart extends BaseClientSideWebPart<IMyMailsWebPartProps> {

  public onInit(): Promise<void> {
    /* App Insights key: */
    let appInsightsKey: string = "a40cf729-7e67-440d-a932-80f32d84f39e";

    AppInsights.downloadAndSetup({ instrumentationKey: appInsightsKey });

    return Promise.resolve<void>();
  }


  public render(): void {
    const element: React.ReactElement<IMyMailsProps > = React.createElement(
      MyMails,
      {
        description: this.properties.description,
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
