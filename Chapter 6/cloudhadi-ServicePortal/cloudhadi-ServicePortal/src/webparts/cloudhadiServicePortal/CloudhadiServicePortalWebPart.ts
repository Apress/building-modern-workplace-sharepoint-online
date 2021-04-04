import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'CloudhadiServicePortalWebPartStrings';
import CloudhadiServicePortal from './components/CloudhadiServicePortal';
import { ICloudhadiServicePortalProps } from './components/ICloudhadiServicePortalProps';
import { sp } from "@pnp/sp";
export interface ICloudhadiServicePortalWebPartProps {
  description: string;
}

export default class CloudhadiServicePortalWebPart extends BaseClientSideWebPart<ICloudhadiServicePortalWebPartProps> {

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
    sp.setup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {
    const element: React.ReactElement<ICloudhadiServicePortalProps > = React.createElement(
      CloudhadiServicePortal,
      {
        description: this.properties.description,
        
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
