import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SpfxpnpjsWebPartStrings';
import Spfxpnpjs from './components/Spfxpnpjs';
import { ISpfxpnpjsProps } from './components/ISpfxpnpjsProps';
import { sp } from '@pnp/sp/presets/all';
export interface ISpfxpnpjsWebPartProps {
  description: string;
  //listName: string;
}

export default class SpfxpnpjsWebPart extends BaseClientSideWebPart<ISpfxpnpjsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISpfxpnpjsProps> = React.createElement(
      Spfxpnpjs,
      {
        description: this.properties.description,
        context : this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }
  protected onInit():Promise<void>{
    console.log("onInit Called!! ");
    return super.onInit().then((_) =>{
      sp.setup({spfxContext : this.context});
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  /*protected get dataVersion(): Version {
    return Version.parse('1.0');
  }*/

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
                  //label: strings.ListNameFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
