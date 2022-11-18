import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'BulkchangeWebPartStrings';
import Bulkchange from './components/Bulkchange';
import { IBulkchangeProps } from './components/IBulkchangeProps';
import { sp } from '@pnp/sp';
export interface IBulkchangeWebPartProps {
  description: string;
}

export default class BulkchangeWebPart extends BaseClientSideWebPart<IBulkchangeWebPartProps> {
  public onInit() {
    sp.setup({ spfxContext: this.context });
    return super.onInit();
  }
  public render(): void {
    const element: React.ReactElement<IBulkchangeProps> = React.createElement(
      Bulkchange,
      {
        //description: this.properties.description
        url: this.context.pageContext.web.absoluteUrl,
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
