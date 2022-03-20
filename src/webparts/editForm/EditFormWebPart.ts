import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'EditFormWebPartStrings';
import EditForm from './components/EditForm';
import { IEditFormProps } from './components/IEditFormProps';

export interface IEditFormWebPartProps {
  description: string;
}

export default class EditFormWebPart extends BaseClientSideWebPart<IEditFormWebPartProps> {


  protected onInit(): Promise<void> {
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IEditFormProps> = React.createElement(
      EditForm,
      {
        description: this.properties.description,
        webURL: this.context.pageContext.web.absoluteUrl,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
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
