import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ColumnPickerAndListViewWebPartStrings';
import ColumnPickerAndListView from './components/ColumnPickerAndListView';
import { IColumnPickerAndListViewProps } from './components/IColumnPickerAndListViewProps';
import { spfi, SPFx } from '@pnp/sp';
// import sp from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
// import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
// import { PropertyFieldColumnPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldColumnPicker';
export interface IColumnPickerAndListViewWebPartProps {
  selectedColumns: any;
  listId: string | undefined;
  description: string;
  ListTitleFieldLabel: any;
}

export default class ColumnPickerAndListViewWebPart extends BaseClientSideWebPart<IColumnPickerAndListViewWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _sp: any;


  public render(): void {
    const element: React.ReactElement<IColumnPickerAndListViewProps> = React.createElement(
      ColumnPickerAndListView,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        sp: this._sp,
        listId: this.properties.listId
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._sp = spfi().using(SPFx(this.context));
    return super.onInit();
  }



  // private _getEnvironmentMessage(): Promise<string> {
  //   if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
  //     return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
  //       .then(context => {
  //         let environmentMessage: string = '';
  //         switch (context.app.host.name) {
  //           case 'Office': // running in Office
  //             environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
  //             break;
  //           case 'Outlook': // running in Outlook
  //             environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
  //             break;
  //           case 'Teams': // running in Teams
  //           case 'TeamsModern':
  //             environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
  //             break;
  //           default:
  //             environmentMessage = strings.UnknownEnvironment;
  //         }

  //         return environmentMessage;
  //       });
  //   }

  //   return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  // }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

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
                PropertyPaneTextField('listId', {
                  label: 'List ID',
                  description: 'Enter the ID of the SharePoint list to use'
                }),
                // PropertyFieldListPicker('listId', {
                //   label: strings.ListTitleFieldLabel,
                //   selectedList: this.properties.listId as string,
                //   includeHidden: false,
                //   orderBy: PropertyFieldListPickerOrderBy.Title,
                //   multiSelect: false,
                //   disabled: false,
                //   onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                //   properties: this.properties,
                //   context: this.context as any,
                //   onGetErrorMessage: undefined,
                //   deferredValidationTime: 0,
                //   key: 'listPickerField'
                // })
                // PropertyFieldListPicker('listId', {
                //   label: 'Select List',
                //   context: this.context as any,
                //   onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                //   properties: this.properties,
                //   key: 'listPicker'
                // }),
                // PropertyFieldColumnPicker('selectedColumns', {
                //   label: 'Select Columns',
                //   context: this.context as any,
                //   listId: this.properties.listId,
                //   selectedColumn: this.properties.selectedColumns,
                //   onPropertyChange: this.onPropertyPaneFieldChanged,
                //   properties: this.properties,
                //   key: 'columnPicker'
                // })
              ]
            },

          ]
        }
      ]
    };
  }
}
