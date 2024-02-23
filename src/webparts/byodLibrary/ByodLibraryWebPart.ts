import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneButton,
  PropertyPaneCheckbox,
  PropertyPaneChoiceGroup,
  PropertyPaneLabel,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { PropertyFieldIconPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldIconPicker';
import { PropertyFieldFilePicker, IFilePickerResult } from "@pnp/spfx-property-controls/lib/PropertyFieldFilePicker";

import * as strings from 'ByodLibraryWebPartStrings';
import ByodLibrary from './components/ByodLibrary';
import { IByodLibraryProps } from './components/IByodLibraryProps';
import { PropertyFieldPeoplePicker, IPropertyFieldGroupOrPerson, PrincipalType } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker'; 
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

export interface IPropertyControlsTestWebPartProps {
  people: IPropertyFieldGroupOrPerson[];
}
export interface IPropertyControlsTestWebPartProps {
  color: string;
}
export interface IByodLibraryWebPartProps {
  description: string;

  context: WebPartContext;
  targetAudience: any;
  siteUrl: string;
  listName: string;
  isExp: boolean;
  color: string;
  openInNewTab: boolean;
  showDivider: boolean;
  sectionTitle: string;
  isCollapsible: boolean;
  iconAlignment: string;
  iconPicker: any;
  thumbnail: any;
  customImgPicker: any;
  groupBy: boolean;
  groupByField: string;
  sectionDescription: string;
  enableSearch: boolean;
  searchPlaceholder: string;
  enableTargetAudience: boolean;
}

export default class ByodLibraryWebPart extends BaseClientSideWebPart<IByodLibraryWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IByodLibraryProps> = React.createElement(
      ByodLibrary,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,

        context: this.context,
        targetAudience: this.properties.targetAudience,
        siteUrl: this.properties.siteUrl,
        listName: this.properties.listName,
        isExp: this.properties.isExp,
        color: this.properties.color,
        openInNewTab : this.properties.openInNewTab,
        showDivider: this.properties.showDivider,
        sectionTitle: this.properties.sectionTitle,
        isCollapsible: this.properties.isCollapsible,
        iconAlignment: this.properties.iconAlignment,
        iconPicker: this.properties.iconPicker,
        thumbnail: this.properties.thumbnail,
        customImgPicker: this.properties.customImgPicker,
        groupBy: this.properties.groupBy,
        groupByField: this.properties.groupByField,
        sectionDescription: this.properties.sectionDescription,
        enableSearch: this.properties.enableSearch,
        searchPlaceholder: this.properties.searchPlaceholder,
        enableTargetAudience: this.properties.enableTargetAudience
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

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

  protected onGotoSiteAssetsClick(){
    if (this.context){
      const siteAssetsUrl = `${this.context.pageContext.web.absoluteUrl}/SiteAssets` ;
      window.open(siteAssetsUrl, '_blank');
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          // header: {
          //   description: strings.PropertyPaneDescription
          // },
          groups: [
            {
              groupName: 'List Properties',
              groupFields: [
                PropertyPaneTextField('siteUrl', {
                  label: 'Site Path',
                  value: this.properties.siteUrl,
                  
                }),
                PropertyPaneTextField('listName', {
                  label: 'List Name',
                  value: this.properties.listName,
                })
              ]
            },
            {
              groupName: 'Section Properties',
              groupFields:[
                PropertyPaneTextField('sectionTitle', {
                  label: 'Title',
                  value: this.properties.sectionTitle
                }),
                PropertyPaneTextField('sectionDescription', {
                  label: 'Description',
                  value: this.properties.sectionDescription,
                  multiline: true
                }),
                // PropertyPaneCheckbox('openInNewTab', {
                //   checked: this.properties.openInNewTab,
                //   text: 'Open link in a new tab',
                // }),
                // PropertyPaneCheckbox('groupBy', {
                //   checked: this.properties.groupBy,
                //   text: 'Group By',
                // }),
                // PropertyPaneTextField('groupByField', {
                //   label: 'Field',
                //   value: this.properties.groupByField,
                //   disabled: !this.properties.groupBy
                // }),
                PropertyPaneCheckbox('enableSearch', {
                  checked: this.properties.enableSearch,
                  text: 'Enable Search',
                }),
                PropertyPaneTextField('searchPlaceholder', {
                  label: 'Search Placeholder',
                  value: this.properties.searchPlaceholder,
                  disabled: !this.properties.enableSearch
                }),
                PropertyPaneToggle('enableTargetAudience', {
                  label: 'Enable audience targeting',
                  onText: 'On',
                  offText: 'Off',
                  checked: this.properties.enableTargetAudience,
                }),
              ]
            },
            {
              groupName: 'Display',
              groupFields:[
                PropertyPaneCheckbox('isCollapsible', {
                  checked: this.properties.isCollapsible,
                  text: 'Make this section collapsible',
                }),
                PropertyPaneToggle('isExp', {
                  label: 'Default display',
                  onText: 'Expanded',
                  offText: 'Collapsed',
                  checked: this.properties.isExp
                }),
                PropertyPaneToggle('iconAlignment', {
                  label: 'Expand/collapse icon alignment',
                  onText: 'Right',
                  offText: 'Left',
                  checked: this.properties.iconAlignment === 'Right'
                })
              ]
            },
            {
              groupName: 'Look and Feel',
              groupFields:[
                PropertyFieldColorPicker('color', {
                  label: 'Color Theme ' + this.properties.color,
                  selectedColor: this.properties.color,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  debounce: 1000,
                  isHidden: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'colorFieldId'
                }),
                PropertyPaneCheckbox('showDivider', {
                  checked: this.properties.showDivider,
                  text: 'Show divider'
                }),
                PropertyPaneChoiceGroup('thumbnail', {
                  label: 'Thumbnail',
                  options : [
                    {key: 'auto', text:'Auto-selected', checked: true}, 
                    {key: 'customImg', text:'Custom image'}, 
                    {key: 'icon', text:'Icon'}, 
                  ],
                }),
                PropertyFieldIconPicker('iconPicker', {
                  // currentIcon: this.properties.iconPicker,
                  key: "iconPickerId",
                  onSave: (icon: string) => { console.log(icon); this.properties.iconPicker = icon; },
                  onChanged:(icon: string) => { console.log(icon);  },
                  buttonLabel: "Change",
                  renderOption: "panel",
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  buttonClassName:  'iconPicker-' + this.properties.thumbnail ,     
                }),
                PropertyFieldFilePicker('customImgPicker', {
                  buttonIcon:'',
                  context: this.context as any,
                  filePickerResult: this.properties.customImgPicker,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onSave: (e: IFilePickerResult) => { console.log(e); this.properties.customImgPicker = e;  },
                  onChanged: (e: IFilePickerResult) => { console.log(e); this.properties.customImgPicker = e; },
                  key: "customImgPickerId",
                  buttonLabel: "Change",
                  buttonClassName: 'customImgPicker-btn filePicker-' + this.properties.thumbnail,
                  hideLocalUploadTab : true,
                  hideLinkUploadTab: true,
                  allowExternalLinks: true,
                  accepts: [".gif", ".jpg", ".jpeg", ".bmp", ".dib", ".tif", ".tiff", ".ico", ".png", ".jxr", ".svg"]
                }),
                PropertyPaneLabel('customImgNote', {
                  text: 'To upload a custom image to site assests please use the button below',
                }),
                PropertyPaneButton('goToSiteAssetsBtn', {
                  text: 'Go to Site Assets',
                  onClick: this.onGotoSiteAssetsClick.bind(this)
                })
              ]
            },
            {
              groupName: 'Target Audience',
              groupFields: [
                PropertyFieldPeoplePicker('targetAudience', {
                  label: 'Target Audience e.g. User(s), Group(s)',
                  initialData: this.properties.targetAudience,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users, PrincipalType.SharePoint, PrincipalType.Security],
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  context: this.context as any,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'peopleFieldId'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
