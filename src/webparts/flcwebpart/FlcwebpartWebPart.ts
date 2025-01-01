import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneButton,
  PropertyPaneCheckbox,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneHorizontalRule,
  PropertyPaneLabel,
  PropertyPaneLink,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'FlcwebpartWebPartStrings';
import Flcwebpart from './components/Flcwebpart';
import { IFlcwebpartProps } from './components/IFlcwebpartProps';

export interface IFlcwebpartWebPartProps {
  description: string;
  textField: string;
  multiLineTextField: string;
  linkField: string;
  dropdownField: string;
  choiceGroupField: string;
  sliderField: number;
  toggleField: boolean;
  checkboxField: boolean;
  buttonField: string;
}

export default class FlcwebpartWebPart extends BaseClientSideWebPart<IFlcwebpartWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IFlcwebpartProps> = React.createElement(
      Flcwebpart,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        textField: this.properties.textField,
        multiLineTextField: this.properties.multiLineTextField,
        linkField: this.properties.linkField,
        sliderField: this.properties.sliderField,
        dropdownField: this.properties.dropdownField,
        choiceGroupField: this.properties.choiceGroupField,
        toggleField: this.properties.toggleField,
        checkboxField: this.properties.checkboxField,
        buttonField: this.properties.buttonField
        
     
        
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
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
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
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
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

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            image: "https://www.google.com/images/branding/googlelogo/1x/googlelogo_color_272x92dp.png",
            description: strings.PropertyPaneHeading

          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('textField', {
                  label: strings.TexboxFieldLabel
                }),
                PropertyPaneTextField('multiLineTextField', {
                  label: 'Multi-line Text Field',
                  multiline: true
                }),
                PropertyPaneLink('linkField', {
                  text: 'Example Link',
                  href: 'https://www.example.com'
                }),
                PropertyPaneDropdown('dropdownField', {
                  label: 'Dropdown',
                  options: [
                    { key: 'Option1', text: 'Option 1' },
                    { key: 'Option2', text: 'Option 2' }
                  ]
                }),
                PropertyPaneChoiceGroup('choiceGroupField', {
                  label: 'Choice Group',
                  options: [
                    { key: 'A', text: 'Option A' },
                    { key: 'B', text: 'Option B' }
                  ]
                }),
                PropertyPaneSlider('sliderField', {
                  label: 'Slider',
                  min: 0,
                  max: 100
                }),
                PropertyPaneToggle('toggleField', {
                  label: 'Toggle'
                }),
                PropertyPaneCheckbox('checkboxField', {
                  text: 'Checkbox'
                }),
                PropertyPaneButton('buttonField', {
                  text: 'Button',
                  onClick: () => alert('Button clicked!')
                }),
                PropertyPaneHorizontalRule()
              ]
            }
          ]
        },
        {
          header: {
            description: "My Web Part Configuration Page 2"
          },
          groups: [
            {
              groupName: "Basic Settings 2",
              groupFields: [
                PropertyPaneLabel('labelField2', {
                  text: 'This is a label'
                }),
                PropertyPaneTextField('textField2', {
                  label: 'Text Field'
                }),
                PropertyPaneTextField('multiLineTextField2', {
                  label: 'Multi-line Text Field',
                  multiline: true
                }),
                PropertyPaneHorizontalRule()
              ]
            }
          ]
        } 
      ]
    };
  }
}
