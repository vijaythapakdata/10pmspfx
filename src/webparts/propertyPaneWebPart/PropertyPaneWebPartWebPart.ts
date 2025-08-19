import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
PropertyPaneDropdown,PropertyPaneSlider,
PropertyPaneChoiceGroup,
PropertyPaneButton, PropertyPaneToggle, PropertyPaneCheckbox
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';


import * as strings from 'PropertyPaneWebPartWebPartStrings';
import PropertyPaneWebPart from './components/PropertyPaneWebPart';
import { IPropertyPaneWebPartProps } from './components/IPropertyPaneWebPartProps';

export interface IPropertyPaneWebPartWebPartProps {
  DropDownField: string;
  SliderField:any;
   ChoiceGroupField:string;
   CheckBoxField:boolean;
 ToggleField:boolean;
 buttonField:string;
}

export default class PropertyPaneWebPartWebPart extends BaseClientSideWebPart<IPropertyPaneWebPartWebPartProps> {



  public render(): void {
    const element: React.ReactElement<IPropertyPaneWebPartProps> = React.createElement(
      PropertyPaneWebPart,
      {
        DropDownField:this.properties.DropDownField,
        SliderField:this.properties.SliderField,
         ChoiceGroupField:this.properties.ChoiceGroupField,
         CheckBoxField:this.properties.CheckBoxField,
          ToggleField:this.properties.ToggleField,
          buttonField:this.properties.buttonField
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

  protected get disableReactivePropertyChanges():boolean{
    return true;
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
               PropertyPaneDropdown('DropDownField',{
                label:'Dropdown',
                options:[
                  {key:'HR',text:'HR'},
                  {key:'IT',text:'IT'}
                ]
               }),
               PropertyPaneSlider('SliderField',{
                label:'Slider',
                min:0,
                max:100,
                step:1
               }),
               PropertyPaneChoiceGroup('ChoiceGroupField',{
                label:'Choice Group',
                options:[
                  {key:'Male',text:'Male'},
                  {key:'Female',text:'Female'}
                ]
               }),
               PropertyPaneButton('buttonField',{
                text:'Button',
                onClick:()=>alert('Button Clicked')
               }),
               PropertyPaneCheckbox('CheckBoxField',{
                text:'Check Box',
               
               }),
               PropertyPaneToggle('ToggleField',{
                label:'Toggle',
                onText:'On',
                offText:'Off'
               })
              ]
            }
          ]
        }
      ]
    };
  }
}
