import * as React from 'react';
// import styles from './PropertyPaneWebPart.module.scss';
import type { IPropertyPaneWebPartProps } from './IPropertyPaneWebPartProps';
// import { escape } from '@microsoft/sp-lodash-subset';

const PropertyPaneWebPart:React.FC<IPropertyPaneWebPartProps>=(props)=>{
  return(
    <>
    <p><strong>Departments: {props.DropDownField}</strong></p>
    <p><strong>Score: {props.SliderField}</strong></p>
     <p><strong>Gender: {props.ChoiceGroupField}</strong></p>
      <p><strong>Permission: {props.CheckBoxField}</strong></p>
       <p><strong>Toggle: {props.ToggleField}</strong></p>
        <p><strong>Button: {props.buttonField}</strong></p>
    </>
  )
}
export default PropertyPaneWebPart;
