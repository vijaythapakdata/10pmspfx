import * as React from 'react';
// import styles from './FunctionalForm.module.scss';
import type { IFunctionalFormProps } from './IFunctionalFormProps';
import { IFunctionalFormState } from './IFunctionalFormState';
import { PrimaryButton,TextField } from '@fluentui/react';
import {Dialog} from "@microsoft/sp-dialog";
import {Web} from "@pnp/sp/presets/all";
import "@pnp/sp/items";
import "@pnp/sp/lists";

const FunctionalForm:React.FC<IFunctionalFormProps>=(props)=>{
  const[formdata,setFormData]=React.useState<IFunctionalFormState>({
    Name:"",
    Email:"",
    Age:"",
    FullAddress:""
  });
  //create form
  const createForm=async()=>{
    try{
const web=Web(props.siteurl); //this is the site url
const list=web.lists.getByTitle(props.ListName);
const item=await list.items.add({
  Title:formdata.Name,
  EmailAddress:formdata.Email,
  Age:parseInt(formdata.Age),
  Address:formdata.FullAddress
});
Dialog.alert(`Item create successfully wit ID:${item.data.Id}`);
console.log(item);
setFormData({
  Name:"",
    Email:"",
    Age:"",
    FullAddress:""
});
    }
    catch(err){
console.error(err);
Dialog.alert(`Error while creating item:${err}`);
    }
  }
  //event handlers
  const handleChange=(fieldvalue:keyof IFunctionalFormState,value:string|number|boolean)=>{
    setFormData(prev=>({...prev,[fieldvalue]:value}))
  }
  return(
    <>
    <TextField
    label='Name'
    value={formdata.Name}
    onChange={(_,value)=>handleChange("Name",value||"")}
    placeholder='Enter your name'
    iconProps={{iconName:'people'}}
    />
     <TextField
    label='Email Address'
    value={formdata.Email}
    onChange={(_,value)=>handleChange("Email",value||"")}
    placeholder='Enter your email address'
    iconProps={{iconName:'mail'}}
    />
     <TextField
    label='Age'
    value={formdata.Age}
    onChange={(_,value)=>handleChange("Age",value||"")}
    
    />
     <TextField
    label='Full Address'
    value={formdata.FullAddress}
    onChange={(_,value)=>handleChange("FullAddress",value||"")}
    placeholder='Enter your address....'
    iconProps={{iconName:'home'}}
    multiline
    rows={5}
    />
    <br/>
    <PrimaryButton
    text='Save' onClick={createForm} iconProps={{iconName:'save'}}/>
    </>
  )
}
export default FunctionalForm;
