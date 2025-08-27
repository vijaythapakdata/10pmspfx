import * as React from 'react';
// import styles from './FunctionalForm.module.scss';
import type { IFunctionalFormProps } from './IFunctionalFormProps';
import { IFunctionalFormState } from './IFunctionalFormState';
import { ChoiceGroup, DatePicker, Dropdown, IDatePickerStrings, IDropdownOption,  PrimaryButton,TextField } from '@fluentui/react';
import {Dialog} from "@microsoft/sp-dialog";
import {Web} from "@pnp/sp/presets/all";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import {PeoplePicker,PrincipalType} from "@pnp/spfx-controls-react/lib/PeoplePicker";

const FunctionalForm:React.FC<IFunctionalFormProps>=(props)=>{
  const[formdata,setFormData]=React.useState<IFunctionalFormState>({
    Name:"",
    Email:"",
    Age:"",
    FullAddress:"",
    Manager:[],
    ManagerId:[],
    Admin:"",
    AdminId:"",
    Department:"",
    Skills:[],
    Gender:"",
    City:"",
    DOB:null
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
  Address:formdata.FullAddress,
  AdminId:formdata.AdminId,
  ManagerId:{results:formdata.ManagerId},
  CityId:formdata.City,
  Department:formdata.Department,
  Gender:formdata.Gender,
  DOB:new Date(formdata.DOB),
  Skills:{results:formdata.Skills}
});
Dialog.alert(`Item create successfully wit ID:${item.data.Id}`);
console.log(item);
setFormData({
  Name:"",
    Email:"",
    Age:"",
    FullAddress:"",
    Manager:[],
    ManagerId:[],
    Admin:"",
    AdminId:"",
     Department:"",
    Skills:[],
    Gender:"",
    City:"",
    DOB:null
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
  //get Admins
  const getAdmin=(items:any[])=>{
    if(items.length>0){
      setFormData(prev=>({...prev,Admin:items[0].text,AdminId:items[0].id}))
    }
    else{
      setFormData(prev=>({...prev,Admin:"",AdminId:""}))
    }
  }
  //Get Managers
  const getManagers=(items:any)=>{
    const managerName=items.map((i:any)=>i.text);
     const managerId=items.map((i:any)=>i.id);
     setFormData(prev=>({...prev,Manager:managerName,ManagerId:managerId}))
  }
//   //on skills change
//   const onSKillsChange=(event:React.FormEvent<HTMLInputElement>,options?:IDropdownProps)=>{
//     if(!options)
//       return;
//     setFormData((prev:any)=>{
// const current=prev.Skills||[];
// const updated=options.selected
//     })
//   }
const onSKillsChange=(event:React.FormEvent<HTMLInputElement>,options:IDropdownOption):void=>{
  const selectedKey=options.selected?[...formdata.Skills,options.key as string]:formdata.Skills.filter((key)=>key!==options.key);
  setFormData(prev=>({...prev,Skills:selectedKey}));
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
    <PeoplePicker
    titleText='Admin'
    context={props.context as any}
    personSelectionLimit={1}
    showtooltip={true}
    required={false}
    ensureUser={true}
    principalTypes={[PrincipalType.User]}
    onChange={getAdmin}
    defaultSelectedUsers={[formdata.Admin?formdata.Admin:""]}
    resolveDelay={1000}
    webAbsoluteUrl={props.siteurl}
    />
    <PeoplePicker
    titleText='Managers'
    context={props.context as any}
    personSelectionLimit={3}
    showtooltip={true}
    required={false}
    ensureUser={true}
    principalTypes={[PrincipalType.User]}
    onChange={getManagers}
    // defaultSelectedUsers={[formdata.Admin?formdata.Admin:""]}
    defaultSelectedUsers={formdata.Manager}
    resolveDelay={1000}
    webAbsoluteUrl={props.siteurl}
    />
    <Dropdown
    placeholder='--select'
    options={props.departmentOptions}
    selectedKey={formdata.Department}
    label='Department'
    onChange={(_,value)=>handleChange("Department",value?.key as string)}
    />
     <Dropdown
    placeholder='--select'
    options={props.cityOptions}
    selectedKey={formdata.City}
    label='City'
    onChange={(_,value)=>handleChange("City",value?.key as string)}
    />
    <ChoiceGroup
     options={props.genderOptions}
    selectedKey={formdata.Gender}
    label='Gender'
    onChange={(_,value)=>handleChange("Gender",value?.key as string)}
    />
      <Dropdown
    placeholder='--select'
    options={props.skillsOptions}
    // selectedKey={formdata.C}
    defaultSelectedKeys={formdata.Skills}
    label='Skills'
    // onChange={(_,value)=>handleChange("City",value?.key as string)}
    onChange={onSKillsChange}
    multiSelect
    />
    <DatePicker
    label='Date of Birth'
    strings={DatePickerStrings}
    value={formdata.DOB}
    formatDate={formateDate}
    onSelectDate={(date)=>setFormData(prev=>({...prev,DOB:date}))}
   
    />
    <br/>
    <PrimaryButton
    text='Save' onClick={createForm} iconProps={{iconName:'save'}}/>
    </>
  )
}
export default FunctionalForm;

export const DatePickerStrings:IDatePickerStrings={
  months:["January","February","March","April","May","June","July","August","September","October","November","December"],
  shortMonths:["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"],
  days:["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"],
  shortDays:["Sun","Mon","Tue","Wed","Thu","Fri","Sat"],
  goToToday:"Go to today",
  prevMonthAriaLabel:"Go to previous month",
  nextMonthAriaLabel:"Go to next month",
  prevYearAriaLabel:"Go to previous year",
  nextYearAriaLabel:"Go to next year"
}
export const formateDate=(date:any):string=>{
  var date1=new Date(date);

  //get year
  var year=date1.getFullYear();
  //get month
  var month=(1+date1.getMonth()).toString();
  month=month.length>1?month:'0'+month;
  //get day
  var day=date1.getDate().toString();
  day=day.length>1?day:'0'+day;

  return month+'/'+day+'/'+year;
  


}
