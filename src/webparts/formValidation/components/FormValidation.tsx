import * as React from 'react';
import styles from './FormValidation.module.scss';
import type { IFormValidationProps } from './IFormValidationProps';
import { Service } from '../../../FormikService/FormikService';
import {sp} from "@pnp/sp/presets/all";
import * as yup from 'yup';
import {Formik,FormikProps} from 'formik';
import { Dialog } from '@microsoft/sp-dialog';
import { DatePicker, Dropdown, Label, PrimaryButton, Stack, TextField } from '@fluentui/react';
import { PeoplePicker,PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
const stackTokens = { childrenGap: 20 };
const FormValidation:React.FC<IFormValidationProps>=(props)=>{
  const [service,setService]=React.useState<Service|null>(null);
  React.useEffect(()=>{
    sp.setup({
      spfxContext:props.context as any
    });
    setService(new Service(props.siteurl));
  },[props.context,props.siteurl]);


  const validationForm=yup.object().shape({
    name:yup.string().required("Task name is required"),
    details:yup.string().min(15,"Minimum 15 charachters are required").required("Task details are required"),
    startDate:yup.date().required("Start date is required"),
    endDate:yup.date().required("End date is reuired"),
    phoneNumber:yup.string().required("Phone number is required").matches(/^[0-9]{10}$/,
      "Phone number must be 10 digits"
    ),
    emailAddress:yup.string().required("Email address is required").matches(/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/,"Invalid email address"),
    projectName:yup.string().required("Project name is required"),
  })

const getFieldProps=(formik:FormikProps<any>,field:string)=>({
  ...formik.getFieldProps(field),errorMessage:formik.errors[field] as string
});
const createRecord=async(record:any)=>{
  try{
const item=await service?.createItems(props.ListName,{
  Title:record.name,
  TaskDetails:record.details,
  StartDate:record.startDate,
  EndDate:record.endDate,
  ProjectName:record.projectName,
  PhoneNumber:record.phoneNumber,
  EmailAddress:record.emailAddress

});
console.log(item);
Dialog.alert("Saved successfully");
  }
  catch(err){
    console.log("Error in creating record: "+err);

  }
}

  return(
    <>
    <Formik
    initialValues={{
      name:"",
      details:"",
      startDate:null,
      endDate:null,
      phoneNumber:"",
      emailAddress:"",
      projectName:""
    }}
    validationSchema={validationForm}
    onSubmit={(values,helpers)=>{
      createRecord(values).then(()=>helpers.resetForm())
    }}
    >
{(formik:FormikProps<any>)=>(
  <form onSubmit={formik.handleSubmit}>
    <div className={styles.formValidation}>
      <Stack tokens={stackTokens}>
        <Label className={styles.lbl}>User Name</Label>
<PeoplePicker
context={props.context as any}
personSelectionLimit={1}
ensureUser={true}
disabled={true}
webAbsoluteUrl={props.siteurl}
defaultSelectedUsers={[props.context.pageContext.user.displayName as any]}
principalTypes={[PrincipalType.User]}
/>
 <Label className={styles.lbl}>Task Name</Label>
 <TextField
 {...getFieldProps(formik,"name")}
 />
  <Label className={styles.lbl}>Email Address</Label>
  <TextField
  {...getFieldProps(formik,"emailAddress")}
  />
   <Label className={styles.lbl}>Phone Number</Label>
   <TextField
   {...getFieldProps(formik,"phoneNumber")}
   />
    <Label className={styles.lbl}>Project Name</Label>
    <Dropdown
    options={[
      {key:"Project 1",text:"Project 1"},
      {key:"Project 2",text:"Project 2"},
      {key:"Project 3",text:"Project 3"},
      {key:"Project 4",text:"Project 4"},
    ]}
    
  selectedKey={formik.values.projectName}
  onChange={(_,options)=>formik.setFieldValue("projectName",options?.key)}
  errorMessage={formik.errors.projectName as string}
    />
     <Label className={styles.lbl}>Start Date</Label>
     <DatePicker
     id="startDate"
     value={formik.values.startDate}
     textField={{...getFieldProps(formik,"startDate")}}
      onSelectDate={(date)=>formik.setFieldValue("startDate",date)}
     />
       <Label className={styles.lbl}>End Date</Label>
     <DatePicker
     id="endDate"
     value={formik.values.endDate}
     textField={{...getFieldProps(formik,"endDate")}}
      onSelectDate={(date)=>formik.setFieldValue("endDate",date)}
     />
     <Label className={styles.lbl}>Task Details</Label>
   <TextField
   {...getFieldProps(formik,"details")}
   multiline
   rows={5}
   />
      </Stack>
      <PrimaryButton
      className={styles.btn}
      text="Submit"
      type="submit"
      iconProps={{iconName:"Save"}}
      />
      <PrimaryButton
      className={styles.btn}
      text="Cancel"
       iconProps={{iconName:"cancel"}}
       onClick={formik.handleReset as any}
      />
      </div>

  </form>
)}
    </Formik>
    </>
  )
}
export default FormValidation;