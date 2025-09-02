import * as React from 'react';
// import styles from './LargeList.module.scss';
import type { ILargeListProps } from './ILargeListProps';
import { ILargeListState } from './ILargeListState';
import { ServiceClassLL } from '../../../LargeListServiceFile/Service';
import { DetailsList } from '@fluentui/react';
const LargeList:React.FC<ILargeListProps>=(props)=>{
  const [ListResult,setListResult]=React.useState<ILargeListState[]>([]);

const _service=new ServiceClassLL(props.context);

React.useEffect(()=>{
  const fetchData=async()=>{
    try{
const result=await _service.getLargeListItems(props.ListName);
setListResult(result);
    }
    catch(err)
{
console.log("Error while fetching data....",err);
throw err;
};

}
fetchData();
},[props.ListName,_service])
  return(
    <>
    <DetailsList
    items={ListResult}
    />
    </>
  )
}
export default LargeList;