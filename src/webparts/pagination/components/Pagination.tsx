import * as React from 'react';
// import styles from './Pagination.module.scss';
import type { IPaginationProps } from './IPaginationProps';
import {Table,Input} from 'antd';
import {sp} from '@pnp/sp/presets/all';
const Pagination:React.FC<IPaginationProps>=(props)=>{
  const [items,setItems]=React.useState<any[]>([]);
  const[search,setSearch]=React.useState<string>('');

  React.useEffect(()=>{
    sp.setup({
      spfxContext:props.context as any
    });
    sp.web.lists.getByTitle(props.ListName).items.select("Id","Title","EmailAddress","Age","Admin/Title","City/Title").expand("Admin","City").get().then((data)=>{
      const _readItems=data.map((e)=>({
        key:e.Id,
        Title:e.Title,
        EmailAddress:e.EmailAddress,
        Age:e.Age,
        Admin:e.Admin?.Title,
        City:e.City?.Title
      }));
      setItems(_readItems)
    })
    .catch((err)=>(
      console.log(err)
    ))

  },[props.context]);

  const columns=[
    {
      title:"Name",
      dataIndex:"Title",
      key:"Title",
      sorter:(a:any,b:any)=>a.Title.localCompare(b.Title)
    },
     {
      title:"Email Address",
      dataIndex:"EmailAddress",
      key:"EmailAddress",
      sorter:(a:any,b:any)=>a.EmailAddress.localCompare(b.EmailAddress)
    },
    {
      title:"Age",
      dataIndex:"Age",
      Key:"Age"
    },
    {
      title:"City",
      dataIndex:"City",
      Key:"City"
    },
    {
      title:"Admin",
      dataIndex:"Admin",
      Key:"Admin"
    },
  ]
  const handleSearch=(event:React.ChangeEvent<HTMLInputElement>)=>{
    setSearch(event.target.value);
  }
  return(
    <></>
  )

}
export default Pagination;