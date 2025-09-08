import * as React from 'react';
// import styles from './GraphApiKey.module.scss';
import type { IGraphApiKeyProps } from './IGraphApiKeyProps';
import { escape } from '@microsoft/sp-lodash-subset';

const GraphApiKey:React.FC<IGraphApiKeyProps>=(props)=>{
  return(
    <>
    <div>
    <img src={props.apolloMissionImage.links[0].href}/>
    <div>
      <strong>Title: </strong> {escape(props.apolloMissionImage.data[0].title)}
    </div>
    <div><strong>Keywords:</strong></div>
    <ul>
      {props.apolloMissionImage&&props.apolloMissionImage.data[0].keywords.map((keyword:string)=>
      <li key={keyword}>

        {escape(keyword)}
      </li>)}
    </ul>
    </div>

    </>
  )
}
export default GraphApiKey;