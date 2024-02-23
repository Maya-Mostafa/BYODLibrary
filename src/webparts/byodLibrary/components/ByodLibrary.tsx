import * as React from 'react';
import styles from './ByodLibrary.module.scss';
import './ByodLibrary.scss';
import { IByodLibraryProps } from './IByodLibraryProps';
import { getGraphMemberOf, isFromTargetAudience, groupBy, getListItemsGraph } from '../services/requests';
// import { escape } from '@microsoft/sp-lodash-subset';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import {  faChevronUp, faChevronDown } from '@fortawesome/free-solid-svg-icons';
import LibraryItem from './LibraryItem/LibraryItem';
import Search from './Search/Search';

export default function ByodLibrary(props: IByodLibraryProps) {

  const [showBasedOnTargetAudience, setShowBasedOnTargetAudience] = React.useState(false);
  const [items, setItems] = React.useState([]);
  const [filteredItems, setFilteredItems] = React.useState(items);
  const [isExp, setIsExp] = React.useState(props.isExp && props.isCollapsible);
  const [memberOfGroups, setMemberofGroups] = React.useState(null);
  
  const [categories, setCategories] = React.useState([]);
  const [categItems, setCategItems] = React.useState(null);
  const [filteredCategItems, setFilteredCategItems] = React.useState(categItems);

  React.useEffect(()=>{
    console.log("props.context", props.context);

    getListItemsGraph(props.context).then(res => {
      console.log("graph", res);
      if (props.groupBy){
        const groupedArr = groupBy(res, props.groupByField);
        setCategItems(groupedArr);
        setFilteredCategItems(groupedArr);
        setCategories(Object.keys(groupedArr));
      }
      setItems(res);
      setFilteredItems(res);
    });
    getGraphMemberOf(props.context).then((res: any) => {
      console.log("grpahMemberOf", res);
      setMemberofGroups(res);
    });

    /*getListItems(props.context, props.siteUrl, props.listName).then(res => {
      
      if (props.groupBy){
        const groupedArr = groupBy(res, props.groupByField);
        setCategItems(groupedArr);
        setFilteredCategItems(groupedArr);
        setCategories(Object.keys(groupedArr));
        
        console.log("groupBy", groupedArr);
        console.log("Object.keys(res)", Object.keys(groupedArr))
      }

      setItems(res);
      setFilteredItems(res);
    });*/

    if (props.targetAudience && props.targetAudience.length > 0){
      setShowBasedOnTargetAudience(isFromTargetAudience(props.context, memberOfGroups, props.targetAudience, 'fullName'));
    }else{
      setShowBasedOnTargetAudience(true);
    }
  }, []);

  const onSearchChanged = (_: any, text: string): void => {
    if (props.groupBy) setFilteredCategItems(items.filter(item => item.Title.toLowerCase().indexOf(text.toLowerCase()) >= 0));
    else setFilteredItems(items.filter(item => item.fields.Title.toLowerCase().indexOf(text.toLowerCase()) >= 0));
  };
  

  return (
    <>
    {showBasedOnTargetAudience &&
      <section className={`${styles.byodLibrary} ${props.hasTeamsContext ? styles.teams : ""}`}>
        {/* <div>Web part property value:{" "} <strong>{escape(props.description)}</strong></div>
        <img alt='' className={styles.welcomeImage}
          src={
            props.isDarkTheme
              ? require("../assets/welcome-dark.png")
              : require("../assets/welcome-light.png")
          }
        /> */}

        <div className={styles.main}>
          <div className={styles.librarySection} style={{borderBottom: props.showDivider ? '1px solid #ddd' : '1px solid #fff'}}>
            <h6 className={styles.sectionHdr} style={{color: props.color, borderBottomColor: props.color}} onClick={()=> props.isCollapsible && setIsExp(prev=>!prev)}>
              <span>{props.sectionTitle ? props.sectionTitle : props.listName}</span>
              {props.isCollapsible &&
                <span className={props.iconAlignment ? styles.colExpRight : styles.colExpLeft}>
                  {isExp ?
                    <FontAwesomeIcon icon={faChevronDown} />
                    :
                    <FontAwesomeIcon icon={faChevronUp} />
                  }
                </span>
              }
            </h6>
            
            <div style={{display: props.isCollapsible ? (isExp ? 'block' : 'none') : 'block'}}>
              <div>{props.sectionDescription}</div>
              {props.enableSearch &&
                <Search 
                  searchPlaceholder={props.searchPlaceholder}
                  onSearchChanged={onSearchChanged}
                />
              }
              {props.groupBy ?
                categories.map((categ: any) => {
                  return(
                    <div key={categ}>
                      <h5 className={styles.secCategory}>{categ}</h5>
                      <ul className={styles.cards}>
                        {filteredCategItems[categ].map((item: any) => {
                          return(
                            <LibraryItem 
                              item={item.fields} 
                              customImgPicker={props.customImgPicker} 
                              iconPicker={props.iconPicker} 
                              thumbnail={props.thumbnail} 
                              key={item.fields.id} 
                          />
                          )
                        })}
                      </ul>
                    </div>
                  );
                })
                :
                <ul className={styles.cards}>
                  {filteredItems.map((item: any)=>{
                    if (!props.enableTargetAudience || 
                        !item.fields._ModernAudienceTargetUserField || 
                        props.enableTargetAudience && 
                        memberOfGroups && 
                        item.fields._ModernAudienceTargetUserField && 
                        isFromTargetAudience(props.context, memberOfGroups, item.fields._ModernAudienceTargetUserField, 'LookupValue')){
                      return(                                            
                        <LibraryItem 
                          item={item.fields} 
                          customImgPicker={props.customImgPicker} 
                          iconPicker={props.iconPicker} 
                          thumbnail={props.thumbnail} 
                          key={item.fields.id} 
                        />
                      )
                    }
                  })}
                </ul>
              }

            </div>
          </div>
        </div>
      </section>
    }
    </>
  );

}

