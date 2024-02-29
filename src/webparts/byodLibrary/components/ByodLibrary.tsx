import * as React from 'react';
import styles from './ByodLibrary.module.scss';
import './ByodLibrary.scss';
import { IByodLibraryProps } from './IByodLibraryProps';
import { getGraphMemberOf, isFromTargetAudience, groupBy, getListItemsGraph, isUserManage, deleteItem } from '../services/requests';
// import { escape } from '@microsoft/sp-lodash-subset';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import {  faChevronUp, faChevronDown } from '@fortawesome/free-solid-svg-icons';
import LibraryItem from './LibraryItem/LibraryItem';
import Search from './Search/Search';
import EditControls from './EditControls/EditControls';
import { useBoolean } from '@uifabric/react-hooks';
import { IFrameDialog } from "@pnp/spfx-controls-react/lib/IFrameDialog";
import { DefaultButton, Dialog, DialogFooter, DialogType, PrimaryButton } from 'office-ui-fabric-react';
import Preloader from './Preloader/Preloader';

export default function ByodLibrary(props: IByodLibraryProps) {

  const [showBasedOnTargetAudience, setShowBasedOnTargetAudience] = React.useState(false);
  const [items, setItems] = React.useState([]);
  const [filteredItems, setFilteredItems] = React.useState(items);
  const [isExp, setIsExp] = React.useState(props.displayState && props.isCollapsible);
  const [memberOfGroups, setMemberofGroups] = React.useState(null);
  
  const [categories, setCategories] = React.useState([]);
  const [categItems, setCategItems] = React.useState(null);
  const [filteredCategItems, setFilteredCategItems] = React.useState(categItems);

  const [showEditControls, {toggle: toggleEditControls}] = useBoolean(false);
  const handleEditChange = (ev: React.MouseEvent<HTMLElement>, checked: boolean) =>{
    toggleEditControls();
  };
  const [iframeUrl, setIframeUrl] = React.useState('');
  const [iframeShow, setIframeShow] = React.useState(false);
  const [iframeState, setIframeState] = React.useState('Add');
  const [hideDeleteDialog, {toggle: toggleHideDeleteDialog}] = useBoolean(true);
  const [isDataLoading, { toggle: toggleIsDataLoading }] = useBoolean(false);
  const [libItemId, setLibItemId] = React.useState(null);

  React.useEffect(()=>{
    console.log("props.context", props.context);

    getListItemsGraph(props.context, props.siteUrl, props.listName).then(res => {
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
    getGraphMemberOf(props.context).then((memberOfGroupsRes: any) => {
      console.log("grpahMemberOf", memberOfGroupsRes);
      setMemberofGroups(memberOfGroupsRes);

      if (props.isCollapsible){
        if (props.displayState === 'expanded') setIsExp(true);
        if (props.displayState === 'collapsed') setIsExp(false);
        if (props.displayState === 'expandedTargetAudience'){
          if(props.targetAudience && props.targetAudience.length > 0)
            setIsExp(isFromTargetAudience(props.userEmail, memberOfGroupsRes, props.targetAudience, 'fullName'));
          else setIsExp(true);
        }
      }

      if (props.showBasedOnTargetAudience && props.targetAudience && props.targetAudience.length > 0){
        setShowBasedOnTargetAudience(isFromTargetAudience(props.userEmail, memberOfGroups, props.targetAudience, 'fullName'));
      }else{
        setShowBasedOnTargetAudience(true);
      }

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

    
  }, []);

  const onSearchChanged = (_: any, text: string): void => {
    if (props.groupBy) setFilteredCategItems(items.filter(item => item.Title.toLowerCase().indexOf(text.toLowerCase()) >= 0));
    else setFilteredItems(items.filter(item => item.fields.Title.toLowerCase().indexOf(text.toLowerCase()) >= 0));
  };

  const handleToggleHideDialog = () => {
    setIframeState('Add');
    setIframeUrl(`${props.siteUrl}/lists/${props.listName}/Newform.aspx`);
    setIframeShow(true);
  };
  const handleDeleteDlg = (itemdId: string) => {
    setLibItemId(itemdId);
    toggleHideDeleteDialog();
  };
  const handleEdit = (itemId: string) => {
    setIframeState('Edit');
    setIframeUrl(`${props.siteUrl}/lists/${props.listName}/Editform.aspx?ID=${itemId}`);
    setIframeShow(true);
  };
  const viewAllHandler = () => {
    window.open(`${props.siteUrl}/lists/${props.listName}/Allitems.aspx`, '_blank');
  };
  const deleteItemHandler = () =>{
    toggleIsDataLoading();
    deleteItem(props.context, props.siteUrl, props.listName, libItemId).then(()=>{
      getListItemsGraph(props.context, props.siteUrl, props.listName).then(res => {
        if (props.groupBy){
          const groupedArr = groupBy(res, props.groupByField);
          setCategItems(groupedArr);
          setFilteredCategItems(groupedArr);
          setCategories(Object.keys(groupedArr));
        }
        setItems(res);
        setFilteredItems(res);
        toggleHideDeleteDialog();
      });
    });
  };
  
  const onIFrameDismiss = async (event: React.MouseEvent) => {
    setIframeShow(false);
    toggleIsDataLoading();
    getListItemsGraph(props.context, props.siteUrl, props.listName).then(res => {
      if (props.groupBy){
        const groupedArr = groupBy(res, props.groupByField);
        setCategItems(groupedArr);
        setFilteredCategItems(groupedArr);
        setCategories(Object.keys(groupedArr));
      }
      setItems(res);
      setFilteredItems(res);
      toggleIsDataLoading();
    });
  };
  const onIFrameLoad = async (iframe: any) => {
    let keepOpen: boolean;
    if (iframeState === "Add" || iframeState === "Edit")
      keepOpen = iframe.contentWindow.location.href.indexOf('Newform.aspx') > 0 || iframe.contentWindow.location.href.indexOf('Editform.aspx') > 0;
    else
      keepOpen = iframe.contentWindow.location.href.indexOf('AllItems.aspx') > 0;
    if (!keepOpen) {
      onIFrameDismiss(null);
    }
  };

  return (
    <>
    {showBasedOnTargetAudience &&
      <section className={`${styles.byodLibrary} ${props.hasTeamsContext ? styles.teams : ""}`}>
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
                              showEditControls={showEditControls}
                              handleDelete={handleDeleteDlg}
                              handleEdit={handleEdit}
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
                        isFromTargetAudience(props.userEmail, memberOfGroups, item.fields._ModernAudienceTargetUserField, 'LookupValue')){
                      return(                                            
                        <LibraryItem 
                          item={item.fields} 
                          customImgPicker={props.customImgPicker} 
                          iconPicker={props.iconPicker} 
                          thumbnail={props.thumbnail} 
                          key={item.fields.id} 
                          showEditControls={showEditControls}
                          handleDelete={handleDeleteDlg}
                          handleEdit={handleEdit}
                        />
                      )
                    }
                  })}
                </ul>
              }
              {isUserManage(props.context) &&
                <EditControls
                  toggleHideDialog={handleToggleHideDialog} 
                  handleEditChange={handleEditChange} 
                  viewAllHandler = {viewAllHandler}
                />
              }
            </div>
          </div>
        </div>

        <IFrameDialog 
          url={iframeUrl}
          width={iframeState === "Add" ? '40%' : '70%'}
          height={'90%'}
          hidden={!iframeShow}
          iframeOnLoad={(iframe) => onIFrameLoad(iframe)}
          onDismiss={(event) => onIFrameDismiss(event)}
          allowFullScreen = {true}
          dialogContentProps={{
            type: DialogType.close,
            showCloseButton: true
          }}
        />

        <Dialog
          hidden={hideDeleteDialog}
          onDismiss={toggleHideDeleteDialog}
          dialogContentProps={{type: DialogType.close, title: "Delete Item"}}>
          <p>Are you sure you want to delete this item? </p>
          <Preloader isDataLoading={isDataLoading} />
          <DialogFooter>
              <PrimaryButton onClick={deleteItemHandler} text="Yes" />
              <DefaultButton onClick={toggleHideDeleteDialog} text="No" />
          </DialogFooter>
        </Dialog>

      </section>
    }
    </>
  );

}

