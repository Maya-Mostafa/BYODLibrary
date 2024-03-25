import * as React from 'react';
import styles from '../ByodLibrary.module.scss';
import {EditControlsProps} from './EditControlsProps';

import {CommandBarButton, IIconProps, Toggle} from '@fluentui/react';
import { initializeIcons } from '@uifabric/icons';


export default function EditControls (props: EditControlsProps) {
  
  initializeIcons();
  const addIcon: IIconProps = { iconName: 'CalculatorAddition' };
  const viewAllIcon: IIconProps = { iconName: 'Documentation'};
  const sortIcon: IIconProps = { iconName: 'Sort'};

  return (
    <div className={styles.editControls}>
        <CommandBarButton className={styles.controlsBtn} iconProps={addIcon} text="Add Item" onClick={props.toggleHideDialog} />
        <CommandBarButton className={styles.controlsBtn} iconProps={viewAllIcon} text="View All" onClick={props.viewAllHandler} />
        <CommandBarButton className={styles.controlsBtn} iconProps={sortIcon} text="Reorder" onClick={props.orderHandler} />
        <Toggle className={styles.editToggle} label="Edit Library" inlineLabel onChange={props.handleEditChange} />
    </div>
  );
}