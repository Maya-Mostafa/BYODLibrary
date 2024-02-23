import * as React from 'react';
import styles from '../ByodLibrary.module.scss';
import { SearchProps } from './SearchProps';
import { SearchBox } from 'office-ui-fabric-react';


export default function Search(props: SearchProps){

    return(
        <SearchBox 
            placeholder={props.searchPlaceholder} 
            onChange={props.onSearchChanged}
            className={styles.searchBox} 
        />
    );
}