import * as React from 'react';
import { PreloaderProps } from './PreloaderProps';

import {Spinner, SpinnerSize, Overlay} from '@fluentui/react';

export default function IPreloader (props:PreloaderProps) {

    return(
        <>
            {props.isDataLoading &&
                <>
                    <Overlay />
                    <div>
                        <Spinner size={SpinnerSize.medium} label="Please Wait, Updating Library Items..." ariaLive="assertive" labelPosition="right" />
                    </div>
                </>
            }
        </>
    );
}