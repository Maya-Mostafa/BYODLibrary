import * as React from 'react';
import styles from '../ByodLibrary.module.scss';
import { LibraryItemProps } from './LibraryItemProps';
import { Icon, TeachingBubble } from 'office-ui-fabric-react';
import FlagBtn from '../FlagBtn/FlagBtn';
import { copyTextToClipboard } from '../../services/requests';
import { useBoolean, useId } from '@fluentui/react-hooks';
import { faUser, faLock, faCircleInfo } from '@fortawesome/free-solid-svg-icons';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';

export default function LibraryItem(props: LibraryItemProps) {

    console.log("LibraryItem props", props);

    const buttonId = useId('targetButton');
    const [teachingBubbleVisible, { toggle: toggleTeachingBubbleVisible }] = useBoolean(false);
    const copyToClipboardHandler = (text: string) => {
        copyTextToClipboard(text);
    };

    return (
        <li className={styles.cardItem} key={props.item.ID}>
            <div className={styles.card}>
                <div className={styles.cardImage}>
                    {props.thumbnail === 'auto' &&
                        <img src={props.item.Image? props.item.Image.Url : require('../../assets/lib5.svg')} />
                    // <img src={props.item.image ? JSON.parse(props.item.image)['serverRelativeUrl'] : require('../../assets/lib5.svg')} />
                    }
                    {props.thumbnail === 'icon' &&
                        <Icon iconName={props.iconPicker}/>
                    }
                    {props.thumbnail === 'customImg' &&
                        <img height={45} src={props.customImgPicker.fileAbsoluteUrl} />
                    }
                </div>
                <div className={styles.cardContent}>
                    <h2 className={styles.cardTitle}>
                        <a title={props.item.link ? props.item.link.Description : ''} href={props.item.link ? props.item.link.Url: ''} 
                            rel="noreferrer" target={props.item.NewTab ? "_blank" : "_self"} data-interception="off">
                            {props.item.Title}
                        </a>
                    </h2>
                    <div className={styles.cardText}>
                        <p>{props.item.Short_x0020_Description}</p>
                        {props.item.login && props.item.pwd &&
                            <div className={styles.cardFlag}>
                                <FlagBtn 
                                    icon={faUser} 
                                    tooltipText='Click to copy username'
                                    calloutText='Copied'
                                    onClick={()=>copyToClipboardHandler(props.item.login)}>
                                    {props.item.login}
                                    </FlagBtn>
                                <FlagBtn 
                                    icon={faLock} 
                                    tooltipText='Click to copy password'
                                    calloutText='Copied'
                                    onClick={()=>copyToClipboardHandler(props.item.pwd)}>
                                    {props.item.pwd}
                                </FlagBtn>
                            </div>
                        }
                        {props.item.LoginDisclaimer &&
                            <>
                                <div className={styles.cardFlag}>
                                    <span className={styles.flagItem} id={buttonId} onClick={toggleTeachingBubbleVisible}>
                                    <FontAwesomeIcon icon={faCircleInfo} />Login Info
                                    </span>
                                </div>

                                {teachingBubbleVisible && (
                                    <TeachingBubble
                                        illustrationImage={{src: require('../../assets/login_info_8.png'), alt: '', height: '110px', style:{paddingLeft: '7px'}}}
                                        isWide={true}
                                        hasSmallHeadline={true}
                                        hasCloseButton={true}
                                        closeButtonAriaLabel="Close"
                                        target={`#${buttonId}`}
                                        onDismiss={toggleTeachingBubbleVisible}
                                        headline="Login Information">
                                        {props.item.LoginDisclaimer}
                                    </TeachingBubble>
                                )}

                            </>
                        }
                    </div>
                </div>
            </div>
        </li>
    );
}
