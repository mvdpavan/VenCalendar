//import React from 'react';
import * as React from 'react';
import styles from './CalendarTemplate.module.scss';

export default function emptyHtml(props) {
    return (<div className= {styles.EmptyCalendar}>
                <div className={styles.container}>
                <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>
                    <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                    <span className="ms-font-xl ms-fontColor-white">{props.title}</span>
                    <p className="ms-font-l ms-fontColor-white">Edit this web part to continue</p>
                    <a href="https://aka.ms/spfx" className={styles.button}>
                        <span className={styles.label}>View More Modern Web Parts</span>
                    </a>
                    </div>
                </div>
                </div>  
            </div>
    );
  }