import * as React from 'react';
import styles from './CloudhadiServicePortal.module.scss';
import { TextField, Dropdown, Stack, IStackTokens, PrimaryButton, DefaultButton } from '@fluentui/react';

const stackTokens: IStackTokens = { childrenGap: 30 };

function LiveChat() {
    return (
        <div className={styles.cloudhadiServicePortal}>
            <div className={styles.container}>
                <div className={styles.row}>
                    <div className={styles.column}>
                        <span className={styles.title}>Chat with Virtual Agent</span>
                        <div id="chatForm">

                        </div>
                    </div>
                </div>
            </div>
        </div>
    );
}

export default LiveChat;

