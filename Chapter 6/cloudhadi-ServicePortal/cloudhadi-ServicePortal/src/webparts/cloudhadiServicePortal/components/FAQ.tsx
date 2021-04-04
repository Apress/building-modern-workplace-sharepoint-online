import * as React from 'react';
import styles from './CloudhadiServicePortal.module.scss';
import { TextField, Dropdown, Stack, IStackTokens, PrimaryButton, DefaultButton } from '@fluentui/react';

const stackTokens: IStackTokens = { childrenGap: 30 };

function FAQ() {
    return (
        <div className={styles.cloudhadiServicePortal}>
            <div className={styles.container}>
                <div className={styles.row}>
                    <div className={styles.column}>
                        <span className={styles.title}>Frequently Asked Questions</span>
                        <div id="faqForm">
                        </div>
                    </div>
                </div>
            </div>
        </div>
    );
}

export default FAQ;

