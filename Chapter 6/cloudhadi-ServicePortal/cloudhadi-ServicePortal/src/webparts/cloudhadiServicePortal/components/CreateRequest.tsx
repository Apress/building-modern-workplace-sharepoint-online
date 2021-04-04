import React, { useState, useEffect } from 'react';
import styles from './CloudhadiServicePortal.module.scss';
import { TextField, Dropdown, Stack, IStackTokens, PrimaryButton } from '@fluentui/react';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react';
import { ICreateRequestProps } from './ICreateRequestProps';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IItemAddResult } from "@pnp/sp/items";
import { IViewRequestsState } from './IViewRequestsState';
const stackTokens: IStackTokens = { childrenGap: 30 };
let csrNumber = "CSR" + Math.floor(Math.random() * (99999));

function CreateRequest(props: ICreateRequestProps) {

    // initiate form element and Title.
    let formTitle = "New Service Request";
    let formStructure = <div className=""></div>;

    // state variable for request status & success
    const [reqStatus, setReqStatus] = useState("New");
    const [success, setSuccess] = useState(false);

    // set field values on change - Dropdown fields.
    const handleDropDownChange = (selOption, setFieldValue) => {
        setFieldValue(selOption.text);
    }
    // Create New service request
    if (!props.ID) {
        // state variables for form fields.
        const [reqTitle, setReqTitle] = useState("");
        const [reqDesc, setReqDesc] = useState("");
        const [relatedTo, setRelatedTo] = useState("");


        // Create a Cloudhadi Service Request on submit button click.
        const createCSR = async () => {

            try {
                csrNumber = "CSR" + Math.floor(Math.random() * (99999));

                const iar: IItemAddResult = await sp.web.lists.getByTitle("Service Portal").items.add({
                    Title: csrNumber,
                    RequestTitle: reqTitle,
                    RequestDescription: reqDesc,
                    Relatedto: relatedTo,

                });

                setSuccess(true);
            }
            catch (error) {
                throw (error);
            }


        }

        // set field values on change - Text fields.
        const handleChange = (event, setFieldValue) => {
            setFieldValue(event.target.value);
        }


        formStructure = <div className={styles.formGridRow}>

            <TextField label="Request Title" onChange={(event) => handleChange(event, setReqTitle)} required></TextField>
            <TextField label="Request Description" onChange={(event) => handleChange(event, setReqDesc)} multiline rows={4} required></TextField>
            <Dropdown
                placeholder="Select an option"
                label="Related to"
                options={[
                    { key: 'Access', text: 'Access' },
                    { key: 'Materials', text: 'Materials' },
                    { key: 'Equipments', text: 'Equipments' },
                    { key: 'General', text: 'General' }

                ]}
                onChanged={(selOption) => { handleDropDownChange(selOption, setRelatedTo); }}
                required
            />
            <Stack horizontal tokens={stackTokens} className={styles.buttonStack}>
                <PrimaryButton className={styles.button} text="Submit" onClick={() => createCSR()} />
            </Stack>
            {success === true &&
                <MessageBar messageBarType={MessageBarType.success} isMultiline={false} >Successfully created Service Request. Reference no:{csrNumber}</MessageBar>
            }
        </div>;
    }

    //Display service request
    else {

        // On component mount
        useEffect(() => {
            loadRequest(props.ID);
        }, [])
        // state variable for Item
        const [currentItem, setCurrentItem] = useState<IViewRequestsState>({
            ID: 0, Title: "", RequestTitle: "", RelatedTo: "",
            RequestStatus: "", RequestAssignedTo: ""
        });

        // Load current item
        const loadRequest = async (reqID) => {

            await sp.web.lists.getByTitle("Service Portal").items
                .getById(reqID)
                .select('ID', 'Title', 'RequestTitle', 'RequestDescription', 'Relatedto', 'RequestStatus', 'RequestAssignedTo/EMail')
                .expand('RequestAssignedTo')
                .get().then((item: any) => {
                    let result: IViewRequestsState = {
                        ID: item.Id, Title: item.Title, RequestTitle: item.RequestTitle, RequestDesc: item.RequestDescription, RelatedTo: item.Relatedto,
                        RequestStatus: item.RequestStatus, RequestAssignedTo: (typeof item.RequestAssignedTo !== "undefined") ? item.RequestAssignedTo.EMail : ""
                    }
                    return result;
                }).then(resultdata => setCurrentItem(resultdata));

        };

        formTitle = `Service Request ${currentItem.Title}`;
        if (!props.isAssigned) {
            formStructure = <div className={styles.formGridRow}>

                <TextField label="Request Title" disabled value={currentItem.RequestTitle} ></TextField>
                <TextField label="Request Description" disabled value={currentItem.RequestDesc} multiline rows={4} ></TextField>
                <TextField label="Related to" disabled value={currentItem.RelatedTo} ></TextField>
                <TextField label="Request Assigned To" disabled value={currentItem.RequestAssignedTo} ></TextField>
                <TextField label="Request Status" disabled value={currentItem.RequestStatus} ></TextField>
                <Stack horizontal tokens={stackTokens} className={styles.buttonStack}>
                    <PrimaryButton className={styles.button} text="Back to My Requests" onClick={props.resetView} />
                </Stack>

            </div>
        }
        else {

            // Update status of  current item
            const updateRequestStatus = async () => {
                try {
                    await sp.web.lists.getByTitle("Service Portal").items
                        .getById(currentItem.ID)
                        .update({
                            RequestStatus: reqStatus
                        });
                    currentItem.RequestStatus = reqStatus;
                    setSuccess(true);
                }

                catch (error) {
                    throw (error);
                }

            };

            formStructure = <div className={styles.formGridRow}>
                <TextField label="Request Title" disabled value={currentItem.RequestTitle} ></TextField>
                <TextField label="Request Description" disabled value={currentItem.RequestDesc} multiline rows={4} ></TextField>
                <TextField label="Related to" disabled value={currentItem.RelatedTo} ></TextField>
                <TextField label="Request Assigned To" disabled value={currentItem.RequestAssignedTo} ></TextField>
                <Dropdown
                    defaultSelectedKey={currentItem.RequestStatus}
                    placeholder="Select an option"
                    label="Request Status"
                    options={[
                        { key: 'New', text: 'New' },
                        { key: 'In Progress', text: 'In Progress' },
                        { key: 'Completed', text: 'Completed' },
                        { key: 'Rejected', text: 'Rejected' }

                    ]}
                    onChanged={(selOption) => { handleDropDownChange(selOption, setReqStatus); }}
                    required
                />
                <Stack horizontal tokens={stackTokens} className={styles.buttonStack}>
                     {currentItem.RequestStatus != "Completed"  && <PrimaryButton className={styles.button} text="Update" onClick={() => updateRequestStatus()}/>}
                    <PrimaryButton className={styles.button} text="Back to Assigned Requests" onClick={props.resetView} />
                </Stack>
                {success === true &&
                    <MessageBar messageBarType={MessageBarType.success} isMultiline={false} >Successfully updated Service Request :{currentItem.Title}</MessageBar>
                }
            </div>
        }
    }

    //render form.
    return (
        <div className={styles.cloudhadiServicePortal}>
            <div className={styles.container}>
                <div className={styles.row}>
                    <div className={styles.column}>
                        <span className={styles.title}>{formTitle}</span>
                        <div id="requestForm">
                            <div className={styles.formGrid}>
                                {formStructure}
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    );
}


export default CreateRequest;

