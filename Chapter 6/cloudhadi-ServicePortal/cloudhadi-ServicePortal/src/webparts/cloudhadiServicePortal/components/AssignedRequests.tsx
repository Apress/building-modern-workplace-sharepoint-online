import React, { useState, useEffect } from 'react';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { IViewRequestsState } from './IViewRequestsState';
import styles from './CloudhadiServicePortal.module.scss';
import { DetailsList, DetailsListLayoutMode } from 'office-ui-fabric-react/lib/DetailsList';
import { Link, VerticalDivider } from '@fluentui/react';
import CreateRequest from './CreateRequest';
// Column headers.
const columns = [
    { key: 'Title', name: 'Request No.', fieldName: 'Title', minWidth: 70, maxWidth: 200, isResizable: true },
    { key: 'RequestTitle', name: 'Request Title', fieldName: 'RequestTitle', minWidth: 160, maxWidth: 200, isResizable: true },
    { key: 'RequestStatus', name: 'Status', fieldName: 'RequestStatus', minWidth: 70, maxWidth: 200, isResizable: true }
];
// Render list of requests.
function AssignedRequests() {

    // On component mount
    useEffect(() => {
        loadMyRequests();
    }, [])

    // Reset to view requests.
    const resetViewRequest = () => {
        setDoViewRequest(false);
    };

    // state variables for request items.
    const [myItems, setMyItems] = useState([]);

    // state variables for viewing individual request.
    const [doViewRequest, setDoViewRequest] = useState(false);
    const [requestID, setRequestID] = useState(0);

    // Load Service requests
    const loadMyRequests = async () => {
        let currentUser = await sp.web.currentUser();
        await sp.web.lists.getByTitle("Service Portal").items
            .filter(`RequestAssignedTo/EMail eq '${currentUser.Email}' and RequestStatus ne 'Completed'`)
            .select('ID', 'Title', 'RequestTitle', 'RequestStatus')
            .get().then((items) => {
                let result: IViewRequestsState[] = [];
                items.forEach(element => {
                    result.push({
                        ID: element.Id, Title: <Link href="#">{element.Title}</Link>, RequestTitle: element.RequestTitle, RequestStatus: element.RequestStatus
                    });
                });
                return result;
            }).then(resultdata => setMyItems(resultdata));

    };
    // On click of item.
    const _onItemInvoked = (item: any): void => {
        // call child component with ID     
        setRequestID(item.ID);
        setDoViewRequest(true);
    };

    // Load all requests.
    if (!doViewRequest) {

        return (
            <div className={styles.cloudhadiServicePortal}>
                <div className={styles.container}>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <span className={styles.title}>Requests Assigned to me</span>
                            <DetailsList
                                items={myItems}
                                columns={columns}
                                layoutMode={DetailsListLayoutMode.justified}
                                onItemInvoked={_onItemInvoked}
                            />
                        </div>
                    </div>
                </div>
            </div>
        );
    }

    // Call to load individual request.
    else {
        return (
            <CreateRequest ID={requestID} resetView={resetViewRequest} isAssigned={true}/>

        );
    }
}
export default AssignedRequests;
