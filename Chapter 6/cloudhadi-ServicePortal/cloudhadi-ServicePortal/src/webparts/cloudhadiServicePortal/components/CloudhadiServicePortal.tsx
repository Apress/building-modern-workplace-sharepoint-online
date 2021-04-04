import React, { useState, useEffect } from 'react';
import styles from './CloudhadiServicePortal.module.scss';
import { CommandBar, ICommandBarItemProps } from 'office-ui-fabric-react/lib/CommandBar';
import CreateRequest from './CreateRequest';
import ViewMyRequests from './ViewMyRequests';
import FAQ from './FAQ';
import LiveChat from './LiveChat';
import AssignedRequests from './AssignedRequests';
import { sp } from "@pnp/sp";
import "@pnp/sp/site-users/web";

function CloudhadiServicePortal() {

  // state variable for setting form.
  const [selectedForm, setSelectedForm] = useState(<CreateRequest />);

  // state variables for group check
  const [serviceExecutive, setServiceExecutive] = useState(false);

  // Set form on command bar menu click.
  const onMenuClick = (form) => {
    setSelectedForm(form);
  }

  // command bar items
  let _items: ICommandBarItemProps[] = [
    {
      key: 'New',
      text: 'New Request',
      iconProps: { iconName: 'Add' },
      onClick: () => onMenuClick(<CreateRequest />)
    },
    {
      key: 'View',
      text: 'View My Requests',
      iconProps: { iconName: 'GroupedList' },
      onClick: () => onMenuClick(<ViewMyRequests />)
    },

    {
      key: 'FAQ',
      text: 'FAQ',
      iconProps: { iconName: 'Questionnaire' },
      onClick: () => onMenuClick(<FAQ />)
    },
    {
      key: 'Chat',
      text: 'Live Chat',

      iconProps: { iconName: 'Chat' },
      onClick: () => onMenuClick(<LiveChat />)
    }

  ];

  // On component mount
  useEffect(() => {
    //Set service executive if logged in user belongs to the group
    checkServiceExecutive();
  }, []);

  // Check if current user belongs to service executive group
  const checkServiceExecutive = async () => {
    let groups: any = await sp.web.currentUser.groups();
    await groups.forEach(group => {
      if (group.LoginName == 'Service Executives') {
        setServiceExecutive(true);
        return;
      }
    }
    );
  }

  // Add 'Assigned to Me' tab if service executive
  if (serviceExecutive) {
    _items.splice(2, 0, {
      key: 'Assigned',
      text: 'Assigned to Me',
      iconProps: { iconName: 'ClipboardList' },
      onClick: () => onMenuClick(<AssignedRequests />)
    });

  }

  return (

    <div>
      <CommandBar
        items={_items}
      />
      <div>
        {selectedForm}
      </div>
    </div>

  );
}

export default CloudhadiServicePortal;



