import * as React from "react";
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import { MSGraphClient } from '@microsoft/sp-http';

export interface ITeamsFooterProps {
      context: any;
}
function TeamsFooter(props: ITeamsFooterProps) {

      // post message to Teams
      const postMessage = () => {

            const chatMessage = {
                  body: {
                        content: 'Hey Team, an urgent issue notified! Please look out in the service portal'
                  }
            };

            props.context.msGraphClientFactory.getClient().
                  then((client: MSGraphClient): void => {
                        client.api('/teams/ed943053-550e-48d5-b679-a2e8af9820a4/channels/19:1740ec0332f24802b31a31538e7801d4@thread.tacv2/messages')
                              .post(chatMessage)
                              .then(success => {
                                    alert("Message posted")
                              }, error => {
                                    alert("failed")
                              });


                  });
      }

      // Data for CommandBar  
      const getItems = () => {
            return [
                  {
                        key: 'Teams',
                        name: 'Notify Urgent Issue to Service Executives Channel',
                        iconProps: {
                              iconName: 'TeamsLogo'
                        },

                        onClick: () => postMessage()
                  }
            ];
      }

      // command bar
      return (
            <div className={"ms-bgColor-themeDark ms-fontColor-white"} >
                  
                  <CommandBar
                        items={getItems()}
                  />
            </div>
      );
}

export default TeamsFooter