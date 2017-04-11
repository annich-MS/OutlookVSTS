import * as React from "react";
import { CommandButton } from "office-ui-fabric-react/lib/Button";
import Constants from "../../models/constants";

export class Feedback extends React.Component<{}, {}> {
        public render(): React.ReactElement<any> {
        if (Office.context.mailbox.diagnostics.hostName === Constants.IOS_HOST_NAME) {
            // display new message isn"t available in mobile
            return (<div/>);
        } else {
            return (
                <div style={{textAlign: "center"}}>
                    <CommandButton onClick={this.feedback.bind(this)}>Give Feedback</CommandButton>
                </div>);
        }
    }

    private feedback(): void {
        Office.context.mailbox.displayNewMessageForm({
            subject: "VSTS add-in feedback",
            toRecipients: ["VSTSaddin_fb@microsoft.com"],
        });

    }
}
