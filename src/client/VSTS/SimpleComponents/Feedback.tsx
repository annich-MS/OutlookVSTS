import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';

export class Feedback extends React.Component<{}, {}> {
        public render(): React.ReactElement<any> {
        if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
            // display new message isn't available in mobile
            return (<div/>);
        } else {
            return (
                <div style={{textAlign: 'center'}}>
                    <Button buttonType={ButtonType.command} onClick={this.feedback.bind(this)}>Give Feedback</Button>
                </div>);
        }
    }

    private feedback(): void {
        Office.context.mailbox.displayNewMessageForm({
            subject: 'VSTS add-in feedback',
            toRecipients: ['VSTSaddin_fb@microsoft.com'],
        });

    }
}
