import * as React from 'react';

export class Feedback extends React.Component<{},{}> {
        public render(): React.ReactElement<any> {
        let style: any = {
            background: 'rgb(255,255,255)',
            border: 'rgb(255,255,255)',
            color: 'rgb(0,122,204)',
            font: '15px arial, ms-segoe-ui',
            'text-align': 'center',
        };

        if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
            // display new message isn't available in mobile
            return (<div/>);
        } else {
            return (
                <div style={{margin:'auto', width:'75%', 'text-align':'center'}}>
                    <br/>
                    <button style={style} onClick={this.feedback.bind(this)}>
                        <span font-family='Arial Black, Gadget, sans-serif' > Give Feedback </span>
                    </button>
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