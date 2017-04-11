import * as React from "react";


export default class PageTitle extends React.Component<{}, {}> {
    public render(): JSX.Element {
        return (<span className="ms-font-l">{this.props.children}</span>);
    }
}