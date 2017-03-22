import * as React from "react";
import { Provider } from "react-redux";

export default class Done extends React.Component<{}, {}> {

  public constructor() {
    super();
    Office.initialize = () => { Office.context.ui.messageParent("done"); };
  }

  public render(): React.ReactElement<Provider> {
    return (<div>You may now close this window.</div>);
  }
}
