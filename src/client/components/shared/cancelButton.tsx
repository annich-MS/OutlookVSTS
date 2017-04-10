import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import NavigationStore from "../../stores/navigationStore";
import NavigationPage from "../../models/navigationPage";

interface ISettingsProps {
  navigationStore: NavigationStore;
  backTarget: NavigationPage;
}

/**
 * renders the cancel button that redirects to CreateWorkItem
 */
export default class CancelButton extends React.Component<ISettingsProps, any> {

  /**
   * Redirects to previous page
   */
  public Cancel(): void {
    this.props.navigationStore.navigate(this.props.backTarget);
  }

  /**
   * Renders the cancel button
   */
  public render(): JSX.Element {
    return (
       <div style={{float: "right"}}>
          <Button buttonType={ButtonType.command} onClick={this.Cancel.bind(this)}>Back</Button>
      </div>
    );
  }
}
