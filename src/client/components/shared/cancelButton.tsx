import * as React from "react";
import { Link } from "office-ui-fabric-react/lib/Link";
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
      <span><Link onClick={this.back.bind(this)}><i className="ms-Icon ms-Icon--Back" /> Back</Link></span>
    );
  }
    private back(): void {
        this.props.navigationStore.navigate(this.props.backTarget);
    }
}
