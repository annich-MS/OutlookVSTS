import * as React from "react";
import { SignInButton } from "./SignInButton";
import { AddInDescription } from "./AddInDescription";
import NavigationStore from "../../stores/navigationStore";

interface ILogInPageProps {
  // child dependancy
  navigationStore: NavigationStore;
}

/**
 * Dumb component
 * Renders the add-in description and sign in button
 */
export class LogInPage extends React.Component<ILogInPageProps, {}> {

  /**
   * Renders the add-in description and sign in button
   */
  public render(): JSX.Element {
    let imageStyle: any = {
      display: "block",
      margin: "auto",
      maxWidth: "325px",
      width: "100%",
    };


    return (<div>
      <image style={imageStyle} src="../../../public/Images/VSTSLogo_Long.png" />
      <AddInDescription />
      <SignInButton navigationStore={this.props.navigationStore} />
    </div>);
  }
}
