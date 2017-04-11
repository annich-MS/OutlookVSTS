import * as React from "react";
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";
import { Overlay } from "office-ui-fabric-react/lib/Overlay";

/**
 * Dumb component
 * Renders saving overlay
 * @class {Saving} 
 */
export default class Saving extends React.Component<{}, {}> {

  /**
   * Renders saving overlay
   */
  public render(): JSX.Element {
    let divStyle: any = {
      alignItems: "center",
      display: "flex",
      height: "100%",
      justifyContent: "center",
    };
    return (
      <Overlay isDarkThemed={true}>
        <div style={divStyle}>
          <Spinner size={SpinnerSize.large} />
        </div>
      </Overlay>);
  }
}




