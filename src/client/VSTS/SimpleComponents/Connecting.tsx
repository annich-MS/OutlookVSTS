import * as React from "react";
import { Spinner, SpinnerType } from "office-ui-fabric-react";

/**
 * Dumb component
 * Renders connecting page
 * @class {Connecting} 
 */
export class Connecting extends React.Component<{}, {}> {

  /**
   * Renders Connecting page
   */
  public render(): JSX.Element {
    let overlayStyle: any = {
      bottom: 0,
      display: "block",
      left: 0,
      position: "absolute",
      right: 0,
      top: 0,
    };
    let divStyle: any = {
      alignItems: "center",
      display: "flex",
      height: "100%",
      justifyContent: "center",
    };
    return (
      <div style={overlayStyle}>
        <div style={divStyle}>
          <Spinner type={SpinnerType.large} label="Connecting..." />
        </div>
      </div>);
  }
}




