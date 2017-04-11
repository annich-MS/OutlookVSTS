import * as React from "react";
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";

/**
 * Dumb component
 * Renders connecting page
 * @class {Connecting} 
 */
export default class Connecting extends React.Component<{}, {}> {

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
          <Spinner size={SpinnerSize.large} label="Connecting..." />
        </div>
      </div>);
  }
}
