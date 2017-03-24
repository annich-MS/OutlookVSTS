import * as React from "react";
import {Spinner, SpinnerType, Overlay} from "office-ui-fabric-react";

/**
 * Dumb component
 * Renders saving overlay
 * @class {Saving} 
 */
export class Saving extends React.Component<{}, {}> {

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
          <Spinner type={ SpinnerType.large }/>
        </div>
      </Overlay>);
  }
}




