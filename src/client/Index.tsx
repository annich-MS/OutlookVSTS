import * as React from "react";
import * as ReactDOM from "react-dom";
import { Dogfood } from "./dogfood";
import { VSTS } from "./VSTS/VSTS";
import Done from "./done";
import { aptCache } from "./stores/aptCache";
import { navigationStore } from "./stores/navigationStore";
import { workItemStore } from "./stores/workItemStore";

class Main extends React.Component<{}, {}> {

  public render(): JSX.Element {
    const route: string = this.getRoute();
    switch (route) {
      case "dogfood":
        return (<Dogfood />);
      case "vsts":
        return (<VSTS aptCache={aptCache} navigationStore={navigationStore} workItemStore={workItemStore} />);
      case "done":
        return (<Done />);
      default:
        return (<div>Route: "{route}" is not a valid route!</div>);
    }
  }

  private getRoute(): string {
    let url: string = document.URL;
    let strings: string[] = url.split("/");
    let output: string = strings[3];
    if (output.includes("?")) {
      output = output.slice(0, strings[3].indexOf("?"));
    }
    return output;
  }

}

ReactDOM.render(<Main />, document.getElementById("app"));
