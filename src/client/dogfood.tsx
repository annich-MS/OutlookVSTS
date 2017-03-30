import * as React from "react";

export class Dogfood extends React.Component<{}, {}> {

  public render(): React.ReactElement<any> {
    return (<div className="ms-font-m">This version of the VSTS add-in has been removed.<br />
            Please uninstall this manifest and install the newer, better version
              <a href="https://raw.githubusercontent.com/annich-MS/OutlookVSTS/master/Manifests/releaseManifest.xml">here</a></div>);
  }
}
