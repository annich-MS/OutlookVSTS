import * as React from "react";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { observer } from "mobx-react";

import APTCache from "../../stores/aptCache";

/**
 * Represents the Title Properties
 */
interface IConfigNameProps {
  cache: APTCache;
}

@observer
export default class ConfigName extends React.Component<IConfigNameProps, {}> {
  /**
   * Dipatches the action to change the value of title in the store 
   */
  public handleChange(value: string): void {
    this.props.cache.setName(value);
  }
  /**
   * Rendersthe Title heading and the Title textbox
   */
  public render(): JSX.Element {
    return (
      <div>
        <TextField
          label="Configuration name"
          onChanged={this.handleChange.bind(this)}
          value={this.props.cache.name} />
      </div>
    );
  }
}