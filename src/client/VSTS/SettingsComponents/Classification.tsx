import * as React from "react";
import ConfigurableDropdown from "./configurableDropdown";
import DropdownConfiguration from "./dropdownConfig";
import NavigationStore from "../../stores/navigationStore";
import APTCache from "../../stores/aptCache";

interface IClassificationProps {
  cache: APTCache;
  navigationStore: NavigationStore;
}

/**
 * Renders the Acccount, Project, and Area components
 * @class {Classification}
 */
export class Classification extends React.Component<IClassificationProps, {}> {
  /**
   * Renders the Account, Project, and Area components
   */
  public render(): JSX.Element {
    return (
      <div>
        <ConfigurableDropdown dropdownConfig={DropdownConfiguration.createAccountConfig(this.props.cache)} navigationStore={this.props.navigationStore} />
        <ConfigurableDropdown dropdownConfig={DropdownConfiguration.createProjectConfig(this.props.cache)} navigationStore={this.props.navigationStore} />
        <ConfigurableDropdown dropdownConfig={DropdownConfiguration.createTeamConfig(this.props.cache)} navigationStore={this.props.navigationStore} />
      </div>
    );
  }
}

