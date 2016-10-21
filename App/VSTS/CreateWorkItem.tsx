/// <reference path="../../office.d.ts" />
import * as React from 'react';
import { Description } from './Description';
import { Title } from './Title';
import { Save } from './Save';
import { WorkItemDropdown } from './WorkItemDropdown';
import { Classification } from './Classification';
import { Gear } from './Gear';
import { Feedback } from './SimpleComponents/Feedback';
import { Error } from './SimpleComponents/Error';

/**
 * Renders all components of the Create page
 * @class { CreateWorkItem }
 */

export class CreateWorkItem extends React.Component<{}, {}> {
  /**
   * Renders the div that contains all the components of the Create page
   */
  public render(): React.ReactElement<{}> {
    console.log('got to create pg');
    return (
      <div>
        <Error />
        <Gear />
        <WorkItemDropdown/>
        <Title/>
        <Description/>
        <Classification/>
        <Save/>
        <Feedback/>
      </div>
    );
  }
}
