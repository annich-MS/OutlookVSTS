/// <reference path="../../office.d.ts" />
import * as React from 'react';
import { Provider } from 'react-redux';

export class CreateWorkItem extends React.Component<{}, {}> {
  public render(): React.ReactElement<Provider> {
    console.log('got to vsts');
    return (<div>
              <h1>Dogfood</h1>
            </div>);
  }
 }