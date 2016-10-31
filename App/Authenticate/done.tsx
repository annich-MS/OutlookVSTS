/// <reference path="../../office.d.ts" />
import * as React from 'react';
import { Provider } from 'react-redux';

export class Done extends React.Component<{}, {}> {

  public componentDidMount(): void {
    window.close();
  }

  public render(): React.ReactElement<Provider> {
    return (<div>You may now close this window.</div>);
  }
 }
