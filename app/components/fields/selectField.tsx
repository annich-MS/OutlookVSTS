/// <reference path='../../../typings/tsd.d.ts' />

import * as React from 'react';
import { Styles } from './styles';

interface ISelectFieldProps {
  label: string;
  options: string[];
  onChange?: ICallback;
  selected: string;
}

interface IEmptyState {

}

interface ICallback { (option: string): void; }

export class SelectField extends React.Component<ISelectFieldProps, IEmptyState> {

  public onSelect(event: any): void {
    this.props.onChange(event.target.value);
  }

  public render(): React.ReactElement<{}> {
    let items: JSX.Element[] = [<option value=''></option>];

    const selected: string = this.props.selected;
    let i: number = 0;
    this.props.options.forEach(element => {
      items.push(<option id={'option' + i} value={element}>{element}</option>);
      i += 1;
    });

    return (<div><div>
      <label className='ms-font-m'>{this.props.label}</label> <br />
      <select style={Styles.select} onChange={this.onSelect.bind(this) } value={selected}>
          {items}
      </select>
      </div></div>
    );
  }
}
