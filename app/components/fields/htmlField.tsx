/// <reference path='../../../typings/tsd.d.ts' />

import * as React from 'react';
import { Styles } from './styles';

interface IHtmlFieldProps {
  label: string;
  text: string;
  onChange?: ICallback;
}

interface ICallback { (option: string): void; }

export class HtmlField extends React.Component<IHtmlFieldProps, {} > {

  public onChange(value: any): void {
    this.props.onChange(value);
  }

  public render(): React.ReactElement<{}> {
    return (<div>
                <div>
                    <label className='ms-font-m'>{this.props.label}</label> <br />
                    <textarea style={Styles.body} onChange={this.onChange.bind(this) } value={this.props.text} />
                </div>
            </div>
    );
  }
}
