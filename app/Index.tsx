/// <reference path="../typings/tsd.d.ts" />
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Store, createStore, applyMiddleware} from 'redux';
import { Provider } from 'react-redux';
import { Dogfood } from './Dogfood/dogfood';
import { VSTS } from './VSTS/VSTS';
import { Done } from './Authenticate/done';
import { completeAddInReducer } from './Redux/GlobalReducer';
import thunkMiddleware from 'redux-thunk';
declare const require: (name: String) => any;
import * as promise from 'es6-promise';
promise.polyfill();

interface IHotModule {
  hot?: { accept: (path: string, callback: () => void) => void };
};

declare const module: IHotModule;

function configureStore(): Store {
  const store: Store = createStore(completeAddInReducer, applyMiddleware(
    thunkMiddleware // lets us dispatch() functions
    // neat middleware that logs actions
  ));

  if (module.hot) {
    module.hot.accept('./reducers', () => {
      const nextRootReducer: any = require('./Redux/LoginReducer').completeAddInReducer;
      store.replaceReducer(nextRootReducer);
    });
  }

  return store;
}

const store: Store = configureStore();


class Main extends React.Component<{}, {}> {

  public getRoute(): string {
    let url: string = document.URL;
    let strings: string[] = url.split('/');
    let output: string = strings[3];
    if (output.includes('?')) {
      output = output.slice(0, strings[3].indexOf('?'));
    }
    return output;
  }

  public getDomain(): string {
    let url: string = document.URL;
    let strings: string[] = url.split('/');
    return strings[2];
  }

  public render(): React.ReactElement<Provider> {
    this.addPolyfill();
    if (this.getDomain().indexOf('outlookvsts') !== -1) {
      return (<Dogfood />);
    }
    const route: string = this.getRoute();
    switch (route) {
      case 'dogfood':
        return(<Dogfood />);
      case 'vsts':
        return(<Provider store = {store}><VSTS /></Provider>);
      case 'done':
        return(<Done />);
      default:
        return(<div>Route: '{route}' is not a valid route!</div>);
    }
  }

  private addPolyfill(): void {
    if (!String.prototype.includes) {
      String.prototype.includes = function(): boolean {
        'use strict';
        return String.prototype.indexOf.apply(this, arguments) !== -1;
      };
    }
    if (typeof Object.assign !== 'function') {
      Object.assign = function(target: Object): Object {
        if (target == null) {
          throw new TypeError('Cannot convert undefined or null to object');
        }

        target = Object(target);
        for (let index: number = 1; index < arguments.length; index++) {
          let source: any = arguments[index];
          if (source != null) {
            for (let key in source) {
              if (Object.prototype.hasOwnProperty.call(source, key)) {
                target[key] = source[key];
              }
            }
          }
        }
        return target;
      };
    }
  }
}

ReactDOM.render(<Main />, document.getElementById('app'));
