import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import Ups from './components/Ups';
import { IUpsProps } from './components/IUpsProps';

export interface IUpsWebPartProps {
}

export default class UpsWebPart extends BaseClientSideWebPart<IUpsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IUpsProps> = React.createElement(
      Ups,
      {
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
