import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { SPFI } from '@pnp/sp';

import { createSp } from '../../core/http';
import CopyJetSetup from './components/CopyJetSetup';
import type { ICopyJetSetupProps } from './components/ICopyJetSetupProps';

export interface ICopyJetSetupWebPartProps {}

export default class CopyJetSetupWebPart extends BaseClientSideWebPart<ICopyJetSetupWebPartProps> {
  private _sp!: SPFI;

  protected onInit(): Promise<void> {
    this._sp = createSp(this.context);
    return super.onInit();
  }

  public render(): void {
    const user = this.context.pageContext.user;
    const element: React.ReactElement<ICopyJetSetupProps> = React.createElement(CopyJetSetup, {
      sp: this._sp,
      siteTitle: this.context.pageContext.web.title,
      createdBy: user.email || user.loginName
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
