import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { SPFI } from '@pnp/sp';

import { createSp } from '../../core/http';
import CopyJetInstall from './components/CopyJetInstall';
import type { ICopyJetInstallProps } from './components/ICopyJetInstallProps';

export interface ICopyJetInstallWebPartProps {}

export default class CopyJetInstallWebPart extends BaseClientSideWebPart<ICopyJetInstallWebPartProps> {
  private _sp!: SPFI;

  protected onInit(): Promise<void> {
    this._sp = createSp(this.context);
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<ICopyJetInstallProps> = React.createElement(CopyJetInstall, {
      sp: this._sp,
      siteTitle: this.context.pageContext.web.title,
      siteUrl: this.context.pageContext.web.absoluteUrl
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
