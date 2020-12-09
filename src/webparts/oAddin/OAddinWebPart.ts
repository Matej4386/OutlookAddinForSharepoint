/// <reference path="../../../node_modules/@types/office-js/index.d.ts" />
import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as $ from 'jquery';
import {sp} from '@pnp/sp';
import PnPTelemetry from '@pnp/telemetry-js';

import * as strings from 'OAddinWebPartStrings';
import OAddin from './components/OAddin';
import { IOAddinProps } from './components/IOAddinProps';

export interface IOAddinWebPartProps {

}

export default class OAddinWebPart extends BaseClientSideWebPart<IOAddinWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      $('#workbenchPageContent').prop('style', 'max-width: none');
      $('.SPCanvas-canvas').prop('style', 'max-width: none');
      $('.CanvasZone').prop('style', 'max-width: none');
      sp.setup(
        {
          spfxContext: this.context,
          sp: {
            headers: {
              Accept: 'application/json; odata=nometadata'
            }
          }
        }
      );
      // Disable PnP Telemetry
      const telemetry: PnPTelemetry = PnPTelemetry.getInstance();
      if (telemetry.optOut) { telemetry.optOut(); }
    });
  }
  public render(): void {
    const element: React.ReactElement<IOAddinProps> = React.createElement(
      OAddin,
      {
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            
          ]
        }
      ]
    };
  }
}
