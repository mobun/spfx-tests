import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Log } from '@microsoft/sp-core-library';
import { IDynamicDataCallables, IDynamicDataPropertyDefinition } from '@microsoft/sp-dynamic-data';

import * as strings from 'Appcustomizer1ApplicationCustomizerStrings';

const LOG_SOURCE: string = 'Appcustomizer1ApplicationCustomizer';

export interface IAppcustomizer1ApplicationCustomizerProperties {
 
  testMessage: string;
}


export default class Appcustomizer1ApplicationCustomizer
  extends BaseApplicationCustomizer<IAppcustomizer1ApplicationCustomizerProperties> {

  private _provider: IDynamicDataCallables | undefined;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    let message: string = this.properties.testMessage;
    if (!message) {
      message = '(No properties were provided.)';
    }


    this._provider = {
      getPropertyDefinitions: this._getPropertyDefinitions,
      getPropertyValue: this._getPropertyValue
    };
    this.context.dynamicDataSourceManager.initializeSource(this._provider);

    return Promise.resolve();
  }

  public onDispose(): void { /* no-op */ }

  private _getPropertyDefinitions = (): ReadonlyArray<IDynamicDataPropertyDefinition> => {
    return [
      { id: 'message', title: 'Message from AppCustomizer 1' },
      { id: 'timestamp', title: 'Timestamp (AppCustomizer 1)' }
    ];
  };

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private _getPropertyValue = (propertyId: string): any => {
    switch (propertyId) {
      case 'message':
        return this.properties.testMessage || 'Hello from AppCustomizer 1';
      case 'timestamp':
        return new Date().toISOString();
      default:
        return undefined;
    }
  };
}
