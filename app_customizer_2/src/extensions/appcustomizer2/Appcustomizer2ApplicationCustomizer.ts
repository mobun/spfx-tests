import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Log } from '@microsoft/sp-core-library';
import { IDynamicDataCallables, IDynamicDataPropertyDefinition } from '@microsoft/sp-dynamic-data';

import * as strings from 'Appcustomizer2ApplicationCustomizerStrings';

const LOG_SOURCE: string = 'Appcustomizer2ApplicationCustomizer';

export interface IAppcustomizer2ApplicationCustomizerProperties {
  testMessage: string;
}


export default class Appcustomizer2ApplicationCustomizer
  extends BaseApplicationCustomizer<IAppcustomizer2ApplicationCustomizerProperties> {
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
      { id: 'title', title: 'Title from AppCustomizer 2' },
      { id: 'random', title: 'Random number (AppCustomizer 2)' }
    ];
  };

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private _getPropertyValue = (propertyId: string): any => {
    switch (propertyId) {
      case 'title':
        return this.properties.testMessage || 'AppCustomizer 2 Title';
      case 'random':
        return Math.random();
      default:
        return undefined;
    }
  };
}
