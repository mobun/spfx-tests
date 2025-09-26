import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Log } from '@microsoft/sp-core-library';
import { IDynamicDataSource } from '@microsoft/sp-dynamic-data';

import * as strings from 'Appcustomizer3ApplicationCustomizerStrings';

const LOG_SOURCE: string = 'Appcustomizer3ApplicationCustomizer';

export interface IAppcustomizer3ApplicationCustomizerProperties {

  testMessage: string;
}
export default class Appcustomizer3ApplicationCustomizer
  extends BaseApplicationCustomizer<IAppcustomizer3ApplicationCustomizerProperties> {
  private _onSourcesChanged?: () => void;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    let message: string = this.properties.testMessage;
    if (!message) {
      message = '(No properties were provided.)';
    }

  const sources: ReadonlyArray<IDynamicDataSource> = this.context.dynamicDataProvider.getAvailableSources();
    // eslint-disable-next-line no-console
  console.log('[AppCustomizer3] Available Dynamic Data Sources:', sources.map(s => ({ id: s.id, metadata: s.metadata })));

    const onSourcesChanged = (): void => {
      const updated = this.context.dynamicDataProvider.getAvailableSources();
      // eslint-disable-next-line no-console
      console.log('[AppCustomizer3] Sources changed:', updated.map(s => ({ id: s.id, metadata: s.metadata })));
    };
    this._onSourcesChanged = onSourcesChanged;
    this.context.dynamicDataProvider.registerAvailableSourcesChanged(onSourcesChanged);

    return Promise.resolve();
  }

  public onDispose(): void {
    if (this._onSourcesChanged) {
      this.context.dynamicDataProvider.unregisterAvailableSourcesChanged(this._onSourcesChanged);
      this._onSourcesChanged = undefined;
    }
  }
}
