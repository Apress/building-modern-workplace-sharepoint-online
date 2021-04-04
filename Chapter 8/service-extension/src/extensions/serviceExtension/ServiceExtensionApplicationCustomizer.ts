import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import * as strings from 'ServiceExtensionApplicationCustomizerStrings';
import * as React from "react";
import * as ReactDOM from "react-dom";
import TeamsFooter, { ITeamsFooterProps } from "./TeamsFooter";

const LOG_SOURCE: string = 'ServiceExtensionApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IServiceExtensionApplicationCustomizerProperties {
  // This is an example; replace with your own property
  Bottom: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class ServiceExtensionApplicationCustomizer
  extends BaseApplicationCustomizer<IServiceExtensionApplicationCustomizerProperties> {


  private _bottomPlaceHolder: PlaceholderContent | undefined;
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // Added to handle possible changes on the existence of placeholders.  
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

    // Call render method for generating the HTML elements.  
    this._renderPlaceHolders();

    return Promise.resolve();
  }

  private _renderPlaceHolders(): void {
   
    // Handling the bottom placeholder  
    if (!this._bottomPlaceHolder) {
      this._bottomPlaceHolder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Bottom,
          { onDispose: this._onDispose });

      // The extension should not assume that the expected placeholder is available.  
      if (!this._bottomPlaceHolder) {
        console.error('The expected placeholder was not found.');
        return;
      }

      const elem: React.ReactElement<ITeamsFooterProps> = React.createElement(
         TeamsFooter,
        {      
          context:this.context
        }
      );
      ReactDOM.render(elem, this._bottomPlaceHolder.domElement);
    }
  }

  private _onDispose(): void {
    console.log('[ReactHeaderFooterApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
}
