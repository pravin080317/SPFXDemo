import * as React from "react";
import * as ReactDOM from "react-dom";

import { Log } from "@microsoft/sp-core-library";
import { BaseFormCustomizer } from "@microsoft/sp-listview-extensibility";

import ListFormCustomizer, {
  IListFormCustomizerProps,
} from "./components/ListFormCustomizer";

import { IListFormInterfaceModel } from "./IListFormInterfaceModel";

import { FormDisplayMode } from "@microsoft/sp-core-library";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import styles from "./components/ListFormCustomizer.module.scss";
/**
 * If your form customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IListFormCustomizerFormCustomizerProperties {
  // This is an example; replace with your own property
  sampleText?: string;
}

const LOG_SOURCE: string = "ListFormCustomizerFormCustomizer";

export default class ListFormCustomizerFormCustomizer extends BaseFormCustomizer<IListFormCustomizerFormCustomizerProperties> {
 
  // Added for the item to show in the form; use with edit and view form
  private _item: IListFormInterfaceModel;
  this_item = {};
  // Added for item's etag to ensure integrity of the update; used with edit form
  private _etag?: string;

  public onInit(): Promise<void> {
    if (this.displayMode === FormDisplayMode.New) {
      // we're creating a new item so nothing to load
      return Promise.resolve();
    }

    // load item to display on the form
    return this.context.spHttpClient
      .get(
        this.context.pageContext.web.absoluteUrl +
          `/_api/web/lists/getbytitle('${this.context.list.title}')/items(${this.context.itemId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            accept: "application/json;odata.metadata=none",
          },
        }
      )
      .then((res) => {
        if (res.ok) {
          // store etag in case we'll need to update the item
          this._etag = res.headers.get("ETag");
          return res.json();
        } else {
          return Promise.reject(res.statusText);
        }
      })
      .then((item) => {
        this._item = item;
        return Promise.resolve();
      });
  }

  public formSaved(): void {
    this.formSaved();
  }

  
  public formClosed(): void {
    this.formClosed();
  }

  public render(): void {
    // Use this method to perform your custom rendering.
    
    const listFormCustomizer: React.ReactElement<{}> = React.createElement(
      ListFormCustomizer,
      {
        context: this.context,
        displayMode: this.displayMode,
        formSaved:this.formSaved,
        domElement:this.domElement,
        formClosed:this.formClosed,
        // onSave: this._onSave,
        // onClose: this._onClose,
        items: this._item,
      } as IListFormCustomizerProps
    );

    ReactDOM.render(listFormCustomizer, this.domElement);
  }

  public onDispose(): void {
    // This method should be used to free any resources that were allocated during rendering.
    ReactDOM.unmountComponentAtNode(this.domElement);
    super.onDispose();
  }

}
