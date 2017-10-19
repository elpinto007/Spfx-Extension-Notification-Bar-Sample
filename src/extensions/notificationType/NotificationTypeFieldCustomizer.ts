import { Log } from "@microsoft/sp-core-library";
import { override } from "@microsoft/decorators";
import { BaseFieldCustomizer, IFieldCustomizerCellEventParameters } from "@microsoft/sp-listview-extensibility";

import * as strings from "NotificationTypeFieldCustomizerStrings";
import styles from "./NotificationTypeFieldCustomizer.module.scss";

/**
 * If your field customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface INotificationTypeFieldCustomizerProperties {
  // this is an example; replace with your own property
  sampleText?: string;
}

const LOG_SOURCE: string = "NotificationTypeFieldCustomizer";

export default class NotificationTypeFieldCustomizer
  extends BaseFieldCustomizer<INotificationTypeFieldCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    // add your custom initialization to this method.  The framework will wait
    // for the returned promise to resolve before firing any BaseFieldCustomizer events.
    Log.info(LOG_SOURCE, "Activated NotificationTypeFieldCustomizer with properties:");
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    Log.info(LOG_SOURCE, `The following string should be equal: "NotificationTypeFieldCustomizer" and "${strings.Title}"`);
    return Promise.resolve();
  }

  @override
  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    
    event.domElement.innerHTML = `<span>${event.fieldValue}</span>`;
    event.domElement.classList.add(styles.NotificationType);

    switch(event.fieldValue) {
      case "Important":
        event.domElement.classList.add(styles.importantCell);
        break;
      case "Warning":
        event.domElement.classList.add(styles.warningCell);
        break;
      default:
        event.domElement.classList.add(styles.infoCell);
        break;
    }
  }

  @override
  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    // this method should be used to free any resources that were allocated during rendering.
    // for example, if your onRenderCell() called ReactDOM.render(), then you should
    // call ReactDOM.unmountComponentAtNode() here.
    super.onDisposeCell(event);
  }
}
