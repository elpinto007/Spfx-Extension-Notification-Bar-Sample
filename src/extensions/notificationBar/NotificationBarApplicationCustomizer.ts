import styles from "./NotificationBar.module.scss";
import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from "@microsoft/sp-application-base";

import * as strings from "NotificationBarApplicationCustomizerStrings";
import * as pnp from "sp-pnp-js";

const LOG_SOURCE: string = "NotificationBarApplicationCustomizer";
import { FIELD_NOTIFICATIONACTIVE, FIELD_NOTIFICATIONTEXT, FIELD_NOTIFICATIONTYPE } from "./../../shared/Constants";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface INotificationBarApplicationCustomizerProperties { }

/** A Custom Action which can be run during execution of a Client Side Application */
export default class NotificationBarApplicationCustomizer
  extends BaseApplicationCustomizer<INotificationBarApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    pnp.setup({
      spfxContext: this.context
    });

    this.getNotification();

    return Promise.resolve();
  }

  private getNotification = (): void => {

    pnp.sp.web.lists.getByTitle("Notifications")
    .items
    .select(FIELD_NOTIFICATIONTEXT, FIELD_NOTIFICATIONTYPE, FIELD_NOTIFICATIONACTIVE)
    .filter(`${FIELD_NOTIFICATIONACTIVE} eq 1`)
    .get()
    .then((items: INotificationItem[]) => {
        console.log(`Got data: ${items.length}`);
        if(items.length > 0) {
          this.renderNotificationInTopPlaceholder(items[0]);
        }

    }, (error) => {
      console.error(error);
    });
  }

  private renderNotificationInTopPlaceholder = (notification: INotificationItem): void => {
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, { onDispose: this._onDispose });

      if (!this._topPlaceholder) { return console.error("The expected placeholder (Top) was not found."); }

      console.log(notification);
      if (this._topPlaceholder.domElement) {
        this._topPlaceholder.domElement.innerHTML = `
              <div class="${styles.appCustomizer}">
                <div class="${styles.notificationBar} ${this.getNotificationType(notification)}">
                  <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i> &nbsp;${notification.Title}
                </div>
              </div>`;
      }
    }
  }

  private getNotificationType = (notification: INotificationItem): any => {
    switch(notification.NotificationType) {
      case "Important":
        return styles.importantNotification;
      case "Warning":
        return styles.warningNotification;
      default:
        return styles.infoNotification;
    }
  }

  private _onDispose(): void {
    Log.info("App Customizer", "[NotificationBarApplicationCustomizer._onDispose] Disposed custom placeholders.");
  }

  private _topPlaceholder: PlaceholderContent | undefined;

}