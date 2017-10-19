import { Item } from "sp-pnp-js/lib/sharepoint";
import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
// tslint:disable-next-line:max-line-length
import { BaseListViewCommandSet, Command, IListViewCommandSetListViewUpdatedParameters, IListViewCommandSetExecuteEventParameters } from "@microsoft/sp-listview-extensibility";
import * as pnp from "sp-pnp-js";

import * as strings from "NotificationActivationCommandSetStrings";
import { FIELD_NOTIFICATIONACTIVE, FIELD_NOTIFICATIONID, FIELD_NOTIFICATIONTEXT, FIELD_NOTIFICATIONTYPE } from "./../../shared/Constants";
import { List } from "sp-pnp-js";



/**
 * if your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * you can define an interface to describe it.
 */
export interface INotificationActivationCommandSetProperties {
  // this is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = "NotificationActivationCommandSet";

export default class NotificationActivationCommandSet extends BaseListViewCommandSet<INotificationActivationCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized NotificationActivationCommandSet");

    pnp.setup({
      spfxContext: this.context
    });

    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    this.context._commands.forEach((command) => {
      if(event.selectedRows.length !== 1) {
        command.visible = false;
      } else {
        if(command.id === "COMMAND_ACTIVATE") {
          command.visible = event.selectedRows[0].getValueByName(`${FIELD_NOTIFICATIONACTIVE}.value`) === "0";
        } else if(command.id === "COMMAND_DEACTIVATE") {
          command.visible = event.selectedRows[0].getValueByName(`${FIELD_NOTIFICATIONACTIVE}.value`) === "1";
        }
      }
    });
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    var itemId: number = parseFloat(event.selectedRows[0].getValueByName("ID"));
    switch (event.itemId) {
      case "COMMAND_ACTIVATE":
        this.updateNotificationStatus(itemId, true);
        break;
      case "COMMAND_DEACTIVATE":
        this.updateNotificationStatus(itemId, false);
        break;
      default:
        throw new Error("Unknown command");
    }
  }

  private async updateNotificationStatus(itemId: number, activate: boolean): Promise<any> {
    var list: List = pnp.sp.web.lists.getByTitle("Notifications");

    if(activate) {
      list.items.select(FIELD_NOTIFICATIONID, FIELD_NOTIFICATIONTEXT, FIELD_NOTIFICATIONTYPE, FIELD_NOTIFICATIONACTIVE)
      .filter(`${FIELD_NOTIFICATIONACTIVE} eq 1`)
      .get().then((items: INotificationItem[]) => {

        if (items.length > 0) {
          items.forEach(async (item: INotificationItem) => {
              await pnp.sp.web.lists.getByTitle("Notifications").items.getById(item.ID).update({
                  Active: false,
              });
            });
          }
      });
    }

    var listItem: Item = list.items.getById(itemId);
    await listItem.update({ Active : activate });
    window.location.reload();
  }
}
