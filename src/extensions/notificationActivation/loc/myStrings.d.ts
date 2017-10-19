declare interface INotificationActivationCommandSetStrings {
  COMMAND_ACTIVATE: string;
  COMMAND_DEACTIVATE: string;
}

declare module 'NotificationActivationCommandSetStrings' {
  const strings: INotificationActivationCommandSetStrings;
  export = strings;
}
