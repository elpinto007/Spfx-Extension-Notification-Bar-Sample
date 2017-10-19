## Spfx Extension Notification Bar Sample

This is a simple sample of a Notification Bar solution. It consists of all three types of extensions, working together. 
- A Notification Bar (Application Customizer extension) that loads a notification from a Notifications List. 
- A Notication Activation (Command Set extension), which can activate or deactivate Notifications on the Notification list. 
- A Notification Type column formatter (Field Customizer extension), which renders the Notification type as a color.

![Alt text](/sample_screenshot.jpg?raw=true "Screenshot of Notification Bar sample")

### Prerequisites

Install all dependencies
```bash
git clone the repo
npm i -g npm@3.x gulp yeoman @microsoft/generator-sharepoint
npm i
```

Use VSCode or any editor of your choice.

### Creating the necessary SharePoint objects

- Create a modern Teamsite on your tenant
- Create a generic list on the Teamsite
- Add two columns: NotificationType (Choice - with options: Info, Warning, Important), Active (Yes/No)
- Add some dummy items, set one to active = yes.

### Run solution in debug mode

```bash
gulp serve --nobrowser
```
Use the --nobrowser attribute because SPFx Extensions cannot be debugged using the local workbench.

Navigate to the Notifications list and paste the following querystring behind the list url (remove all other querystrings first)

```
?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/manifests.js&customActions={"d3359529-056c-4575-aab6-ba589bc70dd2":{"location":"ClientSideExtension.ApplicationCustomizer","properties":{}}, "00ba0713-ed47-40f2-951c-d9b7dc9160c8":{"location":"ClientSideExtension.ListViewCommandSet"}}&fieldCustomizers={"NotificationType":{"id":"ce376f8e-1bc2-467a-baa8-3b76783edabd","properties":{}}}
```

Click the button to 'load debug scripts' on the 'Allow debug scripts' popup.

Now you should be seeing the Notification Bar.
