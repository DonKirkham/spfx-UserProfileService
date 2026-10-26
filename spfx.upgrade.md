# Upgrade project user-profile-service-client-side-solution to v1.23.2

Date: 10/26/2026

## Findings

Following is the list of steps required to upgrade your project to SharePoint Framework version 1.23.2. [Summary](#Summary) of the modifications is included at the end of the report.

### FN001001 @microsoft/sp-core-library | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-core-library

Execute the following command:

```sh
npm i -SE @microsoft/sp-core-library@1.23.2
```

File: [./package.json:17:5](./package.json)

### FN001002 @microsoft/sp-lodash-subset | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-lodash-subset

Execute the following command:

```sh
npm i -SE @microsoft/sp-lodash-subset@1.23.2
```

File: [./package.json:18:5](./package.json)

### FN001003 @microsoft/sp-office-ui-fabric-core | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-office-ui-fabric-core

Execute the following command:

```sh
npm i -SE @microsoft/sp-office-ui-fabric-core@1.23.2
```

File: [./package.json:19:5](./package.json)

### FN001004 @microsoft/sp-webpart-base | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-webpart-base

Execute the following command:

```sh
npm i -SE @microsoft/sp-webpart-base@1.23.2
```

File: [./package.json:21:5](./package.json)

### FN001021 @microsoft/sp-property-pane | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-property-pane

Execute the following command:

```sh
npm i -SE @microsoft/sp-property-pane@1.23.2
```

File: [./package.json:20:5](./package.json)

### FN001023 @microsoft/sp-component-base | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-component-base

Execute the following command:

```sh
npm i -SE @microsoft/sp-component-base@1.23.2
```

File: [./package.json:16:5](./package.json)

### FN001034 @microsoft/sp-adaptive-card-extension-base | Optional

Install SharePoint Framework dependency package @microsoft/sp-adaptive-card-extension-base

Execute the following command:

```sh
npm i -SE @microsoft/sp-adaptive-card-extension-base@1.23.2
```

File: [./package.json:14:3](./package.json)

### FN002002 @microsoft/sp-module-interfaces | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-module-interfaces

Execute the following command:

```sh
npm i -DE @microsoft/sp-module-interfaces@1.23.2
```

File: [./package.json:32:5](./package.json)

### FN002022 @microsoft/eslint-plugin-spfx | Required

Upgrade SharePoint Framework dev dependency package @microsoft/eslint-plugin-spfx

Execute the following command:

```sh
npm i -DE @microsoft/eslint-plugin-spfx@1.23.2
```

File: [./package.json:31:5](./package.json)

### FN002023 @microsoft/eslint-config-spfx | Required

Upgrade SharePoint Framework dev dependency package @microsoft/eslint-config-spfx

Execute the following command:

```sh
npm i -DE @microsoft/eslint-config-spfx@1.23.2
```

File: [./package.json:30:5](./package.json)

### FN002030 @microsoft/spfx-web-build-rig | Required

Upgrade SharePoint Framework dev dependency package @microsoft/spfx-web-build-rig

Execute the following command:

```sh
npm i -DE @microsoft/spfx-web-build-rig@1.23.2
```

File: [./package.json:34:5](./package.json)

### FN002034 @microsoft/spfx-heft-plugins | Required

Upgrade SharePoint Framework dev dependency package @microsoft/spfx-heft-plugins

Execute the following command:

```sh
npm i -DE @microsoft/spfx-heft-plugins@1.23.2
```

File: [./package.json:33:5](./package.json)

### FN010001 .yo-rc.json version | Recommended

Update version in .yo-rc.json

```json
{
  "@microsoft/generator-sharepoint": {
    "version": "1.23.2"
  }
}
```

File: [./.yo-rc.json:16:5](./.yo-rc.json)

### FN017001 Run npm dedupe | Optional

If, after upgrading npm packages, when building the project you have errors similar to: "error TS2345: Argument of type 'SPHttpClientConfiguration' is not assignable to parameter of type 'SPHttpClientConfiguration'", try running 'npm dedupe' to cleanup npm packages.

Execute the following command:

```sh
npm dedupe
```

File: [./package.json](./package.json)

## Summary

### Execute script

```sh
npm i -SE @microsoft/sp-core-library@1.23.2 @microsoft/sp-lodash-subset@1.23.2 @microsoft/sp-office-ui-fabric-core@1.23.2 @microsoft/sp-webpart-base@1.23.2 @microsoft/sp-property-pane@1.23.2 @microsoft/sp-component-base@1.23.2 @microsoft/sp-adaptive-card-extension-base@1.23.2
npm i -DE @microsoft/sp-module-interfaces@1.23.2 @microsoft/eslint-plugin-spfx@1.23.2 @microsoft/eslint-config-spfx@1.23.2 @microsoft/spfx-web-build-rig@1.23.2 @microsoft/spfx-heft-plugins@1.23.2
npm dedupe
```

### Modify files

#### [./.yo-rc.json](./.yo-rc.json)

Update version in .yo-rc.json:

```json
{
  "@microsoft/generator-sharepoint": {
    "version": "1.23.2"
  }
}
```