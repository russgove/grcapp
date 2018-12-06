# Upgrade project C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app to v1.7.0

Date: 2018-12-5

## Findings

Following is the list of steps required to upgrade your project to SharePoint Framework version 1.7.0. [Summary](#Summary) of the modifications is included at the end of the report.

### FN001001 @microsoft/sp-core-library | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-core-library

Execute the following command:

```sh
npm i @microsoft/sp-core-library@1.7.0 -SE
```

File: [./package.json](./package.json)

### FN001002 @microsoft/sp-lodash-subset | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-lodash-subset

Execute the following command:

```sh
npm i @microsoft/sp-lodash-subset@1.7.0 -SE
```

File: [./package.json](./package.json)

### FN001003 @microsoft/sp-office-ui-fabric-core | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-office-ui-fabric-core

Execute the following command:

```sh
npm i @microsoft/sp-office-ui-fabric-core@1.7.0 -SE
```

File: [./package.json](./package.json)

### FN001004 @microsoft/sp-webpart-base | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-webpart-base

Execute the following command:

```sh
npm i @microsoft/sp-webpart-base@1.7.0 -SE
```

File: [./package.json](./package.json)

### FN001005 @types/react | Required

Upgrade SharePoint Framework dependency package @types/react

Execute the following command:

```sh
npm i @types/react@16.4.2 -SE
```

File: [./package.json](./package.json)

### FN001006 @types/react-dom | Required

Upgrade SharePoint Framework dependency package @types/react-dom

Execute the following command:

```sh
npm i @types/react-dom@16.0.5 -SE
```

File: [./package.json](./package.json)

### FN001008 react | Required

Upgrade SharePoint Framework dependency package react

Execute the following command:

```sh
npm i react@16.3.2 -SE
```

File: [./package.json](./package.json)

### FN001009 react-dom | Required

Upgrade SharePoint Framework dependency package react-dom

Execute the following command:

```sh
npm i react-dom@16.3.2 -SE
```

File: [./package.json](./package.json)

### FN002001 @microsoft/sp-build-web | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-build-web

Execute the following command:

```sh
npm i @microsoft/sp-build-web@1.7.0 -DE
```

File: [./package.json](./package.json)

### FN002002 @microsoft/sp-module-interfaces | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-module-interfaces

Execute the following command:

```sh
npm i @microsoft/sp-module-interfaces@1.7.0 -DE
```

File: [./package.json](./package.json)

### FN002003 @microsoft/sp-webpart-workbench | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-webpart-workbench

Execute the following command:

```sh
npm i @microsoft/sp-webpart-workbench@1.7.0 -DE
```

File: [./package.json](./package.json)

### FN006003 package-solution.json isDomainIsolated | Required

Update package-solution.json isDomainIsolated

In file [./config/package-solution.json](./config/package-solution.json) update the code as follows:

```json
{
  "solution": {
    "isDomainIsolated": false
  }
}
```

File: [./config/package-solution.json](./config/package-solution.json)

### FN010001 .yo-rc.json version | Recommended

Update version in .yo-rc.json

In file [./.yo-rc.json](./.yo-rc.json) update the code as follows:

```json
{
  "@microsoft/generator-sharepoint": {
    "version": "1.7.0"
  }
}
```

File: [./.yo-rc.json](./.yo-rc.json)

### FN010007 .yo-rc.json isDomainIsolated | Recommended

Update isDomainIsolated in .yo-rc.json

In file [./.yo-rc.json](./.yo-rc.json) update the code as follows:

```json
{
  "@microsoft/generator-sharepoint": {
    "isDomainIsolated": false
  }
}
```

File: [./.yo-rc.json](./.yo-rc.json)

### FN018001 Web part Microsoft Teams tab resources folder | Optional

Create folder for Microsoft Teams tab resources

Execute the following command:

```sh
mkdir C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_businessRoleReview
```

File: [teams_businessRoleReview](teams_businessRoleReview)

### FN018001 Web part Microsoft Teams tab resources folder | Optional

Create folder for Microsoft Teams tab resources

Execute the following command:

```sh
mkdir C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_grcAzureManagement
```

File: [teams_grcAzureManagement](teams_grcAzureManagement)

### FN018001 Web part Microsoft Teams tab resources folder | Optional

Create folder for Microsoft Teams tab resources

Execute the following command:

```sh
mkdir C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_mitigatingControls
```

File: [teams_mitigatingControls](teams_mitigatingControls)

### FN018001 Web part Microsoft Teams tab resources folder | Optional

Create folder for Microsoft Teams tab resources

Execute the following command:

```sh
mkdir C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_roleToTCode
```

File: [teams_roleToTCode](teams_roleToTCode)

### FN018001 Web part Microsoft Teams tab resources folder | Optional

Create folder for Microsoft Teams tab resources

Execute the following command:

```sh


```

File: [teams_userAccess](teams_userAccess)

### FN018002 Web part Microsoft Teams tab manifest | Optional

Create Microsoft Teams tab manifest for the web part

Execute the following command:

```sh
cat > C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_businessRoleReview\manifest.json << EOF
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.2/MicrosoftTeams.schema.json",
  "manifestVersion": "1.2",
  "packageName": "businessRoleReview",
  "id": "cac695a6-969c-48da-968b-f76624f816a7",
  "version": "0.1",
  "developer": {
    "name": "SPFx + Teams Dev",
    "websiteUrl": "https://products.office.com/en-us/sharepoint/collaboration",
    "privacyUrl": "https://privacy.microsoft.com/en-us/privacystatement",
    "termsOfUseUrl": "https://www.microsoft.com/en-us/servicesagreement"
  },
  "name": {
    "short": "businessRoleReview"
  },
  "description": {
    "short": "businessRoleReview description",
    "full": "businessRoleReview description"
  },
  "icons": {
    "outline": "tab20x20.png",
    "color": "tab96x96.png"
  },
  "accentColor": "#004578",
  "configurableTabs": [
    {
      "configurationUrl": "https://{teamSiteDomain}{teamSitePath}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest={teamSitePath}/_layouts/15/teamshostedapp.aspx%3FopenPropertyPane=true%26teams%26componentId=cac695a6-969c-48da-968b-f76624f816a7",
      "canUpdateConfiguration": true,
      "scopes": [
        "team"
      ]
    }
  ],
  "validDomains": [
    "*.login.microsoftonline.com",
    "*.sharepoint.com",
    "*.sharepoint-df.com",
    "spoppe-a.akamaihd.net",
    "spoprod-a.akamaihd.net",
    "resourceseng.blob.core.windows.net",
    "msft.spoppe.com"
  ],
  "webApplicationInfo": {
    "resource": "https://{teamSiteDomain}",
    "id": "00000003-0000-0ff1-ce00-000000000000"
  }
}
EOF
```

File: [teams_businessRoleReview\manifest.json](teams_businessRoleReview\manifest.json)

### FN018002 Web part Microsoft Teams tab manifest | Optional

Create Microsoft Teams tab manifest for the web part

Execute the following command:

```sh
cat > C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_grcAzureManagement\manifest.json << EOF
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.2/MicrosoftTeams.schema.json",
  "manifestVersion": "1.2",
  "packageName": "GRCAzureManagement",
  "id": "e30095ac-6542-40dc-83c2-965680b847d3",
  "version": "0.1",
  "developer": {
    "name": "SPFx + Teams Dev",
    "websiteUrl": "https://products.office.com/en-us/sharepoint/collaboration",
    "privacyUrl": "https://privacy.microsoft.com/en-us/privacystatement",
    "termsOfUseUrl": "https://www.microsoft.com/en-us/servicesagreement"
  },
  "name": {
    "short": "GRCAzureManagement"
  },
  "description": {
    "short": "GRCAzureManagement description",
    "full": "GRCAzureManagement description"
  },
  "icons": {
    "outline": "tab20x20.png",
    "color": "tab96x96.png"
  },
  "accentColor": "#004578",
  "configurableTabs": [
    {
      "configurationUrl": "https://{teamSiteDomain}{teamSitePath}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest={teamSitePath}/_layouts/15/teamshostedapp.aspx%3FopenPropertyPane=true%26teams%26componentId=e30095ac-6542-40dc-83c2-965680b847d3",
      "canUpdateConfiguration": true,
      "scopes": [
        "team"
      ]
    }
  ],
  "validDomains": [
    "*.login.microsoftonline.com",
    "*.sharepoint.com",
    "*.sharepoint-df.com",
    "spoppe-a.akamaihd.net",
    "spoprod-a.akamaihd.net",
    "resourceseng.blob.core.windows.net",
    "msft.spoppe.com"
  ],
  "webApplicationInfo": {
    "resource": "https://{teamSiteDomain}",
    "id": "00000003-0000-0ff1-ce00-000000000000"
  }
}
EOF
```

File: [teams_grcAzureManagement\manifest.json](teams_grcAzureManagement\manifest.json)

### FN018002 Web part Microsoft Teams tab manifest | Optional

Create Microsoft Teams tab manifest for the web part

Execute the following command:

```sh
cat > C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_mitigatingControls\manifest.json << EOF
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.2/MicrosoftTeams.schema.json",
  "manifestVersion": "1.2",
  "packageName": "MitigatingControls",
  "id": "746860a8-18c2-4110-bd12-62ce3e2fc801",
  "version": "0.1",
  "developer": {
    "name": "SPFx + Teams Dev",
    "websiteUrl": "https://products.office.com/en-us/sharepoint/collaboration",
    "privacyUrl": "https://privacy.microsoft.com/en-us/privacystatement",
    "termsOfUseUrl": "https://www.microsoft.com/en-us/servicesagreement"
  },
  "name": {
    "short": "MitigatingControls"
  },
  "description": {
    "short": "mitigating controls user webpart",
    "full": "mitigating controls user webpart"
  },
  "icons": {
    "outline": "tab20x20.png",
    "color": "tab96x96.png"
  },
  "accentColor": "#004578",
  "configurableTabs": [
    {
      "configurationUrl": "https://{teamSiteDomain}{teamSitePath}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest={teamSitePath}/_layouts/15/teamshostedapp.aspx%3FopenPropertyPane=true%26teams%26componentId=746860a8-18c2-4110-bd12-62ce3e2fc801",
      "canUpdateConfiguration": true,
      "scopes": [
        "team"
      ]
    }
  ],
  "validDomains": [
    "*.login.microsoftonline.com",
    "*.sharepoint.com",
    "*.sharepoint-df.com",
    "spoppe-a.akamaihd.net",
    "spoprod-a.akamaihd.net",
    "resourceseng.blob.core.windows.net",
    "msft.spoppe.com"
  ],
  "webApplicationInfo": {
    "resource": "https://{teamSiteDomain}",
    "id": "00000003-0000-0ff1-ce00-000000000000"
  }
}
EOF
```

File: [teams_mitigatingControls\manifest.json](teams_mitigatingControls\manifest.json)

### FN018002 Web part Microsoft Teams tab manifest | Optional

Create Microsoft Teams tab manifest for the web part

Execute the following command:

```sh
cat > C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_roleToTCode\manifest.json << EOF
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.2/MicrosoftTeams.schema.json",
  "manifestVersion": "1.2",
  "packageName": "roleToTCode",
  "id": "03771b4d-4e98-44f7-98a1-7c41b066f52e",
  "version": "0.1",
  "developer": {
    "name": "SPFx + Teams Dev",
    "websiteUrl": "https://products.office.com/en-us/sharepoint/collaboration",
    "privacyUrl": "https://privacy.microsoft.com/en-us/privacystatement",
    "termsOfUseUrl": "https://www.microsoft.com/en-us/servicesagreement"
  },
  "name": {
    "short": "roleToTCode"
  },
  "description": {
    "short": "role to tcode",
    "full": "role to tcode"
  },
  "icons": {
    "outline": "tab20x20.png",
    "color": "tab96x96.png"
  },
  "accentColor": "#004578",
  "configurableTabs": [
    {
      "configurationUrl": "https://{teamSiteDomain}{teamSitePath}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest={teamSitePath}/_layouts/15/teamshostedapp.aspx%3FopenPropertyPane=true%26teams%26componentId=03771b4d-4e98-44f7-98a1-7c41b066f52e",
      "canUpdateConfiguration": true,
      "scopes": [
        "team"
      ]
    }
  ],
  "validDomains": [
    "*.login.microsoftonline.com",
    "*.sharepoint.com",
    "*.sharepoint-df.com",
    "spoppe-a.akamaihd.net",
    "spoprod-a.akamaihd.net",
    "resourceseng.blob.core.windows.net",
    "msft.spoppe.com"
  ],
  "webApplicationInfo": {
    "resource": "https://{teamSiteDomain}",
    "id": "00000003-0000-0ff1-ce00-000000000000"
  }
}
EOF
```

File: [teams_roleToTCode\manifest.json](teams_roleToTCode\manifest.json)

### FN018002 Web part Microsoft Teams tab manifest | Optional

Create Microsoft Teams tab manifest for the web part

Execute the following command:

```sh
cat > C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_userAccess\manifest.json << EOF
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.2/MicrosoftTeams.schema.json",
  "manifestVersion": "1.2",
  "packageName": "User Access Review",
  "id": "284bb3d2-47d5-44c6-8ca6-457d359f939a",
  "version": "0.1",
  "developer": {
    "name": "SPFx + Teams Dev",
    "websiteUrl": "https://products.office.com/en-us/sharepoint/collaboration",
    "privacyUrl": "https://privacy.microsoft.com/en-us/privacystatement",
    "termsOfUseUrl": "https://www.microsoft.com/en-us/servicesagreement"
  },
  "name": {
    "short": "User Access Review"
  },
  "description": {
    "short": "This webpart is placed on the homepage of a User Access review site so that users can conduct the review",
    "full": "This webpart is placed on the homepage of a User Access review site so that users can conduct the review"
  },
  "icons": {
    "outline": "tab20x20.png",
    "color": "tab96x96.png"
  },
  "accentColor": "#004578",
  "configurableTabs": [
    {
      "configurationUrl": "https://{teamSiteDomain}{teamSitePath}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest={teamSitePath}/_layouts/15/teamshostedapp.aspx%3FopenPropertyPane=true%26teams%26componentId=284bb3d2-47d5-44c6-8ca6-457d359f939a",
      "canUpdateConfiguration": true,
      "scopes": [
        "team"
      ]
    }
  ],
  "validDomains": [
    "*.login.microsoftonline.com",
    "*.sharepoint.com",
    "*.sharepoint-df.com",
    "spoppe-a.akamaihd.net",
    "spoprod-a.akamaihd.net",
    "resourceseng.blob.core.windows.net",
    "msft.spoppe.com"
  ],
  "webApplicationInfo": {
    "resource": "https://{teamSiteDomain}",
    "id": "00000003-0000-0ff1-ce00-000000000000"
  }
}
EOF
```

File: [teams_userAccess\manifest.json](teams_userAccess\manifest.json)

### FN018003 Web part Microsoft Teams tab small icon | Optional

Create Microsoft Teams tab small icon for the web part

Execute the following command:

```sh
cp C:\Users\trwg1\AppData\Roaming\npm\node_modules\@pnp\office365-cli\dist\o365\spfx\commands\project\project-upgrade\assets\tab20x20.png C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_businessRoleReview\tab20x20.png
```

File: [teams_businessRoleReview\tab20x20.png](teams_businessRoleReview\tab20x20.png)

### FN018003 Web part Microsoft Teams tab small icon | Optional

Create Microsoft Teams tab small icon for the web part

Execute the following command:

```sh
cp C:\Users\trwg1\AppData\Roaming\npm\node_modules\@pnp\office365-cli\dist\o365\spfx\commands\project\project-upgrade\assets\tab20x20.png C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_grcAzureManagement\tab20x20.png
```

File: [teams_grcAzureManagement\tab20x20.png](teams_grcAzureManagement\tab20x20.png)

### FN018003 Web part Microsoft Teams tab small icon | Optional

Create Microsoft Teams tab small icon for the web part

Execute the following command:

```sh
cp C:\Users\trwg1\AppData\Roaming\npm\node_modules\@pnp\office365-cli\dist\o365\spfx\commands\project\project-upgrade\assets\tab20x20.png C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_mitigatingControls\tab20x20.png
```

File: [teams_mitigatingControls\tab20x20.png](teams_mitigatingControls\tab20x20.png)

### FN018003 Web part Microsoft Teams tab small icon | Optional

Create Microsoft Teams tab small icon for the web part

Execute the following command:

```sh
cp C:\Users\trwg1\AppData\Roaming\npm\node_modules\@pnp\office365-cli\dist\o365\spfx\commands\project\project-upgrade\assets\tab20x20.png C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_roleToTCode\tab20x20.png
```

File: [teams_roleToTCode\tab20x20.png](teams_roleToTCode\tab20x20.png)

### FN018003 Web part Microsoft Teams tab small icon | Optional

Create Microsoft Teams tab small icon for the web part

Execute the following command:

```sh
cp C:\Users\trwg1\AppData\Roaming\npm\node_modules\@pnp\office365-cli\dist\o365\spfx\commands\project\project-upgrade\assets\tab20x20.png C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_userAccess\tab20x20.png
```

File: [teams_userAccess\tab20x20.png](teams_userAccess\tab20x20.png)

### FN018004 Web part Microsoft Teams tab large icon | Optional

Create Microsoft Teams tab large icon for the web part

Execute the following command:

```sh
cp C:\Users\trwg1\AppData\Roaming\npm\node_modules\@pnp\office365-cli\dist\o365\spfx\commands\project\project-upgrade\assets\tab96x96.png C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_businessRoleReview\tab96x96.png
```

File: [teams_businessRoleReview\tab96x96.png](teams_businessRoleReview\tab96x96.png)

### FN018004 Web part Microsoft Teams tab large icon | Optional

Create Microsoft Teams tab large icon for the web part

Execute the following command:

```sh
cp C:\Users\trwg1\AppData\Roaming\npm\node_modules\@pnp\office365-cli\dist\o365\spfx\commands\project\project-upgrade\assets\tab96x96.png C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_grcAzureManagement\tab96x96.png
```

File: [teams_grcAzureManagement\tab96x96.png](teams_grcAzureManagement\tab96x96.png)

### FN018004 Web part Microsoft Teams tab large icon | Optional

Create Microsoft Teams tab large icon for the web part

Execute the following command:

```sh
cp C:\Users\trwg1\AppData\Roaming\npm\node_modules\@pnp\office365-cli\dist\o365\spfx\commands\project\project-upgrade\assets\tab96x96.png C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_mitigatingControls\tab96x96.png
```

File: [teams_mitigatingControls\tab96x96.png](teams_mitigatingControls\tab96x96.png)

### FN018004 Web part Microsoft Teams tab large icon | Optional

Create Microsoft Teams tab large icon for the web part

Execute the following command:

```sh
cp C:\Users\trwg1\AppData\Roaming\npm\node_modules\@pnp\office365-cli\dist\o365\spfx\commands\project\project-upgrade\assets\tab96x96.png C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_roleToTCode\tab96x96.png
```

File: [teams_roleToTCode\tab96x96.png](teams_roleToTCode\tab96x96.png)

### FN018004 Web part Microsoft Teams tab large icon | Optional

Create Microsoft Teams tab large icon for the web part

Execute the following command:

```sh
cp C:\Users\trwg1\AppData\Roaming\npm\node_modules\@pnp\office365-cli\dist\o365\spfx\commands\project\project-upgrade\assets\tab96x96.png C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_userAccess\tab96x96.png
```

File: [teams_userAccess\tab96x96.png](teams_userAccess\tab96x96.png)

### FN002008 tslint-microsoft-contrib | Required

Install SharePoint Framework dev dependency package tslint-microsoft-contrib

Execute the following command:

```sh
npm i tslint-microsoft-contrib@5.0.0 -DE
```

File: [./package.json](./package.json)

### FN012012 tsconfig.json include property | Required

Update tsconfig.json include property

In file [./tsconfig.json](./tsconfig.json) update the code as follows:

```json
{
  "include": [
    "src/**/*.ts"
  ]
}
```

File: [./tsconfig.json](./tsconfig.json)

### FN012013 tsconfig.json exclude property | Required

Update tsconfig.json exclude property

In file [./tsconfig.json](./tsconfig.json) update the code as follows:

```json
{
  "exclude": [
    "node_modules",
    "lib"
  ]
}
```

File: [./tsconfig.json](./tsconfig.json)

### FN015003 ./tslint.json | Required

Add file ./tslint.json

Execute the following command:

```sh
cat > ./tslint.json << EOF
{
  "rulesDirectory": [
    "tslint-microsoft-contrib"
  ],
  "rules": {
    "class-name": false,
    "export-name": false,
    "forin": false,
    "label-position": false,
    "member-access": true,
    "no-arg": false,
    "no-console": false,
    "no-construct": false,
    "no-duplicate-variable": true,
    "no-eval": false,
    "no-function-expression": true,
    "no-internal-module": true,
    "no-shadowed-variable": true,
    "no-switch-case-fall-through": true,
    "no-unnecessary-semicolons": true,
    "no-unused-expression": true,
    "no-use-before-declare": true,
    "no-with-statement": true,
    "semicolon": true,
    "trailing-comma": false,
    "typedef": false,
    "typedef-whitespace": false,
    "use-named-parameter": true,
    "variable-name": false,
    "whitespace": false
  }
}
EOF
```

File: [./tslint.json](./tslint.json)

### FN015005 ./src/index.ts | Required

Add file ./src/index.ts

Execute the following command:

```sh
cat > ./src/index.ts << EOF
// A file is required to be in the root of the /src directory by the TypeScript compiler

EOF
```

File: [./src/index.ts](./src/index.ts)

### FN001007 @types/webpack-env | Required

Upgrade SharePoint Framework dependency package @types/webpack-env

Execute the following command:

```sh
npm i @types/webpack-env@1.13.1 -SE
```

File: [./package.json](./package.json)

### FN001010 @types/es6-promise | Required

Install SharePoint Framework dependency package @types/es6-promise

Execute the following command:

```sh
npm i @types/es6-promise@0.0.33 -SE
```

File: [./package.json](./package.json)

### FN002005 @types/chai | Required

Upgrade SharePoint Framework dev dependency package @types/chai

Execute the following command:

```sh
npm i @types/chai@3.4.34 -DE
```

File: [./package.json](./package.json)

### FN002006 @types/mocha | Required

Upgrade SharePoint Framework dev dependency package @types/mocha

Execute the following command:

```sh
npm i @types/mocha@2.2.38 -DE
```

File: [./package.json](./package.json)

### FN003001 config.json schema | Required

Update config.json schema URL

In file [./config/config.json](./config/config.json) update the code as follows:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/config.2.0.schema.json"
}
```

File: [./config/config.json](./config/config.json)

### FN004001 copy-assets.json schema | Required

Update copy-assets.json schema URL

In file [./config/copy-assets.json](./config/copy-assets.json) update the code as follows:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/copy-assets.schema.json"
}
```

File: [./config/copy-assets.json](./config/copy-assets.json)

### FN005001 deploy-azure-storage.json schema | Required

Update deploy-azure-storage.json schema URL

In file [./config/deploy-azure-storage.json](./config/deploy-azure-storage.json) update the code as follows:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/deploy-azure-storage.schema.json"
}
```

File: [./config/deploy-azure-storage.json](./config/deploy-azure-storage.json)

### FN006001 package-solution.json schema | Required

Update package-solution.json schema URL

In file [./config/package-solution.json](./config/package-solution.json) update the code as follows:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/package-solution.schema.json"
}
```

File: [./config/package-solution.json](./config/package-solution.json)

### FN007001 serve.json schema | Required

Update serve.json schema URL

In file [./config/serve.json](./config/serve.json) update the code as follows:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/core-build/serve.schema.json"
}
```

File: [./config/serve.json](./config/serve.json)

### FN009001 write-manifests.json schema | Required

Update write-manifests.json schema URL

In file [./config/write-manifests.json](./config/write-manifests.json) update the code as follows:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/write-manifests.schema.json"
}
```

File: [./config/write-manifests.json](./config/write-manifests.json)

### FN010002 .yo-rc.json isCreatingSolution | Recommended

Update isCreatingSolution in .yo-rc.json

In file [./.yo-rc.json](./.yo-rc.json) update the code as follows:

```json
{
  "@microsoft/generator-sharepoint": {
    "isCreatingSolution": true
  }
}
```

File: [./.yo-rc.json](./.yo-rc.json)

### FN010003 .yo-rc.json packageManager | Recommended

Update packageManager in .yo-rc.json

In file [./.yo-rc.json](./.yo-rc.json) update the code as follows:

```json
{
  "@microsoft/generator-sharepoint": {
    "packageManager": "npm"
  }
}
```

File: [./.yo-rc.json](./.yo-rc.json)

### FN010004 .yo-rc.json componentType | Recommended

Update componentType in .yo-rc.json

In file [./.yo-rc.json](./.yo-rc.json) update the code as follows:

```json
{
  "@microsoft/generator-sharepoint": {
    "componentType": "webpart"
  }
}
```

File: [./.yo-rc.json](./.yo-rc.json)

### FN011001 Web part manifest schema | Required

Update schema in manifest

In file [src\webparts\businessRoleReview\BusinessRoleReviewWebPart.manifest.json](src\webparts\businessRoleReview\BusinessRoleReviewWebPart.manifest.json) update the code as follows:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx/client-side-web-part-manifest.schema.json"
}
```

File: [src\webparts\businessRoleReview\BusinessRoleReviewWebPart.manifest.json](src\webparts\businessRoleReview\BusinessRoleReviewWebPart.manifest.json)

### FN011001 Web part manifest schema | Required

Update schema in manifest

In file [src\webparts\grcAzureManagement\GrcAzureManagementWebPart.manifest.json](src\webparts\grcAzureManagement\GrcAzureManagementWebPart.manifest.json) update the code as follows:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx/client-side-web-part-manifest.schema.json"
}
```

File: [src\webparts\grcAzureManagement\GrcAzureManagementWebPart.manifest.json](src\webparts\grcAzureManagement\GrcAzureManagementWebPart.manifest.json)

### FN011001 Web part manifest schema | Required

Update schema in manifest

In file [src\webparts\mitigatingControls\MitigatingControlsWebPart.manifest.json](src\webparts\mitigatingControls\MitigatingControlsWebPart.manifest.json) update the code as follows:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx/client-side-web-part-manifest.schema.json"
}
```

File: [src\webparts\mitigatingControls\MitigatingControlsWebPart.manifest.json](src\webparts\mitigatingControls\MitigatingControlsWebPart.manifest.json)

### FN011001 Web part manifest schema | Required

Update schema in manifest

In file [src\webparts\roleToTCode\RoleToTCodeWebPart.manifest.json](src\webparts\roleToTCode\RoleToTCodeWebPart.manifest.json) update the code as follows:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx/client-side-web-part-manifest.schema.json"
}
```

File: [src\webparts\roleToTCode\RoleToTCodeWebPart.manifest.json](src\webparts\roleToTCode\RoleToTCodeWebPart.manifest.json)

### FN011001 Web part manifest schema | Required

Update schema in manifest

In file [src\webparts\userAccess\UserAccessWebPart.manifest.json](src\webparts\userAccess\UserAccessWebPart.manifest.json) update the code as follows:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx/client-side-web-part-manifest.schema.json"
}
```

File: [src\webparts\userAccess\UserAccessWebPart.manifest.json](src\webparts\userAccess\UserAccessWebPart.manifest.json)

### FN012001 tsconfig.json module | Required

Update module type in tsconfig.json

In file [./tsconfig.json](./tsconfig.json) update the code as follows:

```json
{
  "compilerOptions": {
    "module": "esnext"
  }
}
```

File: [./tsconfig.json](./tsconfig.json)

### FN012002 tsconfig.json moduleResolution | Required

Update moduleResolution in tsconfig.json

In file [./tsconfig.json](./tsconfig.json) update the code as follows:

```json
{
  "compilerOptions": {
    "moduleResolution": "node"
  }
}
```

File: [./tsconfig.json](./tsconfig.json)

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
npm i @microsoft/sp-core-library@1.7.0 @microsoft/sp-lodash-subset@1.7.0 @microsoft/sp-office-ui-fabric-core@1.7.0 @microsoft/sp-webpart-base@1.7.0 @types/react@16.4.2 @types/react-dom@16.0.5 react@16.3.2 react-dom@16.3.2 @types/webpack-env@1.13.1 @types/es6-promise@0.0.33 -SE
npm i @microsoft/sp-build-web@1.7.0 @microsoft/sp-module-interfaces@1.7.0 @microsoft/sp-webpart-workbench@1.7.0 tslint-microsoft-contrib@5.0.0 @types/chai@3.4.34 @types/mocha@2.2.38 -DE
mkdir C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_businessRoleReview
mkdir C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_grcAzureManagement
mkdir C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_mitigatingControls
mkdir C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_roleToTCode
mkdir C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_userAccess
cat > C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_businessRoleReview\manifest.json << EOF
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.2/MicrosoftTeams.schema.json",
  "manifestVersion": "1.2",
  "packageName": "businessRoleReview",
  "id": "cac695a6-969c-48da-968b-f76624f816a7",
  "version": "0.1",
  "developer": {
    "name": "SPFx + Teams Dev",
    "websiteUrl": "https://products.office.com/en-us/sharepoint/collaboration",
    "privacyUrl": "https://privacy.microsoft.com/en-us/privacystatement",
    "termsOfUseUrl": "https://www.microsoft.com/en-us/servicesagreement"
  },
  "name": {
    "short": "businessRoleReview"
  },
  "description": {
    "short": "businessRoleReview description",
    "full": "businessRoleReview description"
  },
  "icons": {
    "outline": "tab20x20.png",
    "color": "tab96x96.png"
  },
  "accentColor": "#004578",
  "configurableTabs": [
    {
      "configurationUrl": "https://{teamSiteDomain}{teamSitePath}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest={teamSitePath}/_layouts/15/teamshostedapp.aspx%3FopenPropertyPane=true%26teams%26componentId=cac695a6-969c-48da-968b-f76624f816a7",
      "canUpdateConfiguration": true,
      "scopes": [
        "team"
      ]
    }
  ],
  "validDomains": [
    "*.login.microsoftonline.com",
    "*.sharepoint.com",
    "*.sharepoint-df.com",
    "spoppe-a.akamaihd.net",
    "spoprod-a.akamaihd.net",
    "resourceseng.blob.core.windows.net",
    "msft.spoppe.com"
  ],
  "webApplicationInfo": {
    "resource": "https://{teamSiteDomain}",
    "id": "00000003-0000-0ff1-ce00-000000000000"
  }
}
EOF
cat > C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_grcAzureManagement\manifest.json << EOF
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.2/MicrosoftTeams.schema.json",
  "manifestVersion": "1.2",
  "packageName": "GRCAzureManagement",
  "id": "e30095ac-6542-40dc-83c2-965680b847d3",
  "version": "0.1",
  "developer": {
    "name": "SPFx + Teams Dev",
    "websiteUrl": "https://products.office.com/en-us/sharepoint/collaboration",
    "privacyUrl": "https://privacy.microsoft.com/en-us/privacystatement",
    "termsOfUseUrl": "https://www.microsoft.com/en-us/servicesagreement"
  },
  "name": {
    "short": "GRCAzureManagement"
  },
  "description": {
    "short": "GRCAzureManagement description",
    "full": "GRCAzureManagement description"
  },
  "icons": {
    "outline": "tab20x20.png",
    "color": "tab96x96.png"
  },
  "accentColor": "#004578",
  "configurableTabs": [
    {
      "configurationUrl": "https://{teamSiteDomain}{teamSitePath}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest={teamSitePath}/_layouts/15/teamshostedapp.aspx%3FopenPropertyPane=true%26teams%26componentId=e30095ac-6542-40dc-83c2-965680b847d3",
      "canUpdateConfiguration": true,
      "scopes": [
        "team"
      ]
    }
  ],
  "validDomains": [
    "*.login.microsoftonline.com",
    "*.sharepoint.com",
    "*.sharepoint-df.com",
    "spoppe-a.akamaihd.net",
    "spoprod-a.akamaihd.net",
    "resourceseng.blob.core.windows.net",
    "msft.spoppe.com"
  ],
  "webApplicationInfo": {
    "resource": "https://{teamSiteDomain}",
    "id": "00000003-0000-0ff1-ce00-000000000000"
  }
}
EOF
cat > C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_mitigatingControls\manifest.json << EOF
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.2/MicrosoftTeams.schema.json",
  "manifestVersion": "1.2",
  "packageName": "MitigatingControls",
  "id": "746860a8-18c2-4110-bd12-62ce3e2fc801",
  "version": "0.1",
  "developer": {
    "name": "SPFx + Teams Dev",
    "websiteUrl": "https://products.office.com/en-us/sharepoint/collaboration",
    "privacyUrl": "https://privacy.microsoft.com/en-us/privacystatement",
    "termsOfUseUrl": "https://www.microsoft.com/en-us/servicesagreement"
  },
  "name": {
    "short": "MitigatingControls"
  },
  "description": {
    "short": "mitigating controls user webpart",
    "full": "mitigating controls user webpart"
  },
  "icons": {
    "outline": "tab20x20.png",
    "color": "tab96x96.png"
  },
  "accentColor": "#004578",
  "configurableTabs": [
    {
      "configurationUrl": "https://{teamSiteDomain}{teamSitePath}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest={teamSitePath}/_layouts/15/teamshostedapp.aspx%3FopenPropertyPane=true%26teams%26componentId=746860a8-18c2-4110-bd12-62ce3e2fc801",
      "canUpdateConfiguration": true,
      "scopes": [
        "team"
      ]
    }
  ],
  "validDomains": [
    "*.login.microsoftonline.com",
    "*.sharepoint.com",
    "*.sharepoint-df.com",
    "spoppe-a.akamaihd.net",
    "spoprod-a.akamaihd.net",
    "resourceseng.blob.core.windows.net",
    "msft.spoppe.com"
  ],
  "webApplicationInfo": {
    "resource": "https://{teamSiteDomain}",
    "id": "00000003-0000-0ff1-ce00-000000000000"
  }
}
EOF
cat > C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_roleToTCode\manifest.json << EOF
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.2/MicrosoftTeams.schema.json",
  "manifestVersion": "1.2",
  "packageName": "roleToTCode",
  "id": "03771b4d-4e98-44f7-98a1-7c41b066f52e",
  "version": "0.1",
  "developer": {
    "name": "SPFx + Teams Dev",
    "websiteUrl": "https://products.office.com/en-us/sharepoint/collaboration",
    "privacyUrl": "https://privacy.microsoft.com/en-us/privacystatement",
    "termsOfUseUrl": "https://www.microsoft.com/en-us/servicesagreement"
  },
  "name": {
    "short": "roleToTCode"
  },
  "description": {
    "short": "role to tcode",
    "full": "role to tcode"
  },
  "icons": {
    "outline": "tab20x20.png",
    "color": "tab96x96.png"
  },
  "accentColor": "#004578",
  "configurableTabs": [
    {
      "configurationUrl": "https://{teamSiteDomain}{teamSitePath}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest={teamSitePath}/_layouts/15/teamshostedapp.aspx%3FopenPropertyPane=true%26teams%26componentId=03771b4d-4e98-44f7-98a1-7c41b066f52e",
      "canUpdateConfiguration": true,
      "scopes": [
        "team"
      ]
    }
  ],
  "validDomains": [
    "*.login.microsoftonline.com",
    "*.sharepoint.com",
    "*.sharepoint-df.com",
    "spoppe-a.akamaihd.net",
    "spoprod-a.akamaihd.net",
    "resourceseng.blob.core.windows.net",
    "msft.spoppe.com"
  ],
  "webApplicationInfo": {
    "resource": "https://{teamSiteDomain}",
    "id": "00000003-0000-0ff1-ce00-000000000000"
  }
}
EOF
cat > C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_userAccess\manifest.json << EOF
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.2/MicrosoftTeams.schema.json",
  "manifestVersion": "1.2",
  "packageName": "User Access Review",
  "id": "284bb3d2-47d5-44c6-8ca6-457d359f939a",
  "version": "0.1",
  "developer": {
    "name": "SPFx + Teams Dev",
    "websiteUrl": "https://products.office.com/en-us/sharepoint/collaboration",
    "privacyUrl": "https://privacy.microsoft.com/en-us/privacystatement",
    "termsOfUseUrl": "https://www.microsoft.com/en-us/servicesagreement"
  },
  "name": {
    "short": "User Access Review"
  },
  "description": {
    "short": "This webpart is placed on the homepage of a User Access review site so that users can conduct the review",
    "full": "This webpart is placed on the homepage of a User Access review site so that users can conduct the review"
  },
  "icons": {
    "outline": "tab20x20.png",
    "color": "tab96x96.png"
  },
  "accentColor": "#004578",
  "configurableTabs": [
    {
      "configurationUrl": "https://{teamSiteDomain}{teamSitePath}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest={teamSitePath}/_layouts/15/teamshostedapp.aspx%3FopenPropertyPane=true%26teams%26componentId=284bb3d2-47d5-44c6-8ca6-457d359f939a",
      "canUpdateConfiguration": true,
      "scopes": [
        "team"
      ]
    }
  ],
  "validDomains": [
    "*.login.microsoftonline.com",
    "*.sharepoint.com",
    "*.sharepoint-df.com",
    "spoppe-a.akamaihd.net",
    "spoprod-a.akamaihd.net",
    "resourceseng.blob.core.windows.net",
    "msft.spoppe.com"
  ],
  "webApplicationInfo": {
    "resource": "https://{teamSiteDomain}",
    "id": "00000003-0000-0ff1-ce00-000000000000"
  }
}
EOF
cp C:\Users\trwg1\AppData\Roaming\npm\node_modules\@pnp\office365-cli\dist\o365\spfx\commands\project\project-upgrade\assets\tab20x20.png C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_businessRoleReview\tab20x20.png
cp C:\Users\trwg1\AppData\Roaming\npm\node_modules\@pnp\office365-cli\dist\o365\spfx\commands\project\project-upgrade\assets\tab20x20.png C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_grcAzureManagement\tab20x20.png
cp C:\Users\trwg1\AppData\Roaming\npm\node_modules\@pnp\office365-cli\dist\o365\spfx\commands\project\project-upgrade\assets\tab20x20.png C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_mitigatingControls\tab20x20.png
cp C:\Users\trwg1\AppData\Roaming\npm\node_modules\@pnp\office365-cli\dist\o365\spfx\commands\project\project-upgrade\assets\tab20x20.png C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_roleToTCode\tab20x20.png
cp C:\Users\trwg1\AppData\Roaming\npm\node_modules\@pnp\office365-cli\dist\o365\spfx\commands\project\project-upgrade\assets\tab20x20.png C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_userAccess\tab20x20.png
cp C:\Users\trwg1\AppData\Roaming\npm\node_modules\@pnp\office365-cli\dist\o365\spfx\commands\project\project-upgrade\assets\tab96x96.png C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_businessRoleReview\tab96x96.png
cp C:\Users\trwg1\AppData\Roaming\npm\node_modules\@pnp\office365-cli\dist\o365\spfx\commands\project\project-upgrade\assets\tab96x96.png C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_grcAzureManagement\tab96x96.png
cp C:\Users\trwg1\AppData\Roaming\npm\node_modules\@pnp\office365-cli\dist\o365\spfx\commands\project\project-upgrade\assets\tab96x96.png C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_mitigatingControls\tab96x96.png
cp C:\Users\trwg1\AppData\Roaming\npm\node_modules\@pnp\office365-cli\dist\o365\spfx\commands\project\project-upgrade\assets\tab96x96.png C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_roleToTCode\tab96x96.png
cp C:\Users\trwg1\AppData\Roaming\npm\node_modules\@pnp\office365-cli\dist\o365\spfx\commands\project\project-upgrade\assets\tab96x96.png C:\Users\trwg1\OneDrive - Tronox\My Documents\GRC\app\teams_userAccess\tab96x96.png
cat > ./tslint.json << EOF
{
  "rulesDirectory": [
    "tslint-microsoft-contrib"
  ],
  "rules": {
    "class-name": false,
    "export-name": false,
    "forin": false,
    "label-position": false,
    "member-access": true,
    "no-arg": false,
    "no-console": false,
    "no-construct": false,
    "no-duplicate-variable": true,
    "no-eval": false,
    "no-function-expression": true,
    "no-internal-module": true,
    "no-shadowed-variable": true,
    "no-switch-case-fall-through": true,
    "no-unnecessary-semicolons": true,
    "no-unused-expression": true,
    "no-use-before-declare": true,
    "no-with-statement": true,
    "semicolon": true,
    "trailing-comma": false,
    "typedef": false,
    "typedef-whitespace": false,
    "use-named-parameter": true,
    "variable-name": false,
    "whitespace": false
  }
}
EOF
cat > ./src/index.ts << EOF
// A file is required to be in the root of the /src directory by the TypeScript compiler

EOF
npm dedupe
```

### Modify files

#### [./config/package-solution.json](./config/package-solution.json)

Update package-solution.json isDomainIsolated:

```json
{
  "solution": {
    "isDomainIsolated": false
  }
}
```

Update package-solution.json schema URL:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/package-solution.schema.json"
}
```

#### [./.yo-rc.json](./.yo-rc.json)

Update version in .yo-rc.json:

```json
{
  "@microsoft/generator-sharepoint": {
    "version": "1.7.0"
  }
}
```

Update isDomainIsolated in .yo-rc.json:

```json
{
  "@microsoft/generator-sharepoint": {
    "isDomainIsolated": false
  }
}
```

Update isCreatingSolution in .yo-rc.json:

```json
{
  "@microsoft/generator-sharepoint": {
    "isCreatingSolution": true
  }
}
```

Update packageManager in .yo-rc.json:

```json
{
  "@microsoft/generator-sharepoint": {
    "packageManager": "npm"
  }
}
```

Update componentType in .yo-rc.json:

```json
{
  "@microsoft/generator-sharepoint": {
    "componentType": "webpart"
  }
}
```

#### [./tsconfig.json](./tsconfig.json)

Update tsconfig.json include property:

```json
{
  "include": [
    "src/**/*.ts"
  ]
}
```

Update tsconfig.json exclude property:

```json
{
  "exclude": [
    "node_modules",
    "lib"
  ]
}
```

Update module type in tsconfig.json:

```json
{
  "compilerOptions": {
    "module": "esnext"
  }
}
```

Update moduleResolution in tsconfig.json:

```json
{
  "compilerOptions": {
    "moduleResolution": "node"
  }
}
```

#### [./config/config.json](./config/config.json)

Update config.json schema URL:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/config.2.0.schema.json"
}
```

#### [./config/copy-assets.json](./config/copy-assets.json)

Update copy-assets.json schema URL:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/copy-assets.schema.json"
}
```

#### [./config/deploy-azure-storage.json](./config/deploy-azure-storage.json)

Update deploy-azure-storage.json schema URL:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/deploy-azure-storage.schema.json"
}
```

#### [./config/serve.json](./config/serve.json)

Update serve.json schema URL:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/core-build/serve.schema.json"
}
```

#### [./config/write-manifests.json](./config/write-manifests.json)

Update write-manifests.json schema URL:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/write-manifests.schema.json"
}
```

#### [src\webparts\businessRoleReview\BusinessRoleReviewWebPart.manifest.json](src\webparts\businessRoleReview\BusinessRoleReviewWebPart.manifest.json)

Update schema in manifest:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx/client-side-web-part-manifest.schema.json"
}
```

#### [src\webparts\grcAzureManagement\GrcAzureManagementWebPart.manifest.json](src\webparts\grcAzureManagement\GrcAzureManagementWebPart.manifest.json)

Update schema in manifest:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx/client-side-web-part-manifest.schema.json"
}
```

#### [src\webparts\mitigatingControls\MitigatingControlsWebPart.manifest.json](src\webparts\mitigatingControls\MitigatingControlsWebPart.manifest.json)

Update schema in manifest:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx/client-side-web-part-manifest.schema.json"
}
```

#### [src\webparts\roleToTCode\RoleToTCodeWebPart.manifest.json](src\webparts\roleToTCode\RoleToTCodeWebPart.manifest.json)

Update schema in manifest:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx/client-side-web-part-manifest.schema.json"
}
```

#### [src\webparts\userAccess\UserAccessWebPart.manifest.json](src\webparts\userAccess\UserAccessWebPart.manifest.json)

Update schema in manifest:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx/client-side-web-part-manifest.schema.json"
}
```
