# byod-library

## Summary

Short summary on functionality and used technologies.

[picture of the solution in action, if possible]

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.17.1-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

> Any special pre-requisites?

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| folder name | Author details (name, company, twitter alias with link) |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.1     | March 10, 2021   | Update comment  |
| 1.0     | January 29, 2021 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

> Include any additional steps as needed.

## Features

Description of the extension that expands upon high-level summary above.

This extension illustrates the following concepts:

- topic 1
- topic 2
- topic 3

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development



Codepen example
https://codepen.io/MayaMostafa/pen/OJqOzGm?editors=1100


## Packages Installations

- Fontawesome
npm i --save @fortawesome/fontawesome-svg-core 
npm i --save @fortawesome/free-solid-svg-icons
npm i --save @fortawesome/free-regular-svg-icons
npm i --save @fortawesome/free-brands-svg-icons
npm i --save @fortawesome/react-fontawesome@latest

- PnP sp & graph
npm install @pnp/sp @pnp/graph --save

- PnP controls
npm install @pnp/spfx-controls-react --save --save-exact

- PnP property controls
npm install @pnp/spfx-property-controls --save --save-exact


## Icons illustrations
https://iconscout.com/illustrations/catalogue

## Getting site & web id
https://www.sharepointdiary.com/2018/04/sharepoint-online-powershell-to-get-site-collection-web-id.html
Here is how to find the ID of a SharePoint Online site collection or subsite with REST endpoints:
1- To Get Site Collection ID, hit this URL in the browser: https://<tenant>.sharepoint.com/sites/<site-url>/_api/site/id
2- To get the subsite ID (or web ID) use: https://<tenant>.sharepoint.com/<site-url>/_api/web/id

## Getting list guid
https://sharepoint.sureshc.com/2017/10/how-to-get-sharepoint-list-library-guid-rest.html
https://pdsb1.sharepoint.com/sites/sLibrary/_api/web/lists/getByTitle('Professional')/Id

## Notes - Target Audience
https://stackoverflow.com/questions/66532774/sharepoint-online-audience-targeting-group-id
https://github.com/pnp/pnpjs/issues/2332
https://github.com/pnp/pnpcore/issues/399
https://sharepoint.stackexchange.com/questions/121452/access-target-audience-with-rest-api-the-field-or-property-audience-does-not

https://pdsb1.sharepoint.com/sites/my-site//_api/web/siteusers?$filter=Title%20eq%20%271106-SIEP-Team%20Members%27
https://pdsb1.sharepoint.com/sites/ModernDemos/_catalogs/users/detail.aspx?Paged=TRUE&p_ID=42&p_Title=NAVPREET%20KAUR&View=%7b2C99A7C6%2dC77B%2d4E83%2dA68F%2dAA2497ED1EDC%7d&SortField=ID&SortDir=Asc&PageFirstRow=31&InitialTabId=Ribbon%2ERead&VisibilityContext=WSSTabPersistence
https://pdsb1.sharepoint.com/sites/moderndemos//_api/web/siteusers?$filter=Email%20eq%20%27LearningTechnologySupportServices-DL@peelsb.com%27
