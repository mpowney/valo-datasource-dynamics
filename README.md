# Valo Custom Data Source - Dynamics 365

## Summary

This extension to the Valo Universal Web Part demonstrates the ability to connect to Dynamics 365 data.

[picture of the solution in action, if possible]

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.11-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

* Target Microsoft 365 tenant must first have Valo Modern installed
* The Valo Modern app registration in the target environment must be configured to allow delegated API permissions to the Dynamics CRM (Common Data Service) API

Once this solution is added to the SharePoint app catalog, a new data source is added to the Valo Universal Web Part.

![Data Sources in the Valo Universal Web Part](./screenshot-uwp-data-source.png | width=100)

## Solution

Solution|Author(s)
--------|---------
valo-datasource-dynamics | Mark Powney, Valo [@mpowney](https://twitter.com/mpowney)

## Version history

Version|Date|Comments
-------|----|--------
0.0.1|September 7, 2020|Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

## Features

This solution allows page authors and site owners to connect the Valo Universal Web Part to Dynamics 365 Customer Experience data.

For this solution to work, the target environment must first be configured to securely enable access from the Valo modern site to Dynamics 365 APIs.  The Valo Tokens app, found in the App Registrations of Azure Active Directory, must have the **Dynamics CRM** delegate permission added, and admin consent must then be granted.

![App registrations in Azure Active Directory](./app-registration.png)

After the solution is deployed, a new data source is offered in the Universal Web Part properties.  Once selected, the data source accepts an API URL from the [Dynamics Customer Engagement Web API](https://docs.microsoft.com/en-us/dynamics365/customer-engagement/web-api/about?view=dynamics-ce-odata-9)

This extension provides the following capability

- Authentication to Dynamics 365 CRM Common Data Service via Azure Active Directory
- Provide Common Data Service as a source of data for Valo Universal Web Part templates

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
