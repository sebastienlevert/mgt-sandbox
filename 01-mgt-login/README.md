# Login Component

## Summary

This is the first walk-through of all the components of the [Microsoft Graph Toolkit](https://aka.ms/mgt) highlighting the Login component.

## Prerequisites

* [Office 365 tenant](https://dev.office.com/sharepoint/docs/spfx/set-up-your-development-environment)
* [Live Server VS Code Extension](https://marketplace.visualstudio.com/items?itemName=ritwickdey.LiveServer)
* [Azure CLI](https://docs.microsoft.com/cli/azure/install-azure-cli) version 2.16.0 or higher (_Optional_)

## Version history

Version | Date              | Author                                                                                                                    | Comments
--------|-------------------|---------------------------------------------------------------------------------------------------------------------------|----------------
1.0     | December 14, 2021 | [Sébastien Levert](https://www.linkedin.com/in/sebastienlevert) ([@sebastienlevert](https://twitter.com/sebastienlevert)) | Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

* Clone this repository
* Create you Azure AD Application using the [manual approach](#create-a-new-azure-ad-application)

### Create a new Azure AD Application

1. Go to the [AAD Azure Portal](https://aad.portal.azure.com), then select **Azure Active Directory** from the left-hand side menu.
1. Select **App Registration** and click on **New Registration** button.
1. Fill the details to register an app:
   * Give a name to your application (_MGT Playground_)
   * Select **Accounts in any organizational directory (Any Azure AD directory - Multitenant)** as an access level
   * For the Redirect URI, select "Single-page application (SPA)" and use the following URL `http://localhost:5500`
   * Select **Register**

4. Copy the generated Application (client) ID. You will need this value later.

## Customize your `index.html` page

1. Open the `01-mgt-login` folder in Visual Studio Code
1. In the `index.html` file, replace the %YOUR_CLIENT_ID% placeholder with the generated Application (client) ID you copied earlier

## Run the sample

1. Start your web server
    * If using Live Server, right click in the index.html file and click on "Open with Live Server". For more details, see Live Server [documentation](https://github.com/ritwickdey/vscode-live-server#shortcuts-to-startstop-server).
1. Navigate to `http://localhost:5500`
1. Click on "Sign in"
1. Select your account and consent to the presented permission scopes

## Features

This sample illustrates the following concepts on top of MGT :

* Loading the Microsoft Graph Toolkit - [Docs](https://docs.microsoft.com/graph/toolkit/get-started/overview?tabs=html)
* Utilizing the Microsoft Graph Toolkit MSAL2 Provider - [Docs](https://docs.microsoft.com/graph/toolkit/providers/msal2)
* Utilizing the Microsoft Graph Toolkit login component - [Docs](https://docs.microsoft.com/graph/toolkit/components/login)