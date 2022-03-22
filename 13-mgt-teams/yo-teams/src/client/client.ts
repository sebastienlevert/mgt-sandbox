// Default entry point for client scripts
// Automatically generated
// Please avoid from modifying to much...
import * as ReactDOM from "react-dom";
import * as React from "react";
import { Providers } from "@microsoft/mgt-react";
import { TeamsMsal2Provider } from "@microsoft/mgt-teams-msal2-provider";
import * as MicrosoftTeams from "@microsoft/teams-js";
export const render = (type: any, element: HTMLElement) => {
    ReactDOM.render(React.createElement(type, {}), element);
};
// Automatically added for the yoTeamsTab tab
export * from "./yoTeamsTab/YoTeamsTab";

TeamsMsal2Provider.microsoftTeamsLib = MicrosoftTeams;

Providers.globalProvider = new TeamsMsal2Provider({
    clientId: process.env.CLIENT_ID!,
    scopes: ['User.Read'],
    authPopupUrl: '/auth.html'
});