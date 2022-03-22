import * as React from "react";
import { Provider, Flex, Text, Button, Header } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import { app } from "@microsoft/teams-js";
import { Login, MgtPersonCard, Person, PersonCardInteraction, ViewType } from "@microsoft/mgt-react";
import { Providers } from "@microsoft/mgt-react";
import { TeamsMsal2Provider } from "@microsoft/mgt-teams-msal2-provider";
import * as MicrosoftTeams from "@microsoft/teams-js";

/**
 * Implementation of the yo teams content page
 */
export const YoTeamsTab = () => {

    const [{ inTeams, theme, themeString, context }] = useTeams();
    const [entityId, setEntityId] = useState<string | undefined>();
    const [mgtTheme, setMgtTheme] = useState<string | undefined>('mgt-light');

    useEffect(() => {
        if (inTeams === true) {            
            TeamsMsal2Provider.microsoftTeamsLib = MicrosoftTeams;

            Providers.globalProvider = new TeamsMsal2Provider({
                clientId: process.env.CLIENT_ID!,
                scopes: ['User.Read'],
                authPopupUrl: '/auth.html'
            });
            app.notifySuccess();
        } else {
            setEntityId("Not in Microsoft Teams");
        }
    }, [inTeams]);

    useEffect(() => {
        if (context) {
            setEntityId(context.page.id);
        }
    }, [context]);
    
    useEffect(() => {
        if (themeString) {
            setMgtTheme(themeString === 'default' ? 'mgt-light' : 'mgt-dark');
        }
    }, [themeString]);

    /**
     * The render() method to create the UI of the tab
     */
    return (
        <Provider theme={theme}>
            <Flex fill={true} column styles={{
                padding: ".8rem 0 .8rem .5rem"
            }} className={mgtTheme}>
                <Flex.Item push>
                    <div>
                        <Login></Login>
                    </div>
                </Flex.Item>
                <Flex.Item>
                    <Header content="This is your MGT tab" />
                </Flex.Item>
                <Flex.Item>
                    <div>
                        Welcome <Person personQuery="me" view={ViewType.oneline} personCardInteraction={PersonCardInteraction.hover}></Person>
                    </div>
                </Flex.Item>
            </Flex>
        </Provider>
    );
};
