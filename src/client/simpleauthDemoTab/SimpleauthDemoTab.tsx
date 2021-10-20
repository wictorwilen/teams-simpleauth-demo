import * as React from "react";
import { Provider, Flex, Text, Button, Header } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import jwtDecode from "jwt-decode";
import * as teamsFx from "@microsoft/teamsfx";

/**
 * Implementation of the Simpleauth Demo content page
 */
export const SimpleauthDemoTab = () => {

    const [{ inTeams, theme, context }] = useTeams();
    const [entityId, setEntityId] = useState<string | undefined>();
    const [name, setName] = useState<string>();
    const [error, setError] = useState<string>();
    const [presence, setPresence] = useState<string>("Unknown");

    useEffect(() => {
        if (inTeams === true) {
            microsoftTeams.authentication.getAuthToken({
                successCallback: async (token: string) => {
                    const decoded: { [key: string]: any; } = jwtDecode(token) as { [key: string]: any; };
                    setName(decoded!.name);
                    teamsFx.loadConfiguration({
                        authentication: {
                            initiateLoginEndpoint: `https://${process.env.PUBLIC_HOSTNAME}/ile`,
                            clientId: process.env.TAB_APP_ID,
                            tenantId: (decoded as any).tid,
                            authorityHost: "https://login.microsoftonline.com",
                            applicationIdUri: process.env.TAB_APP_URI,
                            simpleAuthEndpoint: `https://${process.env.PUBLIC_HOSTNAME}`
                        }
                    });
                    const credential = new teamsFx.TeamsUserCredential();
                    const graphClient = teamsFx.createMicrosoftGraphClient(credential, ["Presence.Read"]);
                    const result = await graphClient.api("/me/presence").get();
                    setPresence(" is " + result.availability);
                    microsoftTeams.appInitialization.notifySuccess();
                },
                failureCallback: (message: string) => {
                    setError(message);
                    microsoftTeams.appInitialization.notifyFailure({
                        reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
                        message
                    });
                },
                resources: [process.env.TAB_APP_URI as string]
            });
        } else {
            setEntityId("Not in Microsoft Teams");
        }
    }, [inTeams]);

    useEffect(() => {
        if (context) {
            setEntityId(context.entityId);
        }
    }, [context]);

    /**
     * The render() method to create the UI of the tab
     */
    return (
        <Provider theme={theme}>
            <Flex fill={true} column styles={{
                padding: ".8rem 0 .8rem .5rem"
            }}>
                <Flex.Item>
                    <Header content="This is your tab" />
                </Flex.Item>
                <Flex.Item>
                    <div>
                        <div>
                            <Text content={`Hello ${name}, your presence is ${presence}`} />
                        </div>
                        {error && <div><Text content={`An SSO error occurred ${error}`} /></div>}

                        <div>
                            <Button onClick={() => alert("It worked!")}>A sample button</Button>
                        </div>
                    </div>
                </Flex.Item>
                <Flex.Item styles={{
                    padding: ".8rem 0 .8rem .5rem"
                }}>
                    <Text size="smaller" content="(C) Copyright Wictor Wilén" />
                </Flex.Item>
            </Flex>
        </Provider>
    );
};
