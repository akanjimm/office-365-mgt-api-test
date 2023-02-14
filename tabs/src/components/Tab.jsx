import { useContext, useState } from "react";
import { Welcome } from "./helperComponents/Welcome";
import { TeamsFxContext } from "./Context";
import config from "./utils/config";

import { BearerTokenAuthProvider, createApiClient } from "@microsoft/teamsfx";
import { useData } from "@microsoft/teamsfx-react";
import Consent from "./helperComponents/Consent";
import { Loader } from "@fluentui/react-northstar";
import * as microsoftTeams from '@microsoft/teams-js';
import { TeamsFx } from '@microsoft/teamsfx';
import { toasterErrorMessage } from "./utils/errorHandlingUtils";
import { Toaster } from "react-hot-toast";
import Configure from "./helperComponents/Configure";

/* TODO
  * implement checking if a tenant is configured to use the application
*/

export default function Tab() {
  const { themeString } = useContext(TeamsFxContext);
  const [isConfigured, setIsConfigured] = useState();
  const [needConsent, setNeedConsent] = useState();

  // stop native loading indicator defined in manifest when app loads
  useData(async () => {
    try {
      await microsoftTeams.app.initialize();
      const context = await microsoftTeams.app.getContext();

      if (Object.values(microsoftTeams.HostName).includes(context.app.host.name)) {
        microsoftTeams.app.notifySuccess();
      }
    } catch (error) {
      microsoftTeams.app.notifyFailure(
        {
          reason: microsoftTeams.app.FailedReason.Timeout,
          message: error
        }
      )
    }
  })

  // check if current tenant is configured to use the application
  // this is applicable if the registered Azure AD app is single tenant
  useData(async () => {
    try {
      let teamsfx = new TeamsFx();
      await teamsfx.getUserInfo();
      setIsConfigured(true);
    } catch (err) {
      if (err.message?.includes("resourceDisabled")) {
        setIsConfigured(false)
      } else if (err.message?.includes("Get SSO token failed")) {
        setIsConfigured(false)
      } else {
        toasterErrorMessage("An error occured!")
      }
    }
  })

  // Create API client
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  if (!teamsUserCredential) {
    // TODO: Replace this with a toaster error popup.
    throw new Error("TeamsFx SDK is not initialized.");
  }
  const apiBaseUrl = config.apiEndpoint + "/api/";
  const apiClient = createApiClient(
    apiBaseUrl,
    new BearerTokenAuthProvider(async () => (await teamsUserCredential.getToken("")).token)
  );

  // function to show or hide consent page when consent is needed here or down in the component tree
  const triggerConsent = (booleanValue) => {
    setNeedConsent(booleanValue);
  }

  // call azure functions consent endpoint to check if there is a need to consent to permissions
  const { loading } = useData(async () => {
    try {
      const response = await apiClient.get("consent");
      if (response.data === "True") {
        triggerConsent(false);
      }
    } catch (error) {
      let errorMessage = error.response.data.error;
      if (errorMessage.includes("invalid_grant")) {
        triggerConsent(true);
      } else {
        toasterErrorMessage("An error occured!");
      }
    }
  });

  return (
    <div className={themeString === "default" ? "" : "dark"}>
      {loading && <Loader />}
      {!isConfigured && !loading && <Configure />}
      {isConfigured && !loading && <div>{needConsent ? <Consent triggerConsent={triggerConsent} /> : <Welcome triggerConsent={triggerConsent} apiClient={apiClient} />}</div>}
      <Toaster toastOptions={{ duration: 5000 }} />
    </div>
  );
}
