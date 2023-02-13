import { useContext, useState } from "react";
import { Welcome } from "./helperComponents/Welcome";
import { TeamsFxContext } from "./Context";
import config from "./utils/config";

import { BearerTokenAuthProvider, createApiClient } from "@microsoft/teamsfx";
import { useData } from "@microsoft/teamsfx-react";
import Consent from "./helperComponents/Consent";
import { Loader } from "@fluentui/react-northstar";
import { toasterErrorMessage } from "./utils/errorHandlingUtils";
import { Toaster } from "react-hot-toast";

const functionName = config.apiName || "myFunc";

const showFunction = Boolean(config.apiName);
/* TODO
  * implement checking if a tenant is configured to use the application
*/

export default function Tab() {
  const { themeString } = useContext(TeamsFxContext);
  const [needConsent, setNeedConsent] = useState();

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
    setNeedConsent(booleanValue)
  }

  // call azure functions to check if there is a need to consent to permissions
  const { loading } = useData(async () => {
    try {
      const response = await apiClient.get("consent");
      if (response.data === "True") {
        triggerConsent(false);
      }
      return response.data;
    } catch (error) {
      let errorMessage = error.response.data.error;
      if (errorMessage.includes("invalid_grant")) {
        triggerConsent(true);
      } else {
        toasterErrorMessage("An error occured!")
      }
    }
  });

  return (
    <div className={themeString === "default" ? "" : "dark"}>
      {loading && <Loader />}
      {!loading && <div>{needConsent ? <Consent triggerConsent={triggerConsent}/> : <Welcome triggerConsent={triggerConsent} apiClient={apiClient} />}</div>}
      <Toaster toastOptions={{ duration: 5000 }} />
    </div>
  );
}
