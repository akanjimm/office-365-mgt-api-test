import { useContext, useState } from "react";
import { Button, Loader } from "@fluentui/react-northstar";
import { BearerTokenAuthProvider, createApiClient } from "@microsoft/teamsfx";
import { Toaster } from 'react-hot-toast';

import { TeamsFxContext } from "../Context";
import config from "../utils/config";
import { toasterErrorMessage } from '../utils/errorHandlingUtils';

const functionName = config.apiName || "myFunc";

export function AzureFunctions(props) {
  const { codePath, docsUrl } = {
    codePath: `api/${functionName}/index.js`,
    docsUrl: "https://aka.ms/teamsfx-azure-functions",
    ...props,
  };
  const [apiData, setApiData] = useState(undefined);
  const [isClicked, setIsClicked] = useState(false);

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

  // Function to handle button click and calling azure functions (API)
  const handleSubmit = async () => {
    setIsClicked(true);
    try {
      const response = await apiClient.get(functionName);
      setIsClicked(false);
      setApiData(response.data);
    } catch (error) {
      setIsClicked(false);
      toasterErrorMessage("Failed to retrieve your Microsoft 365 data");
    }
  }

  return (
    <div>
      <h2>Call your Azure Function</h2>
      <p>An Azure Functions app is running. Authorize this app and click below to call it for a response:</p>
      <Button primary content="Call Azure Function" onClick={handleSubmit} />
      {isClicked && !apiData && (
        <pre className="fixed">
          <Loader />
        </pre>
      )}
      {!isClicked && !apiData && <pre className="fixed"></pre>}
      {apiData && <pre className="fixed">{JSON.stringify(apiData, null, 2)}</pre>}
      <h4>How to edit the Azure Function</h4>
      <p>
        See the code in <code>{codePath}</code> to add your business logic.
      </p>
      <Toaster toastOptions={{ duration: 5000 }} />
    </div>
  );
}
