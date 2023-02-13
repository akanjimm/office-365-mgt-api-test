import { useState } from "react";
import { Button, Loader } from "@fluentui/react-northstar";
import { Toaster } from 'react-hot-toast';

import config from "../utils/config";
import { toasterErrorMessage } from '../utils/errorHandlingUtils';

const functionName = config.apiName || "myFunc";

export function AzureFunctions(props) {
  const { codePath, triggerConsent, apiClient } = {
    codePath: `api/${functionName}/index.js`,
    ...props,
  };
  const [apiData, setApiData] = useState(undefined);
  const [isClicked, setIsClicked] = useState(false);

  // Function to handle button click and calling azure functions (API)
  const handleSubmit = async () => {
    setIsClicked(true);
    try {
      const response = await apiClient.get("user");
      setIsClicked(false);
      setApiData(response.data);
    } catch (error) {
      setIsClicked(false);
      let errorMessage = error.response.data.error;
      if (errorMessage.includes("invalid_grant")) {
        triggerConsent(true);
      } else {
        toasterErrorMessage("Failed to retrieve your Microsoft 365 data");
      }
    }
  }

  return (
    <div>
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
      </div>
      <Toaster toastOptions={{ duration: 5000 }} />
    </div>
  );
}
