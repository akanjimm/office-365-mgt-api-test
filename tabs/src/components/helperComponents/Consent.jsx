import './Consent.css';
import { useState } from "react";
import { TeamsFx } from '@microsoft/teamsfx';
import { Toaster } from 'react-hot-toast';

import appLogo from '../../images/logo_2_30_x_30.png';
import { toasterErrorMessage } from '../utils/errorHandlingUtils';

const teamsfx = new TeamsFx();
const scope = ["User.Read", "User.ReadBasic.All", "User.Read.All"];

export default function Consent(props) {
    const { triggerConsent } = props;
    const [consentErrorHeader, setConsentErrorHeader] = useState("");
    const [consentErrorText, setConsentErrorText] = useState("");

    // function to show login popup page for user to grant consent to permissions
    const handleLogin = async () => {
        try {
            await teamsfx.login(scope);
            triggerConsent(false); // close consent page by updating parent state
        } catch (err) {
            let message;
            if (err instanceof Error && (err.message?.includes("CancelledByUser") || err.message?.includes("User declined"))) {
                message = "The consent process was cancelled. Please grant consent or contact an administrator in order to use the application.";
                setConsentErrorHeader("Consent Cancelled!");
                setConsentErrorText(message);
            } else if (err instanceof Error && err.message?.includes("browser is blocking the url to open")) {
                message = "The consent process was blocked. Kindly consent in Microsoft Teams and refresh the application here or contact an administrator.";
                setConsentErrorHeader("Consent Blocked!");
                setConsentErrorText(message);
            } else {
                toasterErrorMessage("An error occured!");
            }
        }
    }

    return (
        <div>
            <div className='consent-wrapper'>
                <div className='consent'>
                    <div><img src={appLogo} alt='Employee Lookup Logo' className='logo' /></div>
                    <p className='consent-head'>{(consentErrorHeader) ? consentErrorHeader : "Consent Required!"}</p>
                    <p className='consent-body'>{(consentErrorText) ? consentErrorText : "To continue using the app, some permissions are needed. Kindly click the consent button below to grant consent or contact an administrator."}</p>
                    <div className='consent-button-wrapper'>
                        <button className='consent-button' onClick={handleLogin}>Consent</button>
                    </div>
                </div>
            </div>
            <Toaster toastOptions={{ duration: 5000 }} />
        </div>
    )
}