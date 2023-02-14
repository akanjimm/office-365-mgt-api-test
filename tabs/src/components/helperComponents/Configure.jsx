import './Configure.css'
import appLogo from '../../images/logo_2_30_x_30.png';

export default function Configure() {

    return (
        <div>
            <div className='install-error-wrapper'>
                <div className='install-error'>
                    <div><img src={appLogo} alt='Employee Lookup Logo' className='logo' /></div>
                    <p className='install-error-head'>Almost there! Just a couple configurations required.</p>
                    <p className='install-error-body'>You are seeing this page because your tenant is not properly configured to run the application.</p>
                    <div className='error-decision'>
                        <div>
                            <p>Please contact support using this email</p>
                            <p className='contact'>be@relianceinfosystems.com</p>
                        </div>
                        <div> Or</div>
                        <div>
                            <p>Visit our product page to learn more</p>
                            <div><a className='contact link' href="{{state.fx-resource-frontend-hosting.endpoint}}" target="_blank" rel="noopener noreferrer">Click to visit our product page</a></div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    )
}