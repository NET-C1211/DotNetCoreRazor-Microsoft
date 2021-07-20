## .NET Core Razor Pages with Microsoft Graph

1. Ensure you have .NET 5+ installed on your machine. You can download and install it from the following link:

    https://dot.net

1. Create a Microsoft 365 developer tenant if you don't already have one:

    https://developer.microsoft.com/en-us/microsoft-365/dev-program

    You can view a video that covers key tips here:

    https://www.youtube.com/watch?v=DhhpJ1UjbJ0

1. Register a new app in Azure Active Directory:

    - Login to the Azure Portal.
    - Select Azure Active Directory.
    - Select `App registrations` in the `Manage` section.
    - Select `New registration` in the toolbar.
    - Give the app a name.
    - Select `Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)` in the `Supported account types`.
    - In the Redirect URI section select `Web` and enter the following URL:

        https://localhost:5001

    - After the app registration is created, note the `clientId` value shown (you'll use it later) .
    - Click the `Authentication` option on the left.
    - Add the following URL into the `Web` section's `Redirect URIs`:

        https://localhost:5001/signin-oidc

    - Check the `ID tokens` checkbox.
    - Save your changes.
    - Click `Certificates & secrets` and create a new client secret. Ensure that you copy and store the secret somewhere since this is the only time you'll be able to access it. You'll need it in the next step.

1. Perform the following steps in `appsettings.json`:
    - Update the `Domain` property value with your Microsoft 365 tenant name.
    - Update the `ClientId` property value with the `clientID` that was created when you did the app registration steps above.
    - Update the `ClientSecret` property value with the secret that you created earlier during the app registration steps above.
1. Run `dotnet restore`
1. Run `dotnet build`
1. Run `dotnet run`

1. Once the app is running, navigate to https://localhost:5001 and login using one of your Microsoft 365 tenant users.
1. Once you're logged in you should see your user name displayed. Click on the menu items at the top to view the user's email, calendar (you may need to add calendar items for the user), and files.

NOTE: If you get an SSL certificate error, you can generate a dev certificate for your machine using the following command:

```dotnet dev-certs https -t```