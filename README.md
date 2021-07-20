## .NET Core Razor Pages with Microsoft Graph

1. Create a Microsoft 365 developer tenant if you don't already have one:

    https://developer.microsoft.com/en-us/microsoft-365/dev-program

1. Register a new app in Azure Active Directory:

    - Login to the Azure Portal
    - Select Azure Active Directory
    - Select `App registrations` in the `Manage` section
    - Select `New registration` in the toolbar
    - Give the app a name
    - Iin the Redirect URI section select `Web` and enter the following URL:

        https://localhost:5001/signin-oidc

    - After registering the app, click the Authentication option on the left
    - Check the `Access tokens` and `ID tokens` checkboxes
    - Save your changes
    - Click `Certificates & secrets` and create a new client secret

1. Update `appsettings.json` with your AAD clientID and client secret values.
1. Run `dotnet restore`
1. Run `dotnet build`
1. Run `dotnet run`

1. Once the app is running, navigate to https://localhost:5001 and login using one of your Microsoft 365 tenant users.

NOTE: If you get an SSL certificate error, you can generate a dev certificate for your machine using the following command:

```dotnet dev-certs https -t```