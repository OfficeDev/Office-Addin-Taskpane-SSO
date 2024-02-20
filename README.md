# Office-Addin-TaskPane-SSO

This repository contains the source code used by the [Yo Office generator](https://github.com/OfficeDev/generator-office) when you create a new Office Add-in that appears in the task pane. You can also use this repository as a sample to base your own project from if you choose not to use the generator.

## TypeScript

This template is written using [TypeScript](http://www.typescriptlang.org/). For the JavaScript version of this template, go to [Office-Addin-TaskPane-SSO-JS](https://github.com/OfficeDev/Office-Addin-TaskPane-SSO-JS).

## Instructions

- Run the following command to configure single-sign on for your add-in project.

```bash
    $ npm run configure-sso
```

- A web browser window will open and prompt you to sign in to Azure. Sign in to Azure using your Microsoft 365 administrator credentials. These credentials will be used to register a new application in Azure and configure the settings required by SSO.

- Run the following command to build the project, start the local web server, and sideload your add-in in the previously selected Office client application.

```bash
    $ npm start
```

    > [!NOTE]
    > Office Add-ins should use HTTPS, not HTTP, even when you are developing. If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provide.

- In the Office client application that opens when you run the previous command (i.e., Excel, Word or PowerPoint), make sure that you're signed in with a user that's a member of the same Microsoft 365 organization as the Microsoft 365 administrator account that you used to connect to Azure while configuring SSO in step 3 of the previous section. Doing so establishes the appropriate conditions for SSO to succeed.

- In the Office client application, choose the Home tab, and then choose the Show Task Pane button in the ribbon to open the add-in task pane. The following image shows this button in Excel.


    ![Example of task pane ribbon button in an Excel worksheet](https://learn.microsoft.com/en-us/office/dev/add-ins/images/excel-quickstart-addin-3b.png?raw=true)

- At the bottom of the task pane, choose the Get My User Profile Information button to initiate the SSO process.

    If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication. This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft account or Microsoft 365 Education or Work account. Choose the Accept button in the dialog window to continue.

- The add-in retrieves profile information for the signed-in user and writes it to the document.

    ![Example of profile information written to an Excel worksheet](https://learn.microsoft.com/en-us/office/dev/add-ins/images/sso-user-profile-info-excel.png?raw=true)

## Debugging

This template supports debugging using any of the following techniques.

- [Use a browser's developer tools](https://learn.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://learn.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://learn.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)

## Questions and comments

We'd love to get your feedback about this sample. You can send your feedback to us in the *Issues* section of this repository.

Questions about Office Add-ins development in general should be posted to [Microsoft Q&A](https://learn.microsoft.com/answers/questions/185087/questions-about-office-add-ins.html). If your question is about the Office JavaScript APIs, make sure it's tagged with [office-js-dev].

## Join the Microsoft 365 Developer Program

Join the [Microsoft 365 Developer Program](https://aka.ms/m365devprogram) to get resources and information to help you build solutions for the Microsoft 365 platform, including recommendations tailored to your areas of interest.

You might also qualify for a free developer subscription that's renewable for 90 days and comes configured with sample data; for details, see the [FAQ](https://learn.microsoft.com/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-).

## Additional resources

- [Office Add-ins documentation](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- More Office Add-ins samples at [OfficeDev on Github](https://github.com/officedev)

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Copyright

Copyright (c) 2021 Microsoft Corporation. All rights reserved.
