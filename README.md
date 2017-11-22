# About FederatedFBTroubleshooter
The idea behind the FederatedFBTroubleshooter is to create a single, open-source script that can use standard Exchange PowerShell cmdlets to determine and resolve common Federation/FB issues that we see in Support.  In the initial release, the script's only functionality is to reset WSSecurity on the EWS and Autodiscover virtual directories, per this article: https://support.microsoft.com/en-us/help/2752387/users-from-a-federated-organization-cannot-see-the-free-busy-informati

Additional features and issue detection algorithms will be added in the near future.

# How To Run
The script MUST be run via Windows Powershell - *NOT* Exchange Management Shell - on the latest version of Exchange.  For example, if you have Exchange 2010, 2013 and 2016, run from an Exchange 2016 server.

To run, type the following in the Exchange Management Shell:

*.\FederatedFBTroubleshooter.ps1*

After detecting the servers in your environment, you will be given 5 choices:

1) Reset WSSecurity on ALL servers
2) Reset WSSecurity on ALL servers in the server's AD site
3) Reset WSSecurity on ALL servers of a specific version
4) Reset WSSecurity on ALL servers of a specific version in the server's AD site
5) Reset WSSecurity on a specific server

*Note, if you have multiple geographical sites, running the script against ALL servers may take a very long time.  It is recommended, instead, that the script be run from each individual site*

# Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.microsoft.com.

When you submit a pull request, a CLA-bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., label, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
