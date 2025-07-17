# [ARCHIVED] Excel-Content-Add-in-Humongous-Insurance

**Note:** This sample has been moved to the main Office Add-ins Samples repo at [Excel content add-in: Humongous Insurance](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-content-add-in). The sample here is archived and no longer actively maintained. Security vulnerabilities may exist in the project, or its dependencies. If you plan to reuse or run any code from this repo, be sure to perform appropriate security checks on the code or dependencies first. Do not use this project as the starting point of a production Office Add-in. Always start your production code by using the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) or maintained samples, and follow security best practices as you develop the add-in.

The Humongous Insurance content add-in shows how you can use the new JavaScript API for Microsoft Excel 2016 to create a compelling Excel add-in. This add-in shows how you can embed rich, interactive objects into Office documents. The following figure show the main screens of this add-in.

[![Title: images/image1471975566906.Png](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance/blob/master/images/image1471975566906.Png)](https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance/blob/master/images/image1471975566906.Png)

## Table of Contents

*   [Prerequisites](#prerequisites)
*   [Run the project](#run-the-project)
*   [Additional resources](#additional-resources)https://github.com/OfficeDev/Excel-Content-Add-in-Humongous-Insurance/blob/master/images/image1471975566906.Png

## Prerequisites

You'll need:

*   [Visual Studio 2015](https://www.visualstudio.com/downloads/download-visual-studio-vs.aspx)
*   [Office Developer Tools for Visual Studio](https://www.visualstudio.com/en-us/features/office-tools-vs.aspx)
*   Excel 2016, version 6769.2011 or later

## Run the project

1.  To avoid unnecessary security scans on an archived sample, the **packages.config** file was renamed. To restore the project's dependencies, rename **packages-archive.config** to **packages.config**.
1.  Copy the project to a local folder. Ensure that the file path is not too long, otherwise you might run into an error in Visual Studio when it tries to install the NuGet packages necessary for the project.
1.  Then open the `HumongousInsuranceAgentsAddin.sln` in Visual Studio.
1.  Set the startup document to the included Excel spreadsheet, HumongousInsuranceAgentsAddin.xls.
1.  Press F5 to build and deploy the sample add-in. Excel launches and displays the add-in in the spreadsheet. You can play with the add-in by interacting with the controls.

## Additional resources

*   [Office Dev Center](http://dev.office.com/)

## Copyright

Copyright (c) 2016 Microsoft. All rights reserved.


This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
