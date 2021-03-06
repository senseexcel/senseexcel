# Sense Excel Detailed User Guide

![SE Cover](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Cover.PNG)

- [Sense Excel Detailed User Guide](#sense-excel-detailed-user-guide)
  * [1. Disclaimer](#1-disclaimer)
  * [2. Introduction](#2-introduction)
    + [2.1 Who We Are](#21-who-we-are)
    + [2.2 Why We Do What We Do](#22-why-we-do-what-we-do)
  * [3. What You Can Expect from Sense Excel](#3-what-you-can-expect-from-sense-excel)
    + [Sense Excel can/does not...](#sense-excel-can-does-not)
    + [Sense Excel can/does.....](#sense-excel-can-does)
  * [4.  Prerequisites and System Setup](#4--prerequisites-and-system-setup)
    + [4.1 System Requirements](#41-system-requirements)
    + [4.2 Qlik Sense Server Configuration - NOT REQUIRED FOR USE WITH QLIK SENSE DESKTOP ONLY](#42-qlik-sense-server-configuration---not-required-for-use-with-qlik-sense-desktop-only)
      - [4.2.1 Add a Security rule to your Qlik Sense Server.](#421-add-a-security-rule-to-your-qlik-sense-server)
      - [IDENTIFICATION](#identification)
      - [BASIC](#basic)
      - [ADVANCED](#advanced)
      - [4.2.2  Create a Content Library on the Qlik Sense Server.](#422--create-a-content-library-on-the-qlik-sense-server)
      - [4.2.3 Create and Upload "license.txt" File.](#423-create-and-upload--licensetxt--file)
      - [4.2.4  View / Update an Existing "license.txt" File.](#424--view---update-an-existing--licensetxt--file)
  * [5. Installation Guide](#5-installation-guide)
    + [5.1 Download Software](#51-download-software)
    + [5.2 Installer Program](#52-installer-program)
  * [6.  Explaining the User Interface](#6--explaining-the-user-interface)
    + [6.1 "Connection"](#61--connection-)
      - [6.1.1 Manual Connection Creation](#611-manual-connection-creation)
      - [6.1.2 Create a Connection Using a Wizard](#612-create-a-connection-using-a-wizard)
      - [6.1.3 Edit Existing Connection](#613-edit-existing-connection)
    + [6.2 "Hub"](#62--hub-)
    + [6.3 “Load data” and “Data Load Editor”](#63--load-data--and--data-load-editor-)
    + [6.4 “Table”](#64--table-)
      - [Select Table Object to Import:](#select-table-object-to-import-)
      - [Pivot Tables](#pivot-tables)
      - [Table Definition Interface](#table-definition-interface)
    + [6.5 "Bookmarks"](#65--bookmarks-)
      - [6.5.1 "Default Bookmark"](#651--default-bookmark-)
    + [6.6 “Selections tool”](#66--selections-tool-)
    + [6.7 Sense Filter Toolbar Navigation](#67-sense-filter-toolbar-navigation)
    + [6.8 Alternate States](#68-alternate-states)
    + [6.9 "Report Preview"](#69--report-preview-)
    + [6.10 "Settings"](#610--settings-)
      - [Language:](#language-)
      - [Open Info Tab:](#open-info-tab-)
      - [Settings:](#settings-)
      - [Support:](#support-)
      - [Versions:](#versions-)
    + [6.11 "About"](#611--about-)
    + [6.12 Sheet Loop](#612-sheet-loop)
    + [6.13 Export Field/Dimension Values](#613-export-field-dimension-values)
  * [7. Using the Sense Excel Demo Report](#7-using-the-sense-excel-demo-report)
  * [8. Create Your Own Reports](#8-create-your-own-reports)
    + [8.1 Using the “Table” import button](#81-using-the--table--import-button)
    + [8.2 Fixed Cell Report Definition](#82-fixed-cell-report-definition)
    + [8.3 Sense Excel Formulas and Syntax](#83-sense-excel-formulas-and-syntax)
      - [8.3.1 SenseConnected](#831-senseconnected)
      - [8.3.2 SenseFilter](#832-sensefilter)
      - [8.3.3 SenseVariable](#833-sensevariable)
      - [8.3.4 SenseEV](#834-senseev)
  * [9. Seven Steps to Your First Report in Sense Excel](#9-seven-steps-to-your-first-report-in-sense-excel)
    + [Step 1: Connect to Qlik Sense & select an App](#step-1--connect-to-qlik-sense---select-an-app)
    + [Step 2: Press the "Table" button](#step-2--press-the--table--button)
    + [Step 3: Press the Check button to create the table in your Excel workbook.](#step-3--press-the-check-button-to-create-the-table-in-your-excel-workbook)
    + [Step 4. Edit Table Definition](#step-4-edit-table-definition)
    + [Step 5: Open the Selections tool](#step-5--open-the-selections-tool)
    + [Step 6: Select a Bookmark](#step-6--select-a-bookmark)
    + [Step 7: Create an Interactive Chart or Pivot table](#step-7--create-an-interactive-chart-or-pivot-table)
  * [10. Terminal Server installtion](#10-how-to-install-on-a-terminal-server)
    + [Check: Access to the add-in](#check-access-to-the-add-in)
    + [Check: The Office Version](#check-officeversions)
    + [Important NOTE](#important-note)
    + [Update a version](#update-a-version)
  * [11. Contact Information](#11-contact-information)
    + [Contact Support:](#contact-support-)
    + [Contact Sales](#contact-sales)
      - [Europe/German Speaking Countries:](#europe-german-speaking-countries-)
      - [Americas:](#americas-)
      - [Rest of the World:](#rest-of-the-world-)
      - [Address:](#address-)
  * [Appendix 1 - Older Version Installation Instructions](#appendix-1---older-version-installation-instructions)
    + [Appendix 1.1 Extract Download Package](#appendix-11-extract-download-package)
    + [Appendix 1.2 New Sense Excel Installation](#appendix-12-new-sense-excel-installation)
    + [Appendix 1.3 Alternate Installation Technique](#appendix-13-alternate-installation-technique)
    + [Appendix 1.4 Upgrade Sense Excel Installation](#appendix-14-upgrade-sense-excel-installation)
  * [Appendix 2 - Sense EV Example Formulas](#appendix-2---sense-ev-example-formulas)


## 1. Disclaimer 

This document will cover how to get Sense Excel up and running on your system, what to expect and ways the product can be used. This document is written to provide quick answers and will be continuously updated. Not all options or possibilities of Sense Excel are covered within the scope of this document but will be complimented by best practices and use cases documented elsewhere in this repository. 

 AKQUINET SENSE EXCEL LICENSE AGREEMENT 

IMPORTANT: BY ACCEPTING, DOWNLOADING OR USING THIS SOFTWARE, YOU ACCEPT AND AGREE TO THE TERMS OF THIS AKQUINET SENSE EXCEL LICENSE AGREEMENT (“DLA”) AS MAY BE UPDATED FROM TIME TO TIME AND PUBLISHED AT WWW.AKQUINET.COM. BY ACCEPTING THESE TERMS, OR USING THE SOFTWARE AS AN EMPLOYEE, CONTRACTOR OR AGENT ON BEHALF OF A CORPORATE OR OTHER ENTITY, YOU (“USER”) REPRESENT AND WARRANT THAT YOU HAVE AUTHORITY TO BIND SUCH ENTITY TO THESE TERMS. DIRECT COMPETITORS AND THEIR EMPLOYEES AND AGENTS MAY NOT ACCESS THE SOFTWARE 
WITHOUT PRIOR WRITTEN CONSENT OF AKQUINET. THE SOFTWARE MAY NOT BE USED FOR PURPOSES OF BENCHMARKING, COLLECTING AND PUBLISHING SOFTWARE PERFORMANCE DATA OR ANALYSIS, OR ANY OTHER COMPETITIVE PURPOSES. 
 
NOTICE: THE SOFTWARE CONTAINS FUNCTIONALITY INTENDED TO LIMIT THE DURATION OF ITS USE AND IS INTENDED TO COLLECT CERTAIN USAGE METRICS. THE INSTALLATION OF THIS SOFTWARE WILL INSTALL FILES NECESSARY TO OPERATE THE SOFTWARE ONTO THE USER’S COMPUTER AND OTHER SYSTEM FILES MAY BE INSTALLED OR UPDATED. AS WITH ALL INSTALLATIONS, BACK UP OF THE USER’S HARD DRIVE IS RECOMMENDED BEFORE INSTALLING THE SOFTWARE. 

You can review the rest of this disclaimer by pressing the "About" button on the "SENSE" Ribbon. 


## 2. Introduction 


### 2.1 Who We Are

akquinet AG is a stock corporation headquartered in Hamburg, Germany with over 1000 employees. We offer a wide range of IT and BI implementation services including Enterprise Resource Planning (ERP) systems utilizing Microsoft and SAP, custom ERP development using JAVA as well as managed hosting services. akquinet's Business Intelligence team is located in Jena, Germany is and responsible for development of the Sense Excel product line. 


### 2.2 Why We Do What We Do

We believe that all stakeholders should have easy access to business critical information when and where they need it. Enabling all users to access their data in a familiar interface while, at the same time, fully reaping the benefits of the "single version of truth" aspect of an Enterprise BI Platform is our primary goal.


## 3. What You Can Expect from Sense Excel 

In this chapter we will explain what you can and can't expect from Sense Excel. 
 

### Sense Excel can/does not... 

• …import a complete Qlik Sense Dashboard into Excel

• …import visualization objects from Qlik Sense into Excel

• …change Excel in any way other than extending its reporting capabilities 


### Sense Excel can/does.....

• …provide the same UI (user interface) as Qlik Sense including filter selections and bookmarks for ease of use.

• …provide all Dimensions, Measures, Fields, Bookmarks etc. that are available in the connected Qlik Sense app.

• …allow for creation of data tables in Excel utilizing the familiar workflow of Qlik Sense.

• …allow for a quick download of existing Qlik Sense table objects into Excel.

• …allow editing and reloading of the Qlik Sense application script from directly inside of Excel.

• …allow access to the complete Qlik Sense function library, including Set Analysis, directly within an Excel formula.

• …take full advantage of the best features of Excel including copy/paste and fill down/across to rapidly increase your speed of development.

• …allow for population of Qlik Sense formula and Set Analysis parameters utilizing embedded Excel cell references.

• …allow you take full advantage of all Excel features including charting and pivot tables on data that comes directly from Qlik Sense.

• …allow for access to multiple apps (.qvf files) within a single Excel workbook.  

To summarize, with Sense Excel you get the full power and flexibility of Microsoft Excel combined with all of the speed, governance and navigation capabilities of Qlik Sense. We are confident that Sense Excel will enhance both your Qlik Sense and Excel experience and provide you with consistent, reliable, real-time data for your reporting efforts.


## 4.  Prerequisites and System Setup


### 4.1 System Requirements 

• TO use Sense Excel, you must be able to access either Qlik Sense Desktop or Qlik Sense Server software.

• Confirm that you have a .NET Framework >= 4.5.1 installed, otherwise download it from https://dotnet.microsoft.com/download and install it. 

• You have Microsoft Excel 2013, 2016, 2019 or Office 365-Desktop (32 or 64 bit) installed.

To find out which Microsoft Office edition you are running within Excel, go to FILE >> ACCOUNT >> "About Excel".  Behind the MSO number, it will display 32-Bit or 64-Bit.

![Excel Version](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Excel-Version.PNG)


### 4.2 Qlik Sense Server Configuration - NOT REQUIRED FOR USE WITH QLIK SENSE DESKTOP ONLY

To use Sense Excel only with Qlik Sense Desktop skip directly to Chapter 5. You DO NOT need to perform the steps in Chapter 4.2.

To use Sense Excel with a Qlik Sense Enterprise installation, you must perform the below configuration steps prior to attempting to connect. Sense Excel license keys are required to use Sense Excel with a Qlik Sense Enterprise installation. 

30-day Sense Excel trial licenses are available for existing Qlik Sense customers.  Please contact an authorized reseller or your appropriate sales contact listed at the end of this document. 

EXCEPTIONS: Qlik Sense Desktop and all installations utilizing a valid Qlik Sense Trial, Internal or Partner keys DO NOT require a separate Sense Excel license key.  


 
#### 4.2.1 Add a Security rule to your Qlik Sense Server. 

QMC > MANAGE RESOURCES > Security Rules > + Create new > SER License

#### IDENTIFICATION
|Setting |Value
|------------------|------------------|
|Name | SER License |
|Create Rule from Template | Unspecified |
|Disabled | Leave Unchecked |
|Description | |

#### BASIC
|Setting |Value|
|---|---|
|Resource Filter | License_* |
|Actions | Read |

#### ADVANCED
Add the below value manually into the Conditions table:

|Setting    |Value             |
|-----------|------------------|
|Conditions | !user.IsAnonymous() |
|Context    | Both in hub and QMC   |

Validate Rule > Add Rule

![SER License Security Rule](https://github.com/senseexcel/senseexcel-reporting/blob/master/docs/Security-Rule-SER-License.PNG)


#### 4.2.2  Create a Content Library on the Qlik Sense Server. 

1. Go to QMC >> Content libraries >> Create New >> Enter the EXACT name “senseexcel” >> Apply.
2. When prompted to create an associated Security Rule select "User" is "Like" * to make Sense Excel available to all licensed Qlik Sense users upon assignment of a Sense Excel license as shown below.


#### 4.2.3 Create and Upload "license.txt" File. 

1. Copy the contents of your organization's Qlik Sense LEF file or trial key into a new text document. 
2. If your organization uses Token licensing, only the contents of your LEF or trial key are required. For Named User licensing strategies please follow the additional steps below.
3. Append the LEF information with EXCEL_NAME; followed by the Qlik Sense "User directory" and "User ID" of the designated Sense Excel users like shown in the example below. 
4. The FROM; TO information is optional.  This allows time limits for the duration of license assignments for your named users. The example below shows a time limited license assignment followed by a reassignment to another user.

![SE License Example LEF](https://github.com/senseexcel/senseexcel/blob/master/images/SE-License-Example-LEF.png)


6. Please confirm that there are no spaces at the end of each line.
7. Save this file with the EXACT file name "license.txt".
8. Upload the "license.txt" file to the senseexcel content library. 
   
 QMC >> Content Libraries >> senseexcel >> Contents >> Upload >> Select File “license.txt”. 

   
#### 4.2.4  View / Update an Existing "license.txt" File. 

Once installed, the "license.txt" in the "senseexcel" Content library is used to manage the assignment of Sense Excel named user licenses. It can be viewed, updated and overwritten by following the steps below.  

1. Go to QMC >> Start >> Content libraries >> senseexcel >> Associated Items >> Contents

2. Copy the contents of the "URL path" and paste it immediately after your Qlik Sense Server Url in a new browser tab/window like the below example

e.g. https:// your.qlikserver.com/content/sensexcel/license.txt

![SE Check Existing License](https://github.com/senseexcel/senseexcel/blob/master/images/SE-License-Check-Existing-License.png)

3.  Entering this Url will display the contents of an existing "license.txt" file. to make changes, copy the contents from the screen into a new text file, make your desired changes, save the document as "license.txt" and upload/overwrite the existing license.txt file in the "senseexcel" content library.

4. You also can use the Sense Excel Web Management Console to manage your Sense Excel licenses.  Please go to https://www.senseexcel.com to learn more.

IMPORTANT: The names of the "senseexcel" Content library and "license.txt" file need match EXACTLY and ARE case sensitive. 


## 5. Installation Guide

Version 4.0 includes a new Install-Updater program.

If you have Sense Excel software but no active Internet connection or have any reason to need access to older Sense Excel versions and installation workflows, follow the manual installation steps described in Appendix 1 at the end of this document.


### 5.1 Download Software

You can acquire the Sense Excel Install-Updater program from the link below:

https://m.sense2go.net/asset/63:sense-excel-install--updater.  

Clicking or pasting the link in a browser will download to your designated Downloads folder.  

BEST PRACTICE:  Move the SE-Install-Updater.exe file from your Downloads folder to a new folder in your Documents or Desktop called Sense Excel.  The log file generated by this process will remain in the same folder that the program is started in.

The Sense Excel Install-Updater is also available within a Sense-Excel-All-In-One Download package inside of the Sense Excel X.X folder.

![SER Installer Download Folder](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Installer-Download-Folder.png)


### 5.2 Installer Program

1. Double click on the SE-Install-Updater.exe file to run it.  IF you receive and error message, instruct Windows to Ignore the warning and run the program anyway.

2. The first window will prompt you to save any open Excel processes.  Clicking "Next" will terminate all Excel processes and unsaved data will be lost.

![SER Installer Close Windows](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Installer-Close-Windows.png)

3. The next window will bring you into "Normal" mode. This will show you only the most up to data version of Sense Excel available.  


![SER Installer Normal Mode](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Installer-Normal-Mode.png)

Pressing the "Update" button will download that most current version and install it in the Plugin installation path designated.  

The default location is the Microsoft/Addins folder in the Roaming profile of the current user. IF you need to navigate to that location you can do so by typing %appdata% into your Windows Search box or the address bar in File Manager.

![SER Installer Updating](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Installer-Updating.png)

4. Choose Expert Mode by switching the expert mode switch to "On".  Expert mode allows you to access and download Sense Excel Addins for versions of Excel that differ from the one currently running and all available previous Sense Excel versions.

![SER Installer Expert Mode](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Installer-Expert-Mode.png)

To access older versions type "old' in the Room ID field and use the "Online Versions" dropdown to make your selection.

![SER Installer Expert Mode Old Versions](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Installer-Expert-Mode-Old-Versions.png)


## 6.  Explaining the User Interface 

To get the most out of your Sense Excel experience, it is necessary to understand the Sense Excel UI (User Interface). 

The picture below shows the “SENSE” Ribbon as well as the features that make up Sense Excel.

![Toolbar](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar.png)


### 6.1 "Connection"

![Toolbar Connection](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Connection.png)

The "Connection" portion of the Sense Excel Ribbon allows you to create and manage the connections Sense Excel makes to your Qlik Sense environment(s).  Once a connection is configured, use the Sign In and Sign Out buttons as described below.

a. Sign-In : Pressing the sign in button will attempt to connect to the Connection shown in the drop down box.

b. Sign-Out: This will end your active connection to the server. 

There are two different Connection creation techniques available, a manual process as well as a step-by-step Installation Wizard:


#### 6.1.1 Manual Connection Creation

1. Press the "..." next to the drop-down box.
2. Press the New Connection button
3. Fill the fields in the form with the appropriate values as described below.

a. "Connection Name" field.  Enter a unique name for the connection. 

BEST PRACTICE: Choose a Connection name which you can easily associate with the source system and find within a drop-down list. 

b. Ignore Certificate Errors Checkbox. 

Sometimes Qlik Sense is run on a server without a valid security certificate.  Checking this box will ignore https://certificate errors  that are encountered and proceed with the connection.  Unchecking the box will allow the https:// chain to remain intact and provides the most security. The default setting is "checked".

![Connection Edit Property Panel Ignore Certificate Errors](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Connection-Edit-Property-Panel-Ignore-Certificate-Errors.png)

c. Connection Type

Depending on your Qlik Sense authorization strategy and which browser you are using, this parameter can be set to allow Sense Excel to work in any environment.  

![Connection Edit Property Panel Connection Type](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Connection-Edit-Property-Panel-Connection-Type.png)

There are 4 different Connection types available:

a. "Qlik Sense Desktop".  This setting is included in the default Sense Excel installation and points to the local machine using 127.0.0.1 as the IP address.

If do not require use of Qlik Sense Desktop as a data source, this connection can be deleted.  It can be re-created if needed by creating a new connection using the following parameters:  Type: Qlik Sense Desktop and Url: ws: //127.0.0.1:4848.

b. "Sense Server - Current Windows User"

Like Qlik Sense, the default behavior of Sense Excel uses Windows authentication and passes the credentials of current user to Qlik Sense.  This works with Qlik Sense Server implementations on a local machine as well as with Windows based single sign on strategies.

c. "Sense Server - Enter username/password"

This approach is used when you have different credentials for Qlik Sense than your local windows machine.  Signing in using this approach will prompt for Qlik Sense credentials.

d. "Sense Server - Custom authentication (via embedded Browser)"

This technique will open a new, embedded browser window, prompt for credentials and pass them to Qlik Sense via the browser session cookie to log in. Use this approach when logging into a server on different domain or using a third-party single sign on system such as OKTA.

5. Url 

Enter the Url of your target server. https: //your.qlikserver.com.  This can be easily copied from the address bar of a browser connected to your Qlik Sense environment. Do not include /hub in this address.

6. 3rd Party Sigle Sign On

If using OKTA or similar, choose "Sense Server - Custom authentication (via embedded Browser)" as your "Type" and append your Url as follows:  https:// your.qlikserver.com/okta

7. "Session cookie header name" - Alternate Proxy Server

Sense Excel supports defining connections that point to different virtual proxy servers in a distributed Qlik Sense environment.  

A virtual proxy name that begins with X-Qlik-Session can be used by appending the name of the virtual proxy directly to the Url as follows: 

https:// your.qlikserver.com/development

If your virtual proxy has a header name that begins with anything other than "X-Qlik-Session" enter it in the "Session cookie header name" field as follows: "Otherheadername"

8. Save your new Connection by pressing the Green Check Button.
 

#### 6.1.2 Create a Connection Using a Wizard

You can also use use a step by step Wizard process to create a new Connection by following the steps below.

In the "SENSE" ribbon, use the Connections dropdown box and select "Create new connection to server" at the bottom of the list.
 
![Connection Edit New Connection Property Panel Wizard](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Connection-New-Connection-Property-Panel-Wizard.png)
 
This will open the Wizard process in the Sense Excel Property Panel on the right side of your screen, prompt you to input required values and show detailed instructions.
 
1. Connection Name:  

Please enter a unique name for the connection you would like to create.  

BEST PRACTICE: Choose a Connection name that will clearly identify the underlying source and be easy to find in a drop down list. 
  
![Connection Edit New Connection Property Panel Wizard Step 1 Connection Name](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Connection-New-Connection-Property-Panel-Wizard-Step-1-Connection-Name.png)

2. Enter Url.  Type this entry in manually or copy it from the address bar of a browser connected to Qlik Sense.

https:// qlik.yourcompany.com
    
  ![Connection Edit New Connection Property Panel Wizard Step 2 Connection Url](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Connection-New-Connection-Property-Panel-Wizard-Step-2-Connection-Url.png)
  
Hitting the "Next" button will first attempt the default strategy of authenticating using the Windows credentials of the current user.  If this approach works, you will be connected and displayed a "Success" message.   
  
3. If your current Windows credentials won't connect, you will be prompted for Step 3.  The underlying Connection type will be changed to "Sense Server - Custom authentication (via embedded Browser)" and prompt you for the appropriate credentials.

![Connection Edit New Connection Property Panel Wizard Step 3 Custom Authentication](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Connection-New-Connection-Property-Panel-Wizard-Step-3-Custom-Authentication.png)

Upon successful connection to the Qlik server the "Success" message shown below will be displayed.

![Connection Edit New Connection Property Panel Wizard Step 4 Success](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Connection-New-Connection-Property-Panel-Wizard-Step-4-Success.png)

If you complete all steps of the Wizard and your Connection attempt is unsuccessful, you will need to begin the Wizard process again with different parameter values.  Confirm that your Url, user name and password are correct and check with your Qlik Sense administrator for any Virtual Proxy or Header settings that may have been customized in your environment.

To exit the Sense Excel Property Panel press the grey "X" at the top right of the window.


#### 6.1.3 Edit Existing Connection

Once you have created Connections using either technique you can edit or delete them by doing the following:

1. Press "Sign Out" to close any active connections.
2. Press the "..." next to the Connections drop down box.
3. To Delete, select the desired connection and press the "Trash" icon.
4. To Edit, select the desired connection, press the "Pencil" icon and make your edits. 
5. Press the Green Check button to save your changes.
6. Press the Grey "X" to exit the Sense Excel property panel.



### 6.2 "Hub"

![Open Hub](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Open-Hub.png)

This button allows you to select a Qlik Sense App and access the Dimensions, Measures, Fields and underlying Object data and definitions associated with it.  By default, a single workbook is associated with a single Qlik Sense App.  Just like within Qlik Sense itself, all bookmarks, filter selections etc. will apply and act upon all Qlik Sense related data within the workbook. 

It is possible to cause specific data elements to not respond to filters/bookmark selections. Utilizing an Altered State (see Section 6.7), Table column definition or SenseEV() (see section 8.3) formula you can specify a different behavior or source App.

![Open Hub Dialogue](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Open-Hub-Dialogue.png)


### 6.3 “Load data” and “Data Load Editor”

![Load Data](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Load-Data.png)
 
These buttons allow you, with proper credentials, to edit and execute the load script associated with your application.

Sense Excel provides the ability to edit, save and reload Qlik Sense application scripts directly from Excel. This feature can be limited to administrators and certain report developers (power users) to maintain adequate security and consistency. 

Click the "Data Load Editor" Button to edit your script, make your desired changes.

Click the “Load data” button once to accept and save and changes to your App.

Clicking the "Load data" button once will also reload and recalculate all connected elements of your report.

Click the "Load Data" button a second time to execute script reload in the background. 

Notes: You will not see a reload dialog, but the updated results will be automatically reflected in your report after reloading. 

In order to use the "Data Load Editor" and the “Load Data” options, you need to be able to reload the application you wish to edit. The source app for the provided Executive Dashboard example report can't be reloaded and therefore can't be used for testing and understanding these features.

![Data load editor](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Data-load-editor.png)


### 6.4 “Table” 

![Table](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Table.png)

This button will activate the Table Property Panel which allows you to import an existing Table from your app Qlik Sense app directly into an Excel table or create a new one.

Use the Dropdown to Select an exisiting Table Object from Qlik Sense you would like to import.  


#### Select Table Object to Import:

![Import Table](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Table-Property-Panel-Import-Table.png)

Once a Table Object to import is selected, all Columns within that table will be placed into the Data section of the Sense Excel Table Property Panel.

This feature allows ense Excel access to Custom Formulas a Qlik Sense App developer may have only used to define Columns within a  Table Object but have not added as data model Fields or defined as Master Items.

![Import Table With Formulas](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Table-Property-Panel-Import-Table-With-Formulas.PNG)


#### Pivot Tables

Swithching "On" the Pivot Table setting will allow you to either define a new pivot table or Import and existing Pivot Table Object into Excel.


![Import Pivot Table](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Table-Property-Panel-Import-Pivot-Table.png)

Sense Excel Pivot Table definitions require use of one Measure object and allows the use of multiple Row objects. There is a limit of a single Column object.  

If you require a Pivot Table utilizing multiple Column objects, you can import a regular Sense Excel Table and utilize the native Excel Pivot Table functionality.

![Pivot Table](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Table-Property-Panel-Pivot-Table.png)
 

#### Table Definition Interface
 
Once Table Columns are chosen utilizing "Select Table Object to Import", "Add Column" or "Fast Add", a user can then reorder, delete or change properties prior to Loading or Reloading the Table.  The user can also define Custom sorting or additional parameters like Null Supporession and Alternate States.

See the example below showing the use of a Formula to define a Dynamic Background Color.

![Column Edit](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Table-Property-Panel-Column-Edit.png)

For more on this topic, see section 8.3 below.

 
### 6.5 "Bookmarks" 

 ![Bookmarks](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Bookmarks.png)

This option allows you to access existing Bookmarks defined in your Qlik Sense app or save your current filter selections as a new Bookmark that can be reused later. Bookmarks created in Sense Excel are not re-imported into the connected Qlik Sense app.

The Bookmarks command will open a window in the Sense Excel info tab.  You can a create a new Bookmark by using the "Create New Bookmark" button as well as edit or delete any existing bookmarks by pressing the Pencil or Trash Can icons respectively. 

![Bookmarks Info Tab](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Bookmarks-Info-Tab.png)



#### 6.5.1 "Default Bookmark"

Sense Excel brings all requested data into Excel on the client machine. Setting a Default Bookmark can be beneficial when connecting to large source data sets and to minimize processing and network traffic.  

You can set a default Bookmark by checking the "App Default" box shown below.  Once a default Bookmark is defined, all Sense Excel users will have the default Bookmark and associated filter selections automatically applied upon app opening.  
 
![Bookmarks Define](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Bookmarks-Define.png)


### 6.6 “Selections tool” 

![Selection Tool Button](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Selection-Tool-Button.png)

This button will show or hide the Sense Filter Bar where all filter selections will be displayed. 

Once this toolbar is activated, you can use the square box on the right side to pull up all of your available dimensions as list boxes from which to select values for as filters.

![Selection Tool Activate](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Selection-Tool-Activate.png)

Once inside of this interface, you can also use the Search function to reduce the displayed elements to those relevant to your query.

![Selection Tool Search](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Selection-Tool-Search.png)
 
![Selection Tool Search Value](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Selection-Tool-Search-Value.png)
 
![Selection Tool List Box Select](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Selection-Tool-List-Box-Select.png)

![Selection Tool Sense Filter Bar](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Selection-Tool-Filters.png)
  
  
### 6.7 Sense Filter Toolbar Navigation

![Selections Tool Toolbar Navigation](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Selections-Tool-Toolbar-Navigation.png)


The buttons here allow you to manipulate your selections e.g. “Step back”, “Step forward” and “Clear all selections. Furthermore, when a filter is set, you can change the values of your selection by clicking on it and using the Check and X buttons to select and deselect filter values respectively.

### 6.8 Alternate States

Sense Excel can support the creation and use of the Alternate States functionality in Qlik Sense.  Alternate States can be applied to Sense Excel tables causing them to respond only to the filters and filter selections within an Alternate State definition an not to changes in filter selections or bookmark application

To Apply an Existing Alternate State:

1. Open the Sense Excel Selection tool by pressing the Selection Tool button in the "SENSE" Ribbon and pressing the Universal Selector button.
2. Drop down the Alternate States box on the top right of the screen.
3. Select the existing Alternate State you desire.  The filters and filter selections will be applied and displayed immediately.

![Toolbar Selection Tool Altered States](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Selection-Tool-Altered-States-New.png)

To Define a new Alternate State:

1. Open the Sense Excel Selection tool by pressing the Selection Tool button in the Sense Excel Ribbon and pressing the Global Selector button (a Square icon with an arrow thru it on the right side of the Sense Excel Selection tool).
2. Choose the Dimension and Field Values you want to filter and make your filter selections within them.
3. Drop Down the Alternate States box on the top right of the screen.
4. Press the "..." icon
5. Give a name to your new Alternate State
6. Press the Green Check to save.
7. Drop down and choose the Altered State you would like to apply.

![Toolbar Selection Tool Alternate States New](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Selection-Tool-Altered-States-New.png)

An Altered State can be applied to any table defined or imported into Sense Excel.

![Toolbar Selection Tool Alternate States Table](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Selection-Tool-Altered-States-Table.png)

1. Edit an existing table definition by selecting the table and pressing the "Table" button in the Sense Excel Ribbon or Right Mouse Click within the Table and select "Edit Table" to enter the Sense Excel Table Property Panel.
2. If creating a new table, use Table Import, Add Column or Fast Add to place the desired Columns into the Sense Table Property Panel.
3. Expand the "Appearance" section within the Sense Excel Property panel.
4. Choose the Alternate State you would like to apply.
5. Press the Green Check to save your table definition and import your data.

Once an Alternate State is applied to a table definition, the filters and filter selections specified in the Alternate State will apply to that table alone. Bookmarks and/or filter selections made with the Sense Selection tool or Global Selector will be ignored. 


### 6.9 "Report Preview"

![Toolbar Report Preview Button](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Report-Preview.PNG)


The "Report Preview" button will only be enabled if Sense Excel Reporting is installed and running in your Qlik Sense deployment.


![Toolbar Report Preview Show Dialogue](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Report-Preview-Show-Dialogue.png)


![Toolbar Report Preview Property Panel](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Report-Preview-Property-Panel.png)

Executing the "Report Preview" button prompts the Sense Excel Reporting engine to execute your active Sense Excel report including all defined "Sheet Loops" (see section 6.11 for more information) then download a pdf version of the report to your client machine.


### 6.10 "Settings"

![Toolbar Settings](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Settings.png).

Pressing the "Wheel" icon will take you directly to the Sense Excel Property Panel Settings Menu.  Press the Triangle Icon to expand the menu.


#### Language:

"Language" allows you to choose your preferred language. 

![Toolbar Settings Language](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Settings-Language.png)
 

If your preferred language is not available, please send an email to support@qlik2go.net to request including it in the product.
 

![Toolbar Settings Language Expanded](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Settings-Language-Expanded.png)


#### Open Info Tab:

"Open Info Tab" will activate the Sense Excel Property Panel (Info Tab) and display all active processes.

![Toolbar Settings Open Info Tab](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Settings-Info-Tab-Property-Panel.png)


#### Settings:

![Toolbar Settings Settings](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Settings-Settings.png)

Chosing "Settings" from this menu will activate the Sense Excel Settings Panel.  This interface allows a number of behaviors to be enabled or disabled as well as parameters values set on a local or global (your Sense Excel installation) basis.

![Toolbar Settings Property Panel](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Settings-Property-Panel.png)

###### Apply selections while selections tool is open	
While the selections tool is open, changing a selection has no effect to Tables/SenseEV's. After the selections tool is closed, all changes will be applied in one step. So the selections tool is faster.	

default: true
	
###### Apply selections while a listbox is open 	
While a listbox of a selectionitem is open, changing a selection has no effect to Tables/SenseEV's.After the listbox is closed, all changes will be applied in one step. So the listbox is faster.	

default: true
	
###### Open app in a shared session	
If you open an app in your Sense Server/Desktop which is also opend in Senseexcel, then all selections made in Senseexcel will be applied in Senseserver/Desktop and vice versa.	

default: true
	
###### Show UserInfo's when Senseexcel is busy	
…	
default: true
 
###### Show UserInfo for SenseEV Calculations
While calculating SenseEV Formulas a Panel with the progress is opened.	

default: true
	
###### Show Enigma Logging	
Enigma logs Qlik Server communication.  This results in a larger logfile.  In normal cases this is not nessecary.	

default: false
	
###### Keep Qlik session alive while Excel is open
"The Qlik Session times out after a certain amount of time.
In order to continue working with Qlik, this will require a reconnect. When this Option is enabled, Senseexcel periodically sends messages to Qlik to kep the session in use and alive."	

default: true
	
###### Show UserInfo's in english (Support Only)
…	

default: false
	
###### Check for remainig excelprocess	
Check if after closing excel a 'ghost' process remains and close it.	

default: true
	
###### Open the Hub after a successful connect	true
…	

default: true

###### Show Warnmessage when Tabledata is getting truncated	
…	

default: false
	
###### Time in ms after a app-open proccess is canceled	
If an open app proccess has not been finished in this time it will be canceled. Adjusting this Time is usefull on slow machines or slow network connection	

default: 5000
	
###### Time in ms after a connection-open proccess is canceled	
If an open connection process has not been finished in this time it will be canceled. Adjusting this Time is usefull on slow machines or slow network connection	

default: 5000
	
###### Timeout for detecting that a Qlik session is expired	
…	

default: 3000
	
###### Max Number of parallel SenseEV Calculations	
"Recommended is 200. 
A higher value leads to faster SenseEV-calculation but can cause problems on Servers because of to many Requests fired in a short amount of time."	

default: 200
	
###### Time in milliseconds after check for a 'ghost' process is started 	
Check if after closing excel a 'ghost' process remains and close it.	

default: 5000


#### Support:

The “Support” option will open your default e-mail client and create a message with an attached log file addressed to our support-team.  Sometimes this process might be hidden behind other active windows. Minimize your active window to find the dialogue box seen below.  

![Toolbar Settings Support Email](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Settings-Support-Email.png)

Press the "Allow" button to allow Sense Excel to open your default email client and automatically generate a message to Sense Excel support with the information needed to identify you and your configuration.
 
![Toolbar Settings Support Email Message](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Settings-Support-Email-Message.png)

#### Versions:

![Toolbar Settings Versions](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Settings-Versions.png)
 

Selecting "Versions" will activate the Versions window in the Sense Excel Property Panel (Info Tab)

![Toolbar Settings Versions Property Panel](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Settings-Versions-Property-Panel.png)

Once in this interface, available versions can be selected and installed.  Entering a code in the "RoomID" box will allow you to access additional versions made available by the Development Team for preview or testing purposes.


### 6.11 "About" 

 ![Toolbar About](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-About.png)

Clicking on the question mark brings up an additional window with the Sense Excel version, add-in location and disclaimer information. 

 ![Toolbar About Version](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-About-Version.png)
 

### 6.12 Sheet Loop

Sense Excel offers a function called "Sheet Loop" for use with reporting content authored for the On-Demand and Distribution capabilities of Sense Excel Reporting.  
 
When a Sheet Loop is applied, the Sense Excel Reporting Engine will create a discreet Excel worksheet tab filtered for each member of the field/dimension specified.
 
To apply a sheet loop, right mouse click on the worksheet you would like to include it within.
 
![Sheet Loop Open](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Feature-Shootloop-Open.png)
 
Use the drop-down box to choose the Field/Dimension you would like the loop applied to. Once the dimension/field is selected, a dialogue box will show the syntax created by the selection.

![Sheet Choose Dimension-1](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Feature-Sheetloop-Choose-Dimension-1.png) 

You also need to specify whether or not you would like to export the root node and if you would like a dynamic name applied to the worksheets other than the dimension values themselves.

![Sheet Choose Dimension-2](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Feature-Sheetloop-Choose-Dimension-2.png)

Export Root Node:

If you were to specify a Sheet Loop for a field/dimension called "Year" which includes values "2017" "2018" and "2019", your output workbook would have the below worksheets included:

With Export Root Node Checked/True:
 
Worksheet 1: 2017 

Worksheet 2: 2018 

Worksheet 3: 2019
 
With Export Root Node Unchecked/False:
 
Worksheet1: 2018 

Worksheet2: 2019

The default value of Export Root Node is Checked/true.
 

Sheet Name:

The default setting is to name the individual worksheets with the values of the field/dimension specified in the Sheet Loop.  

Having "Sheet Name" unchecked or the "Sheet Name" formula empty would name the worksheets "2017", "2018" and "2019" respectively.

You can also use a Qlik Sense formula to dynamically name the worksheet by checking the "Sheet Name" box and if the formula output meets the following criteria:

1. The formula output values do not exceed 30 characters. This is an Excel limitation and will be truncate all characters above the max of 30.

2. The formula output values do not include any special characters not allowed in an Excel worksheet name such as  ?, /, (, ), ]  etc. 

Example: The formula below;

=only([Fiscal Year])&'|'&'Report'

would create and name a worksheet for each possible value of [Fiscal Year] concatenated with "|Report" as the engine makes its dynamic selections for [Fiscal Year] like below:
 
 2017|Report 
 
 2018|Report
 
 2019|Report
 


### 6.13 Export Field/Dimension Values

Sense Excel gives you the ability to quickly export the values within a field or dimension to a list in Excel.  This can be very useful and reduce developement time considerably especially when used in conjunction with SenseEV formulas utilizing Set Analysis parameters populated with relative Excel formulas references.

Acces the list box mode by either clicking on an existing filter box in the Sense Selection tool or alternatively using the Global selector to choose from all available Dimensions and Fields 

![Feature Export Field Values](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Feature-Export-Field-Values.png) 

Press the select data button to copy the values in Field or Dimnnsion of the displayed List Box.

![Feature Export Field Values Export Data](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Feature-Export-Field-Values-Export-Data.png)

Example:  You have a list box with 5 values.  Value 1, Value 2, Value 3, Value 4, Value 5.  

With No values selected, pressing "Export data" will copy the list of values to the clipboard with a vertical orientation.  Pasting in Excel will give you the following:

Value 1

Value 2

Value 3

Value 4

Value 5

If you have values Selected (green): Value 2 and Value 4, Pressing the "Export data" button pasting the list into Excel will do so, again with a vertical orientation and the selected values displaying at the top of the list 

Value 2

Value 4

Value 1

Value 3

Value 5

Alternatively, you can use the "Add" button instead of "Export data" to copy and paste the values to Excel in a Transposed (oriented horizontally) Fashion as below:

![Feature Export Field Values Export Data Transpose](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Feature-Export-Field-Values-Export-Data-Transpose.png)

No values Selected (green): Value 1, Value 2, Value 3, Value 4, Value 5  

Value 2 and Value 4 Selected (green): Value 2, Value 4, Value 1, Value 3, Value 5

This concludes the introduction of the Sense Excel User Interface. If there are any questions that have not been answered or should be described in more detail, please feel free to contact us with your feedback.


 ## 7. Using the Sense Excel Demo Report 

To help enable Sense Excel users to understand, create and edit a report we have created an example Excel workbook and App. These files, 5_Executive_Dashboard_Template.xlsx and 5_Executive_Dashboard_App are located within the Examples folder of a Sense Excel All in One download package or can be downloaded by clicking the link below.

https://www.senseexcel.com/downloads.html#Examples

Once you have the App copy it to the %user%\Documents\Qlik\Apps folder for use with Qlik Sense Desktop or upload it into your Qlik Sense server via Apps section of the QMC.

To use the Sense Excel 5_Executive_Dashboard_Template.xlsx report, please follow the following steps: 

1. Start Excel and make sure the Qlik Sense add-in is loaded and Toolbar is visible. 

2. Open the 5_Executive_Dashboard_Templte.xlsx file.

![Demo Report](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Demo-Report.png)

3. Connect to Qlik Sense Desktop or Server via the Connection portion of the Sense Excel Ribbon.

Go to the Sense Excel Ribbon >> Drop Down to Select Your Connection >> Press the Hub button (if necessary) >> Double Click on the “5_Executive_Dashboard_App".

![Open Hub Dialogue](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Open-Hub-Dialogue.png)

Now that the Excel report is connected to the Qlik Sense app (.qvf-file) any changes made to the data in app in Qlik Sense will be automatically reflected in the corresponding data in your Excel report. 

4. Use the Global Selector by pressing the "Selections tool" button and pressing the "Square with Arrow Inside Button" on the right side,

![Selection Tool Activate](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Selection-Tool-Activate.png)

5. Use the search box the find particular dimensions or fields you would like to filter on or use the slider on the bottom of the screen to look thru all available, make some filter selections and see how it automatically updates the figures in the Excel report in the background. 

6. View the filters that were just set in the Sense Filter Toolbar and click to edit or remove them. 

![Selection Tool Filter Selection](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Selection-Tool-Filter-Selection.png)

7. To see an example of Qlik syntax being used within and in conjunction with an Excel formula, click on any “ACTUAL” field within the Excel Demo report. There are four separate functions being utilized in the formula in this example.  These are explained in further detail in Chapter 8. 

![Formula Only](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Formula-Only.png)

![Formula Expanded](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Formula-Expanded.PNG)

8. This report also features the existing Bookmarks included in the Executive Dashboard app definition. When applied, these will globally update all cell values in your report the same way as it would in Qlik Sense. New Bookmarks can also be named and saved based upon any other filter selections made by the user and reused later. 

 ![Bookmarks Dialogue](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Bookmarks-Dialogue.png)

## 8. Create Your Own Reports

In this chapter we define two approaches to build a report in Excel using Dimensions, Measures, Formulas, Bookmarks and Filter settings directly from Qlik Sense. 

Approach 1 allows existing Table objects to be imported from Qlik Sense into your Excel workbook by utilizing the “Table” button in the Sense Excel Ribbon and dynamically choosing the relevant columns. 

Approach 2 creates a function for each discreet cell within the Excel workbook and displays a value or string. There are a number of available functions that can be utilized in this scenario.

### 8.1 Using the “Table” import button 

Click on the “Table” button to open the “Sense” property panel on the right side your screen.  As you will notice, the behavior within this feature is almost identical to that of Qlik Sense.

The following steps will import tables directly into your worksheet and allow you to filter them using the Global Selector button either before or after your import to create the report that you need. 

Option 1 allows you to import an complete existing Table object from a Qlik Sense app in one click.  In the 5_Executive_Dashboard_App example application, CY Sales vs. LY Sales is the available choice. 

Once an Table object is selected, the all columns within it will be displayed.

Option 2 allows you to choose your own columns using the Add Column button, identifying your column type by Dimension, Measure or Formula then scrolling or using the Search function to find and select your desired columns. 

Note: Master Items - Dimensions and Measures - defined in the Qlik Sense App cannot be used with the “Formula” function. Fields not identified as Master Elements in your data model, you can use the Formula selection and aggregating functions such as sum(), count() etc. on the fly to behave as if they were.

Option 3 "FAST ADD" allows you to select one or more Columns directly from your Qlik Sense Data model.  Available items are displayed to make finding your desired data elements easier by displaying Master Elements first then all other Fields below.  Dimensions, Measures and Fields. 

You can also use the FAST ADD button to add additional columns to any existing table.  Check the Dimension, Measure or Fields you want to include and press the green Check button.  This will add the additional Columns to the bottom of your table definition.

Once you have used on of the approaches above to get your desired Columns you can perform the additional actions prior to importing the data into your Excel workbook:

Data Section:

1. Reorder your Columns by clicking and dragging them up or down
2. Use the Triangle next to the Column name to expand for these additional options:
3. Check to Include or Suppress Null Values
4. Use a Qlik Sense Formula to create a dynamic Background or Text Colors e.g. =If(\[Amount]>100,lightred, )
5. Delete The Column from your table.

Sorting Section

1. Change to Sort Order of your Columns by clicking and dragging them up or down
2. Use the Triangle next to the Column name to expand for additional options.
3. The Default Sorting Setting is "Auto".  Slide the toggle switch left to enter the "Custom" Setting
4. Check the boxes to sort Numerically or Alphabetically and use the Drop-Down boxes to choose Ascending or Descending.

When imported tables are too large (which is often the case), further filtering is necessary to reduce the amount of data returned to a manageable amount.  In this situation, the use of the Global Selector is advised. The Global Selector gives you access to all of the Dimensions and Fields in the app and their associated values and offers a search function to return relevant options for you to filter on. Upon clicking on any value using the Global Selector, it will be added to the “Current selections” toolbar as a filter. The Global Selector can be closed via a click above the dark grey area. Once closed, the filters persist in the Sense Filter Toolbar and, once there, additional selections can be made to adjust your result set(s). 

### 8.2 Fixed Cell Report Definition

In order to take full advantage of the power of Sense Excel, it is very useful to understand  Qlik Sense formula syntax and especially the Set Analysis features. This report creation strategy allows formulas to be entered into a specific cell on an Excel sheet and return a unique value to that cell. This feature is especially valuable when creating Financial reports similar to the example.  Once a single formula is created you can use the copy/paste and fill across/down features of Excel to rapidly expand the scope of your report. 

### 8.3 Sense Excel Formulas and Syntax

Upon installation of the Sense Excel add-in, several additional formulas are made available for report development within the Excel formula library. These formulas, their usage and syntax are outlined below.  

#### 8.3.1 SenseConnected 

SenseConnected is the command with which the connection to Qlik Sense can be tested. Either it returns true (WAHR) or false (FALSCH) which indicates whether or not there is a live connection to Qlik Sense. The Syntax is as follows; =SenseConnected() 

#### 8.3.2 SenseFilter 

SenseFilter is a function that allows the use of two parameters. It can be used as a filter box and allows use of curtain wildcard characters such as " * ". The first parameter can be a fieldname (e.g. Year) in double quotation marks (“), then a semicolon (;), followed by the second parameter (e.g. 2013), which is the filter to be used (without quotation marks) The syntax is as follows: =SenseFilter("[Fiscal Year]"; 2013) 

This function is added to any empty cell within Excel and causes the imported tables and their associated field names to be filtered by a second parameter. The Cell with the Filter will state "Already Filtering". 

#### 8.3.3 SenseVariable 

The next Sense command is SenseVariable, which can use up to two parameters. The first parameter is the name of the variable available in Qlik Sense. The second parameter describes the content of the variable. The syntax is as follows: =SenseVariable(“A”, “Bee”) will place the content “Bee” in all places where variable “A” is implemented using the =SenseVariable() function. 

#### 8.3.4 SenseEV 

The example formula below is in the PRODUCT ANALYZE worksheet of the 5_Executive_Dasboard_Template.xlsx file available in the Sense Excel Examples library.

SenseEV stands for "Sense EValuate" and enables Excel to call any Qlik Sense function including set analysis within an Excel formula in a discreet cell.  The example below explains the syntax step by step and demonstrates the ability to combined the use Qlik Sense Set Analysis along with native Excel components. 

=SenseEV("Sum({<[Product Group Desc]={'"&$B7&"'}, [Fiscal Quarter] = {'"&SUBSTITUTE(C$5,"Q","")&"'}>} [Sales Quantity]\*[Sales Price])")

=SenseEV(" *This starts the connection to Qlik Sense and calculates… 

 Sum({< *…the Sum of \[Sales Quantity]\*\[Sales Price] and… 

[Product Group Desc] = {'"&$E8&"'}, // …is limited to where [Product Group Desc] = the value from cell E8… 

IMPORTANT: In this example, cell $E8 is a relative reference. This allows you to copy/paste and fill down/across in Excel and VERY rapidly populate an array within your report.

 [Fiscal Quarter] = {'"&SUBSTITUTE (L$4;"Q";"")&"'} // This uses an Excel function to replace the "Q" in the value "Q1" with a null and pass the "1" as a set parameter value for [Fiscal Quarter]. 
 
IMPORTANT:  This example demonstrates using "& INSERT EXCEL FORMULA HERE &" which allows you to enter "Excel mode" and access the full Excel formula library from the inside of a Qlik Sense expression.

 >} // 

 [Sales Quantity]\*[Sales Price]) //. 

 ") 

## 9. Seven Steps to Your First Report in Sense Excel 

### Step 1: Connect to Qlik Sense & select an App  

Create your connection, press the Sign In Button.  Once the "Hub" button is active (turns black), press it and select the app you would like to connect to.

### Step 2: Press the "Table" button 

From the Table Property panel on the right side of the screen:

Dropdown the "Choose Select Table Object to Import" box to choose an existing table defined in your Qlik Sense App

OR 

Dropdown "Add Column" to individually select your desired Dimensions, Measures or Formulas 

OR

Use the "Fast Add" button to select multiple objects from your data model

### Step 3: Press the Check button to create the table in your Excel workbook.

### Step 4. Edit Table Definition
You can add columns, change column display order, change your sort order or change your number formats returned prior to importing your table or afterward by right mouse clicking over the table would like to update and choosing "Edit Table". This will re-open the Table Property window and allow you to make any desired changes. You can also reopen the Table Property Panel by selecting a cell within your table and pressing the "Table" button again.

### Step 5: Open the Selections tool 
Press the Selection button and find the field and value your would like to filter on and select it.  Alternatively, use the Search function to limit your choices to those related to your search term.

### Step 6: Select a Bookmark 
Use the Bookmark button to use any of the existing bookmarks in your app or use "Create New" to create a new one based on the filters you have currently selected.

### Step 7: Create an Interactive Chart or Pivot table 
Once your table is created, you can use the Insert Chart or Insert Pivot table commands from the Insert tab of Excel and your first report is done. 

## 10. How to install on a terminal server

### Check: Access to the add-in
To set up the add-in, the user must have access to the add-in (xll file).
### Check: Registry settings  
Furthermore, a registry entry must be set in "HKEY_CURRENT_USER".
The following key contains the path to the add-in: "HKEY_CURRENT_USER\Software\Microsoft\Office\<OfficeVersion>\Excel\Options\"
There is or can be one or more "OPEN" keys OPENx= / R "<path_to_.xll>" see example:
"OPENx" in the above key may have one or more "OPEN" key values. One for each add-in. The following count applies:

e.g.:

OPEN = Path to add-in 1
OPEN1 = Path to add-in 2
OPEN2 = Path to add-in 3

You can put the add-in in a diffrent place that every user has access to, then every user has to set the registry entry in which this path is referenced.

<Path_to_add-in.xlll>
-	Is the .xll file in the path „C:\Users\%USERNAME%\AppData\Roaming\Microsoft\AddIns“ -  then an indication without a path is enough.

/R "akquinet-sense-excel-4.0.4-64.xll"
 
-	If the add-in is on a different path, it must be specified as well (for example, C:\AddIn\)

/R "C:\AddIn\akquinet-sense-excel-4.0.4-64.xll"

### Check: Officeversions
Office 2003 -> “15.0” (HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Excel\Options\)
Office 2016 and newer -> “16.0” (HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Excel\Options\)

### Important NOTE:
If no add-in is registered yet, the value (string) "OPEN" must be created.
If one or more add-ins already exist, a value (string) "OPEN1 ... etc. OPENx" must be created.
Where x is the next number to use (for example, "OPEN4" if Sense Excel is the 4th add-in)


### Update a version:
Our add-in carries the version number in its name. This can be unfavorable for you, if you want to provide the users with a new version, you always have to adjust the registry.
Therefore, the version number part of the file name and the registry entry can be removed, ie without a version number: for example: "akquinet-sense-excel.xll"
This has the consequence that the registry entry only needs to be set once, and to update the add-ins only the xll must be overwritten. Via "Copy and Paste".



## 11. Contact Information 

### Contact Support: 

Please register at http://support.qlik2go.net, login and open a “Bug” or “Feature request” ticket in order for us to provide you with support as soon as possible. We wish to provide you with the best Sense Excel experience possible and depend upon your feedback in order to improve Sense Excel.  

If you are happy with Sense Excel, please tell others. If you’re not,… please tell us how to improve. 

### Contact Sales

#### Europe/German Speaking Countries:

Juliane Tschierske

Telephone: +49 40 881 73 26 21

E-Mail:  [juliane.tschierske@akquinet.de]

#### Americas:

Lance Harris

Master Agent, Sales & Alliances

Telephone: +1 703 625 7738

E-Mail: [lance.harris@senseexcel.com]

#### Rest of the World: 

Michael Walther

Telephone:  +49 40 88173 2617

E-Mail: [michael.walther@akquinet.de]

#### Address: 

akquinet finance & controlling 

Paul-Stritter-Weg 5 

D-22297 Hamburg , Germany


## Appendix 1 - Older Version Installation Instructions

### Appendix 1.1 Extract Download Package

Extract (Unzip) the files of the Sense Excel download package to a location of your choosing. 

BEST PRACTICE: Create a new directory in your Documents folder called Sense Excel with a sub-folder referencing the version number like %User%\Documents\Sense Excel\x.x.x and extract your software here.  

DO NOT load Sense Excel from the inside of a compressed (Zip) file archive.


###  Appendix 1.2 New Sense Excel Installation 

1  Open the location of your Sense Excel software files. Double-click on the .xll file corresponding to your environment (32 or 64 bit).

2. Excel should start automatically, include "SENSE" in the Excel Menu bar and display the Sense Excel Ribbon when "SENSE" is selected. 

3. Go to the "Settings" section of the Sense Excel Ribbon, press the triangle to expand the menu and check "Auto Load". This will automatically start Sense Excel every subsequent time you open Excel.

4. Close Excel and re-open it.  "SENSE" should appear in the Excel Menu Bar.

### Appendix 1.3 Alternate Installation Technique

In limited instances, the Sense Excel Add-in might not register properly using the double-click installation method. If this is the case, please use the alternate installation technique described below.

1. Open Excel and a blank workbook.

2. Navigate to File > Options > Add-Ins.  Press the "Go" button then the "Browse" button. 

3. Navigate to the location of your Sense Excel software and double-click on the appropriate .xll file. 

4. Once the Add-in loads and is registered, go to the "Settings" section of the Sense Excel Ribbon, press the triangle below the "Settings" button and check Auto-Load.  

5. Exit and Restart Excel.

### Appendix 1.4 Upgrade Sense Excel Installation

To upgrade an existing version of Sense Excel please perform the following steps:

1. Open Excel and a blank workbook.

2. Go to File > Options > Add-Ins > Press the "Go" button.

3. Uncheck all "akquinet-sense-excel" Add-ins that are displayed.  Press OK.  Confirm that the "SENSE" entry is removed from the Excel Menu bar.

4. Go to File > Options > Add-Ins > Press the "Browse" button.  This will take you to the %user%\Roaming\Microsoft\Addins directory.

5. Dropdown the File Type Drop Down Box to select "All Files (*.*)". Delete all "akquinet-sense-excel" (.xll and .xlldel) files in the directory.  Press OK.

6. Go to File > Options > Add-Ins > "Go" button > Browse.  Navigate to the location of your updated Sense Excel software files.

7. Double-click the appropriate .xll file and wait for the add-in to load and register.  The "SENSE" entry should re-appear in the Excel Menu bar.

8. Go to Settings on the Sense Excel Ribbon, press the triangle to expand the menu and check "Auto Load".

9. Close and Restart Excel.

10. Confirm that the "SENSE" entry appears in the Excel Menu Bar.


## Appendix 2 - Sense EV Example Formulas

1. Simple Sum with One Measure:

=SenseEV("Sum([Amount])")


2. Sum with Single Set Analysis Parameter.  Excel Formula set to allow filling down.

=SenseEV("Sum({<[Quarter]={'"&$B4&"'}>} [Amount])")


3. Sum with Single Set Analysis Parameter.  Excel Formula set to allow filling down. 2nd Parameter, GUID to Pull from different App.

=SenseEV("Sum({<[Quarter]={'"&$B4&"'}>} [Amount])","eba546df-658a-4b6d-bee5-42d3f9fa55e2")


4. Sum of Revenue measure created from 2 fields. 2 Set Analysis parameters. First parameter set to populate from Excel cell values and fill down.  Second paramter uses Excel "Substitute" formula to extract value from header cell, C5, and able to fill across.

=SenseEV("Sum({<[Product Group Desc]={'"&$B7&"'}, [Fiscal Quarter] = {'"&SUBSTITUTE(C$5,"Q","")&"'}>} [Sales Quantity]*[Sales Price])")


5. Sum of Revenue measure calulated on the fly from 2 fields. First parameter set to populate from Excel cell values and fill down. Second set parameter set to only calculate Revenue for the most recent year in the data.
 
=SenseEV("Sum({<[Product Group Desc]={'"&$B7&"'}, [Fiscal Year] = {'<$(=max([Fiscal Year]))'}>} [Sales Quantity]*[Sales Price])")


6. Adds a function to triple the value of the entire calculation.

=SenseEV("Sum({<[Fiscal Year] = {$(=max([Fiscal Year]))}>} [Sales Quantity]\*[Sales Price])")*3


7. Only calculates for "Last Year"

=SenseEV("Sum({<[Fiscal Year] = {$(=max([Fiscal Year])-1)}>} [Sales Quantity]\*[Sales Price])")


8. Dynamic Title to show the text "PRODUCT SALES VOLUME" along with each selected value of Region delimited by a "/" for the most recent year.

=SenseEV("'PRODUCT SALES VOLUME '&Upper(Concat(distinct [Region],' / ')) &' '& max([Fiscal Year])")


9. Dynamic Title to show each selected value of Region delimited by a "/".

=SenseEV("Upper(Concat(distinct [Region],' / '))")


10. Formula to quickly populate an array.  Measure created on the fly with two set parameters both referencing Excel cell values.  First set value allows filling down, second set value allows for filling across. 

=SenseEV("Count({<[Case Owner Group]={'"&$A4&"'}, [Priority] = {'"&B$3&"'}>} [%CaseId])")


11. 3 Set Values set to dynamically filter data based on column header values from Excel cells.  1st and 2nd levels are fixed. 3rd level allows filling down. 4th and 5th levels filter for only Income Statement and Actual values. 

=SenseEV("Sum({<[1st Level]={'"&$B$12&"'},[2nd Level]={'"&$B$14&"'},[3rd Level]={'"&$C15&"'},Statement={'Income Statement'},DataCategory={'Actual'}>} [Amount])")


12. Same as above additionally filtered for only most recent year's data.

=SenseEV("Sum({<[1st Level]={'"&$B$12&"'},[2nd Level]={'"&$C13&"'},[Year] = {$(=Max([Year])-1)},Statement={'Income Statement'},DataCategory={'Actual'}>} [Amount])")


13. Filters for only the data from 1 year prior to the most recent values of date in the data.

=SenseEV("Sum({<[1st Level]={'"&$B$12&"'},[2nd Level]={'"&$C13&"'},[AsOfMonth]=,[DateCounter] = {$(=Max([DateCounter])-12)},Statement={'Balance Sheet'},DataCategory={'Actual'}>} [Amount])")


14. 2 levels of Set parameters to filter based on Excel cell values.  Additionally populates set parameters to filter for dates derived from an Excel cell value.  Allows for filling across.

=SenseEV("Sum({<[1st Level]={'"&$B$12&"'},[2nd Level]={'"&$C13&"'},Year = {$(=Year('"&E$10&"'))},[Month] = {$(=Month('"&E$10&"'))},Statement={'Income Statement'},DataCategory={'Actual'}>} [Amount])")


15. 3 levels of Set parameters to filter based on Excel cell values.  Additionally populates set parameters to filter for dates derived from an Excel cell value.  Allows for filling across.

=SenseEV("Sum({<[1st Level]={'"&$B$12&"'},[2nd Level]={'"&$B$14&"'},[3rd Level]={'"&$C15&"'},Year = {$(=Year('"&E$10&"'))},[Month] = {$(=Month('"&E$10&"'))},Statement={'Income Statement'},DataCategory={'Actual'}>} [Amount])")


16.  Calculation for a snapshot in time, Balance Sheet, using AsOfMonth as the filter.

=SenseEV("Sum({<[1st Level]={'"&$B$12&"'},[2nd Level]={'"&$C14&"'},[AsOfMonth] ={'"&E$9&"'} ,Statement={'Balance Sheet'},DataCategory={'Actual'}>} [Amount])")


17. 3 set parameters.  1st level is fixed.  Second level allows filling down.  3rd level is fixed.

=SenseEV("Sum({<[1st Level]={'"&$B$11&"'},[2nd Level]={'"&$C12&"'},[CompanyID]={'"&E$9&"'},Statement={'Income Statement'},DataCategory={'Actual'}>} [Amount])")


18. 3 set parameters.  1st and 2nd levels are fixed.  3rd level allow filling down.

=SenseEV("Sum({<[1st Level]={'"&$B$11&"'},[2nd Level]={'"&$B$13&"'},[3rd Level]={'"&$C14&"'},[CompanyID]={'"&E$9&"'},Statement={'Income Statement'},DataCategory={'Actual'}>} [Amount])")
