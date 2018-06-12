# Sense Excel Detailed Installation Guide



 ![Cover](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Cover.PNG)


## Table of Contents 

## 1. Please Note
### 1.1 Installation Notes
### 1.2 Disclaimer 
## 2. Introduction 
### 2.1 Who We Are
### 2.2 Why We Do What We Do
## 3. What You Can Expect from Sense Excel
## 4. Requirements 
## 5. Installation Guide 
### 5.1 Download Software  
### 5.2 Unpack Files
### 5.3 Basic (Desktop) Installation 
### 5.4 Qlik Sense Server Configuration
## 6.  Explaining the User Interface
### 6.1. "Connections"
### 6.2. "Open Hub"
### 6.3. “Load data” and “Data load editor”
### 6.4. “Table” 
### 6.5. "Bookmarks" 
### 6.6. “Selections tool” 
### 6.7. "About" 
### 6.8. Sense Filter Toolbar
## 7. Using the Sense Excel Demo Report
## 8. Create Your Own Reports
### 8.1 Using the “Table” Import Button 
### 8.2 Fixed Cell Report Definition 
### 8.3 Sense Excel Formulas and Syntax
#### 8.3.1 SenseConnected 
#### 8.3.2 SenseFilter
#### 8.3.3 SenseVariable 
#### 8.3.4 SenseEV 
### 8.4 Using “Load Data” and "Data Load Editor"  
## 9. Connection Options
## 10. Seven Steps to Your First Report in Sense Excel
## 11. Contact Information




## 1. Please Note 

### 1.1 Installation Notes 

If previous installations of Sense Excel exist on your system please read the following installation advice. 

1. Close MS Excel and make sure it is not running in the background. 
2. Extract the new downloaded version of Sense Excel into the existing installation folder and overwrite (all) the old installation data. 
3. Start MS Excel, go to the Qlik Sense tab and check the “Auto load” option (if not checked already) under the about menu. This will start Sense Excel when Microsoft® Excel is started. 
4. Check to see that the version number equals the downloaded and installed version of Sense Excel. This can be confirmed by clicking on the question mark in the about section of the Qlik Sense tab. 

### 1.2 Disclaimer 

This document will cover how to get Sense Excel up and running on your system, what to expect and ways the product can be used. This document is written to provide quick answers and will be continuously updated. Not all options or possibilities of Sense Excel are covered within the scope of this document but will be complimented by best practices and use cases documented elsewhere in this repository. 

 AKQUINET SENSE EXCEL LICENSE AGREEMENT 

IMPORTANT: BY ACCEPTING, DOWNLOADING OR USING THIS SOFTWARE, YOU ACCEPT AND AGREE TO THE TERMS OF THIS AKQUINET SENSE EXCEL LICENSE AGREEMENT (“DLA”) AS MAY BE UPDATED FROM TIME TO TIME AND PUBLISHED AT WWW.AKQUINET.COM. BY ACCEPTING THESE TERMS, OR USING THE SOFTWARE AS AN EMPLOYEE, CONTRACTOR OR AGENT ON BEHALF OF A CORPORATE OR OTHER ENTITY, YOU (“USER”) REPRESENT AND WARRANT THAT YOU HAVE AUTHORITY TO BIND SUCH ENTITY TO THESE TERMS. DIRECT COMPETITORS AND THEIR EMPLOYEES AND AGENTS MAY NOT ACCESS THE SOFTWARE 
WITHOUT PRIOR WRITTEN CONSENT OF AKQUINET. THE SOFTWARE MAY NOT BE USED FOR PURPOSES OF BENCHMARKING, COLLECTING AND PUBLISHING SOFTWARE PERFORMANCE DATA OR ANALYSIS, OR ANY OTHER COMPETITIVE PURPOSES. 
 
NOTICE: THE SOFTWARE CONTAINS FUNCTIONALITY INTENDED TO LIMIT THE DURATION OF ITS USE AND IS INTENDED TO COLLECT CERTAIN USAGE METRICS. THE INSTALLATION OF THIS SOFTWARE WILL INSTALL FILES NECESSARY TO OPERATE THE SOFTWARE ONTO THE USER’S COMPUTER AND OTHER SYSTEM FILES MAY BE INSTALLED OR UPDATED. AS WITH ALL INSTALLATIONS, BACK UP OF THE USER’S HARD DRIVE IS RECOMMENDED BEFORE INSTALLING THE SOFTWARE. 

The rest of the disclaimer can be read at Qlik2Go (https://support.qlik2go.net/projects/sense-excel/wiki) under "Terms of use". 

## 2. Introduction 

### 2.1 Who We Are

akquinet AG is a stock corporation headquartered in Hamburg, Germany with approximately 900 employees. We offer a wide range of IT and BI implementation services including Enterprise Resource Planning (ERP) systems utilizing Microsoft and SAP, custom ERP development using JAVA as well as managed hosting services. akquinet's Business Intelligence team is located in Jena, Germany is and responsible for development of the Sense Excel product line. 


### 2.2 Why We Do What We Do

We believe that all stakeholders should have easy access to business critical information when and where they need it. Enabling all users to access their data in a familiar interface while, at the same time, fully reaping the benefits of the "single version of truth" aspect of an Enterprise BI Platform is our primary goal.


## 3. What You Can Expect from Sense Excel 

In this chapter, we will explain what can be expected of Sense Excel and what cannot. 
 
### Sense Excel can/does not... 

• …import a complete Qlik Sense Dashboard into Excel

• …import visualization objects from Qlik Sense into Excel

• …change Excel in any way other than extending its reporting capabilities 

### Sense Excel can/does.....

• …provide the same UI (user interface) as Qlik Sense including filter selections and bookmarks for ease of use.

• …provide all Dimensions, Measures, Fields, Bookmarks etc. that are available in the connected Qlik Sense app.

• …allow for creation of data tables in Excel utilizing the familiar workflow of Qlik Sense.

• …allow for a one click download of underlying object data from Qlik Sense into an Excel table.

• …allow editing and reloading of the Qlik Sense application script from directly inside of Excel.

• …allow access to the complete Qlik Sense function libarary, including Set Analysis, directly within an Excel formula.

• …take full advantage of the best features of Excel including copy/paste and fill down/across to rapidily increase your speed of development.

• …allow for population of Qlik Sense formula and Set Analysis parameters utilizing embedded Excel cell references.

• …allow you take full advantage of all Excel features including charting and pivot tables on data that comes directly from Qlik Sense.

• …allow for access to multiple apps (.qvf files) within a single Excel workbook.  

To summarize, with Sense Excel you get all of the power and flexibility of Microsoft Excel combined with all of the speed, governance and navigation capabilities of Qlik Sense. We are confident that Sense Excel will enhance both your Qlik Sense and Excel experience and provide you with consistent, reliable, real-time data for your reporting efforts.

## 4.  Requirements

### Please make sure the following system requirements are met: 

• You have acccess to either Qlik Sense Desktop or Qlik Sense Server software.

• Confirm that you have a .NET Framework >= 4.5.1, otherwise please install it. 

• You have Microsoft Excel 2013, 2016 or Office 365-Desktop (32 or 64 bit) installed.

To find out which Microsoft Office edition you are running within Excel, go to FILE >> ACCOUNT >> "About Excel".  Behind the MSO number, it will display 32-Bit or 64-Bit.

![Excel Version](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Excel-Version.PNG)


## 5. Installation Guide

### 5.1 Download Software
You can download the Sense Excel software from the following link https://m.sense2go.net/extension-package or from https://www.senseexcel.com. 

### 5.2 Unpack Files

Unpack (unzip) the files from the Sense Excel directory of your download package to a location of your choosing. 

###  5.3 Basic (Desktop) Installation 

1. Choose the .xll-install file corresponding to your environment (32 or 64 bit) in the installation folder and double click it. 

2. Excel should start automatically and add the "Qlik Sense" tab and associated toolbar to Excel. 

3. Click "About" and check the "Auto load" option if it isn't checked already. This will start Sense Excel every time you open MS Excel.  
4.  Please take the time to read our terms of use (click on About) and confirm that the version you are using corresponds to the version listed in the Info window. 

### 5.4  Qlik Sense Server Configuration

To connect Sense Excel to a Qlik Sense Server, the additional configuration steps listed below need to be performed as well. 

Only paid versions of Sense Excel can be connected to a Qlik Sense Enterprise Site (server) installation. EXCEPTION: All installations utilizing Qlik internal license keys or current Qlik trial or partner keys.  

30-day server trial licenses are available for existing Qlik Sense customers.  Please contact an authorized reseller or the appropriate sales contact listed in Chapter 11 of this document. 
 
1. Open the Qlik Management Console (QMC) 

On the server/computer go to: Start >> Programs >> Qlik Sense >> Qlik Management Console or use any browser and type: https://”your_sense_server_ name”/qmc 

2. Add a Security rule to your Qlik Sense Server. 

Within the QMC >> Security rules >> Create New 

Name: LicenseReadAll

Resource filter: License_* 

Condition: !user.IsAnonymous() 

Context: Only in QMC 

Actions: Read 

Press Apply and confirm the rule has been added. 

3.  Create a Content Library on the Qlik Sense Server. 

a. Within the QMC >> Content libraries >> Create New >> Enter the name “senseexcel” >> Apply.

b. Copy the contents of your organzation's LEF file or trial key into a new text document. 

c. For token licensing strategies, the contents of your LEF or trial key alone will suffice.  

For named user licensing strategies please follow the additional steps below.

d. Append the LEF information with the domain and user name of the users which will access Sense Excel in the format below.  The FROM; TO information is optional if you choose to put time limits on the license duration of your named users.

 9999999999999999 

 EXCEL_NAMED;5;;YYYY-MM-DD 

 AAAA-BBBB-CCCC-DDDD-EEEE 

 EXCEL_NAME;DOMAIN\USER1 

 EXCEL_NAME;DOMAIN\USER2;FROM;TO 

 EXCEL_NAME;DOMAIN\USER3;FROM;TO 

 EXCEL_NAME;DOMAIN\USER4;FROM;TO 

 EXCEL_NAME;DOMAIN\USER5;FROM;TO 


e. Please confirm that there are no spaces at the end of each line.

f. Save this file with the name license.txt.

g. Upload the license.txt file to the senseexcel content library. 

QMC >> Content Libararies >> Upload >> Select File “license.txt”. 

This file will be the tool for managing the licenses of all of your Sense Excel users going forward.  It can be updated and overwritten as necessary.

4.  “Qlik Sense Desktop” is the default connection string.  Once the security rule and content library have been added, you can connect to your Qlik Sense Server from the Connection section of the Qlik Sense toolbar.

To connect to a Qlik Sense server, enter the (URL) without “/hub”. 

You should first try to connect using a secured (HTTPS) connection https://your.server.url/. If your Qlik Sense server does not have a valid security certificate (these certificates can be purchased online), you might receive an error message.   

If you do receive a certificate error, try using an unsecured connection (HTTP) http://your.server.url/. This configuration is supported if this option is enabled within the QMC.

QMC> Proxies> Central> Edit> Ports> Check “Allow HTTP”.

More on this topic is covered in Chapter 9 of this document.

## 6.  Explaining the User Interface 

To get the most out of your Sense Excel experience it is necessary to understand the UI (User Interface) used in Sense Excel. The  picture below shows the “Qlik Sense” toolbar as well as the features that make up Sense Excel.

 ![Toolbar](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar.PNG)

### 6.1. "Connections"
This enables you to connect to your Qlik Sense Desktop or Server (URL).
 
 ![Connection](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Connecton.png)
 
a. Sign In : Signing in will either start Qlik Sense Desktop in the background or connect to the Qlik Sense Hub on your server. 
b. Sign Out: This will end your active connection to the server. 
 
### 6.2. "Open Hub"

 ![Open Hub](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Open-Hub.png)

This button allows you to choose your application and access the dimensions, measures and underlying object data associated with it.  

 ![Open Hub Dialogue](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Open-Hub-Dialogue.png)

### 6.3. “Load data” and “Data load editor”
 ![Load Data](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Load-Data.png)

These buttons allow you, with proper credentials, to edit and execute the load script associated with your application. 

 ![Data load editor](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Data-load-editor.png)

### 6.4. “Table” 

 ![Table](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Table.png)

This button will activate the Table Property Panel which allows you to import underlying object data, such as from a chart, directly from your app into an Excel table or create a new table by selecting columns Data elements - Dimensions, Measures, Fields or Formulas and apply sorting and add-on properties.

 ![Table Editor Onopen](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Table-Editor-Onopen.PNG)
 
 
### 6.5. "Bookmarks" 

 ![Bookmarks](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Bookmarks.png)

This option allows you to access existing Bookmarks in your Qlik Sense app or save your current filter selections as a new Bookmark that can be reused later.Bookmarks created in Sense Excel are not re-imported into the connected Qlik Sense app.

 ![Bookmarks Dialogue](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Bookmarks-Dialogue.png)


### 6.6. “Selections tool” 

 ![Selection Tool Button](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Selection-Tool-Button.png)

This button will show or hide the Sense Filter Bar where all filter selections will be displayed. 

![Selection Tool Sense Filter Bar](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Selection-Tool-Sense-Filter-Bar.PNG)

Once this toolbar is activated, you can use the square box on the right side to pull up all of your available dimensions as list boxes from which to select values for as filters.

 ![Selection Tool Activate](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Selection-Tool-Activate.PNG)


Once inside of this interface, you can also use the Search function to reduce the dislayed elements to those relevant to your query.

 ![Selection Tool Search](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Selection-Tool-Search.png)
 
 ![Selection Tool Search Value](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Selection-Tool-Search-Value.png)
 
  ![Selection Tool List Box Select](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Selection-Tool-List-Box-Select.png)

### 6.7. "About" 

 ![Toolbar About](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-About.png)

Clicking on the question mark brings up an additional window with the Sense Excel version and disclaimer information. 

The “Support” option will open your default e-mail client and create a message with an attached log file addressed to our support-team. 

 ![Toolbar About Support](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-About-Support.PNG)

Checking the “Auto load”-option will load the Qlik Sense toolbar automatically when starting Excel. 

 ![Toolbar About Autoload](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-About-Autoload.png)

"Language" allows you to choose which language you would like to work in.

 ![Toolbar About Language](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-About-Language.png)


### 6.8. Sense Filter Toolbar Navigation

![Selection Tool Sense Filter Bar](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Selection-Tool-Sense-Filter-Bar.PNG)

The buttons here allow you to manipulate your selections e.g. “Step back”, “Step forward” and “Clear all selections. Furthermore, when a filter is set, you can change the values of your selection by clicking on it and using the Check and X buttons to select and deselect filter values respectively,

This concludes the introduction of the Sense Excel User Interface. If there are any questions that have not been answered or should be described in more detail, please feel free to contact us with your feedback.

 ## 7. Using the Sense Excel Demo Report 

To enable the user of Sense Excel to understand, create and edit a report we have included an example. This file, ExecutiveDashboard.xlsx, is located in the Examples folder within the Sense Excel Reporting directory of your download package. It is built using the Executive Dashboard example app that ships with Qlik Sense Desktop. If you do not have access to the app in your environment, it can also be found in the Examples directory of the Sense Excel section of your download package.  From there it can be uploaded into your Qlik Sense server via Apps section of the QMC.

To use the Sense Excel Executive Dashboard Report, please follow the following steps: 

1. Start Excel and make sure the Qlik Sense add-in is loaded and Toolbar is visible. 

2. Open the ExecutiveDashboard.xlsx file.

 ![Demo Report](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Demo-Report.PNG)

3. Connect to the Qlik Sense Desktop or Server via the "Log-in" button.

 Once opened, go to the Qlik Sense toolbar >> Log in >> click on “Open Hub” >> click on the “Executive Dashboard” app.

 ![Open Hub Dialogue](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Open-Hub-Dialogue.png)

Now that the Excel report is connected to the Qlik Sense app (.qvf-file) any changes made to the data in app in Qlik Sense will be automatically reflected in the corresponding data in your Excel report. 

4. Use the Global Selector, choose some filters and see how it automatically updates the figures in the Excel report in the background. Close the Global Selector by clicking above the black background. 

 ![Selection Tool Activate](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Selection-Tool-Activate.PNG)

5. View the filters that were just set in the Sense Filter Toolbar and click to edit or remove them. 

6. To see an example of Qlik syntax being used within and in conjunction with an Excel formula, click on any “ACTUAL” field within the Excel Demo report. There are four seperate functions being utilized in the formula in this example.  These are explained in further detail in Chapter 8. 

![Formula Only](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Formula-Only.png)

![Formula Expanded](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Formula-Expanded.PNG)

7. This report also features the existing Bookmarks included in the Executive Dashboard app definition. When applied, these will globally update all cell values in your report the same way as it would in Qlik Sense. New Bookmarks can also be named and saved based upon any other filter selections made by the user and reused later. 

 ![Bookmarks Dialogue](https://github.com/senseexcel/senseexcel/blob/master/images/SE-Toolbar-Bookmarks-Dialogue.png)

## 8. Create Your Own Reports

 In this chapter we define two approaches to build a report in Excel using Dimensions, Measures, Formulas, Bookmarks and Filter settings directly from Qlik Sense. 

Approach 1 allows complete tables to be imported from Qlik Sense into your Excel workbork using the “Table” button and dynamically choosing the relevant columns. 

Approach 2 creates a function for each discreet cell within the Excel workbook and displays a value or string. There are a number of available functions that can be utilized in this scenario.

### 8.1 Using the “Table” import button 

Another way to get the data you need into Excel is through use of the “Table” button. The following steps will import tables directly into your worksheet and allow you to filter them using the Global Selector button either before or after your import to create the report that you need. 

Click on the “Table” button to open the “Sense” property panel on the right side your screen.  As you will notice, the behavior within this feature is almost identical to that of Qlik Sense.

Option 1 allows you to import the underlying data from any existing object within your Qlik Sense app.  In the example application. CY Sales vs. LY Sales is the available choice. Once an object is selected, the columns within it are displayed and you can perfom the following actions prior to importing the data into your Excel workbook:

#### 1. Reorder fields by clicking and dragging them up or down
#### 2. Remove any field by right mouse clicking and choosing Clear.
#### 3. Change number formatting by expanding the column information using the small triangle.
#### 4. Change sort order using the Sorting option and dragging the columns up and/or down into the desired sort order.  
#### 5. Once you have specified the desired behavior in the steps above, pressing the Check button will import the into a table in your adjacent Excel sheet. 

Option 2 allows you to choose your own columns using the Add Column button, identfying your column type by Dimension, Measure or Formula then scrolling or using the Search function to find and select your desired columns. 

Note: All Dimensions and Measures defined as master items in the Qlik Sense App cannot be used with the “Formula” function. However, for Fields not idenfied as Master Elements in your data model, you can use the Formula selection and aggregating functions such as sum(), count() etc. on the fly to behave as if they were.

 When imported tables are too large (which is often the case), further filtering is necessary to reduce the amount of data returned to a manageable amount.  In this situation, the use of the Global Selector is advised. The Global Selector gives you access to all of the Dimensions and Fields in the app and their associated values and offers a search function to return relevant options for you to filter on. Upon clicking on any value using the Global Selector, it will be added to the “Current selections” toolbar as a filter. The Global Selector can be closed via a click above the dark grey area. Once closed, the filters persist in the Sense Filter Toolbar and, once there, additional selections can be made to adjust your result set(s). 

### 8.2 Fixed Cell Report Definition

In order to take full advantage of the power of Sense Excel, it is very useful to understand  Qlik Sense formula syntax and  especially the Set Analysis features. This report creation strategy allows formulas to be entered into a specific cell on an Excel sheet and return a unique value to that cell. This feature is especially valuable when creating Financial reports similar to the example.  Once a single formula is created you can use the copy/paste and fill across/down features of Excel to rapidly expand the scope of your report. 

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

SenseEV stands for "Sense EValuate" and enables Excel to call any Qlik Sense function including set analysis within an Excel formula in a discreet cell.   To understand the syntax the following example of a set analysis combined with excel elements has been added. 

=SenseEV(" // This starts the connection to Qlik Sense and calculates… 

 Sum({< // …the Sum of [Sales Quantity]*[Sales Price] and… 

[Product Group Desc] = {'"&$E8&"'}, // …is limited to where [Product Group Desc] = the value from cell E8… 

 [Fiscal Quarter] = {'"&SUBSTITUTE (L$4;"Q";"")&"'} // …the "1" is parsed from the value "Q1" in the cell and is passed as a set parameter value for [Fiscal Quarter]. 

 >} // 

 [Sales Quantity]*[Sales Price]) //. 

 ") 


### 8.4 Using “Load Data” and "Data Load Editor"  

Sense Excel provides the ability to edit, save and reload Qlik Sense application scripts directly from Excel. This feature can be limited to administrators and cetain report developers (power users) to maintain adequate security and consistency. 

Click the "Data Load Editor" Button to edit your script, make your desired changes then click the “Load data” button once to accept and save them. Click the "Load Data" button a second time to run the script in the background. 

Notes: You will not see a reload dialog, but the updated results will be automatically reflected in your report after reloading. 

In order to use the "Data Load Editor" and the “Load Data” options, you need to be able to reload the application you wish to edit. The data included with the Qlik Sense example applications (e.g.Executive Dashboard) can't be reloaded and therefore can't be used for testing and understanding these features.

## 9. Connection Options 

As mentioned above you can use the “Connection String” option to connect Sense Excel to either the Qlik Sense Desktop or Qlik Sense server via a http:// or https:// server address (without “/hub”!). 

Connections can be further managed by clicking on the small rectangle at the bottom right corner of the Connection section of the Qlik Sense toolbar. Here you can define the following parameters:

Connection Name - This can be changed to a user friendly name such as "Test Server".

Connection Type - Depending on your Qlik Sense authorization strategy and which browser you are using, this parameter can be set to allow Sense Excel to work in your environment.  The Qlik Sense default is to use NTLM, also called Windows Authentication.  In this situation, if the credentials that you use to sign into your Qlik Sense Server do not match your Windows credentials, this value would be set to "Cookie" which would prompt for different credentials to access Qlik Sense.  There are other options avaialble to ensure that Sense Excel will support whichever approach is being used in your situation. 

Connection uri:  The url used for connecting to your Qlik Sense Server.

Ignore Certificate Errors Checkbox. Sometimes https://sites have untrusted certificates and this setting defines how Sense Excel will behave in this scenario.  Checking the box will not stop connecting if https://certificate errors are encountered and ignore these errors.  Uncheking the box (which is usually the default) will allow the https:// chain to remain intact and provides the most security. 


## 10. Seven Steps to Your First Report in Sense Excel 

### Step 1: Connect to Qlik Sense & select an App  
Enter your connection string, press the Sign In Button.  Once the "Open Hub" button is active, press it and select the app you would like to connect to.

### Step 2: Press the "Table" button 
From the Table Property panel on the right side of the screen, choose your desired Dimensions, Measures or Formulas.

### Step 3: Press the Check button to create the table in your Excel workbook.

### Step 4. Edit Table Definition
You can add columns, change column display order, change your sort order or change your number formats returned prior to importing your table or afterward by right mouse clicking over the table would like to update and choosing "Edit Table". This will re-open the Table Property window and allow you to make any desired changes. 

### Step 5: Open the Selections tool 
Press the Selection button and find the field and value your would like to filter on and select it.  Alternatively, use the Search function to limit your choices to those related to your search term.

### Step 6: Select a Bookmark 
Use the Bookmark button to use any of the existing bookmarks in your app or use "Create New" to create a new one based on the filters you have currently selected.

### Step 7: Create an Interactive Chart or Pivot table 
Once your table is created, you can use the Insert Chart or Insert Pivot table commands from the Insert tab of Excel and your first report is done. 


## 11. Contact Information 

### Contact Support: 

Please register at http://support.qlik2go.net, login and open a “Bug” or “Feature request” ticket in order for us to provide you with support as soon as possible. We wish to provide you with the best Sense Excel experience possible and depend upon your feedback in order to improve Sense Excel.  

If you are happy with Sense Excel, please tell others. If you’re not,… please tell us how to improve. 

### Contact Sales

#### Europe/German Speaking Countries:

Juliane Tschierske

Telephone: +49 40 881 73 26 21

E-Mail:  juliane.tschierske@akquinet.de

#### Americas:

Lance Harris

Master Agent, Sales & Alliances

Telephone: +1 703 625 7738

E-Mail:  lharris@strategypoint.com

#### Rest of the World: 

Michael Walther

Telephone:  +49 40 88173 2617

E-Mail: [michael.walther@akquinet.de]

#### Address: 

akquinet finance & controlling 

Paul-Stritter-Weg 5 

D-22297 Hamburg , Germany
