Regression-Tests for Sense Excel
* 1 Ribbon
  * 1.1 Tooltip / Supertip
    
    When moving the mouse over the ribbon-elements, a tooltip should appear, describing what is done by this element.
    Test for all languages. Important are English, Korean, German
    
* 2  Connections Section
  * 2.1 Connection string combobox
  
    The user can insert a serveradress in this box and hit enter. 
    This creates a temporary Connection (NTLM) and starts to connect with it. 
    If the Connection could be established, it is saved in the sense-excel.jsconfig.  
 
  * 2.2 Sign In
  
    Connect to the selected address. Statusmessage should be shown in the Excel status bar.
    If the connection could be established, the Hub Icon gets enabled.
  
  * 2.3 Sign Out
  
    Disconnect from opended connection.
    The Hub Icon gets disabled.
  
* 3 Hub
    
  Clicking the Hub Icon opens the Hub.
      
  * 3.1 Username
     
    In the topleft corner the username should be Visible.
  * 3.2 Streams
        
    For Local connections (connectiontype local) only the personal streams should be visible
        
    For  all other Types, public Streams should be shown.
        
  * 3.3 Searchbar
      
    Entered text should filter apps by appname, filename(app-id), last reload date.
    Should be focused after hub opens. Esc deletes the entered text, Esc again closes the hub.
        
  * 3.4 Sorting
      
    * 3.4.1 sort-direction
      
      click on this button should change the sort direction of the apps
      
    * 3.4.2 sort-by
    
      change the field which is sorted by.
      Atm. only Alphabetically is suported.
    
    * 3.5 Toggle between the icon view and list view of the applist. 
    
    * 3.6 clicking the i-symbol next to an appname should open a panel with further appdescriptions.
    * 3.7 Select an app and open hub again, the selected app should be highlighted.
  * 4 Icon of selected app
  
    if an app is selected, the appname, and its icon should appear nect to the hubicon. 
    The name of the app should be limited to 25 characters.
    
  * 5 Load Data
  * 6 Data load Editor
  * 7 Table
      * 7.1 click on Table opens the table property-editor in a sidepane, table icon is pressed (gray background). After closing the sidepanel, the table button should be unpredssed (normal background)       
      * 7.2 Import Combobox
        
        Should contain all Tables of the app and show their names. 
        If no name is assigned to the table it should show 'No Title(IdOfTable)'
      
      * 7.2.1 selecting a table from the list
      
        After selecting a table its Dimension/Measures/Fields should be inserted in the Data tab 
        in the same order as declared in qlik sense.
        After pressing the green ok-tick, the table should be inserted in Excel.
      
      * 7.3 Pivote mode
              
        If pivotmode changes the defaulttext in the import combobox should change accordingly.
        The View in Data Tab should change too. 
        
      * 7.4 additional tests
      
        * 7.4.1 Pivot mode off
          * 7.4.1.1 select an empty Cell in Excel, in the table panel add a dimension, click green tick button -> a Table with a single column should be inserted into Excel
          * 7.4.1.2 select an empty cell in Excel, the tableeditor should be cleared and the importcomboboxgets visible. Select a cell of the created table, the previously created table should be loaded in the tableeditor, tamplimüport combobox should disapper.
          * 7.4.1.3 select an empty cell in Excel, insert two dimensions (dim1, dim2) and two measures (mea1, mea2) resulting in the following list (dim1,dim2,mea1,mea2). Now drag mea2 to the top of the list an dim 1 to the end (mea2,dim2,mea1,dim2) and accept the table, select a empty cell, select a cell of the table, validate if the tableditor matches the table, move all columns back in the original order (dim1,dim2,mea1,mea2) and acccept, check if the table is updated correct.
          In the section Sorting change some sortorder and check
          
        * 7.4.2 Pivot mode on
          * 7.4.2.1 select an empty Cell in Excel, in the tableeditor switch to pivot mode, add a dimension for row and one for column an one measure accept with the green tick, table should be created, select en empty cell, tableeditor shoeuld be cleared. Now select a cell of the created table, check the ableeditor, add a dimension for rowand drag it to the top of the rows dimensions, then accept.  
          
          * 7.4.2.2 A Warning message should be shown if one attempt to create/import a table with:
            * more than one column-Dimensions
            * no row dimension
            * more than one measure         
          
      * 7.5 Table Misc.
        * 7.5.1 Error in dimension formula 
          * 7.5.1.1 create a table insert a dimension. Resolve dimension to filed via the chain-button next to the dimension-Name.          
          * 7.5.1.2  change the formula so that it has errors (use a field that didn't exists in Qlik,...)
          * 7.5.1.3  Table should collapse to one header row containing '-' and one empty data row. 
        * 7.5.2 Error in Measure Formula 
          * 7.5.2.1 create a table insert a measure. Resolve dimension to filed via the chain-button next to the measure-Name.
          * 7.5.2.2  change the formula so that it has errors (use a field that didn't exists in Qlik,...)
          * 7.5.2.3  The table should show normaly but in the measure-column witch contains errors, all rows should show a '-'
        * 7.5.3 if a table is inserted an the Excel-format of a column of this table is set, this format should also be set after the table is changed via the tableeditor
        * 7.5.4 If no foreground/backgroundcolor is set explicitly, the fore/back-color of a cell within a table should change according to the selected table-style. (light tablestyle uses black foreground color, dark tablestyle uses white foreground,...)
        * 7.5.5 if a Excel-totalrow is attached to a table, it should not be overridden after a selection changes the number of rows of the table.  
          
