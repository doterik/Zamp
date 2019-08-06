Add-in Express (TM) 2010 for Microsoft Office and .NET samples - readme.txt
http://www.add-in-express.com

-------------------------------------------------------------------------------
This file describes the sample projects, which are located in the following
folders:

- Advanced Outlook Regions
- Documentation Samples
- Other Samples
- Outlook Security Manager
- Toolbar Controls for Microsoft Office

All sample projects are available in four variations:
- VB.NET solution for VS 2005
- C# solution for VS 2005
- VB.NET solution for VS 2010
- C# solution for VS 2010

Below are listed the folders and sample projects that each of the folders
contains.


1. Advanced Outlook Regions
===========================
The projects below show how to develop advanced Outlook regions.

Context-Dependent Regions
-------------------------
This sample project creates a new Outlook folder named "Add-in Express" and
adds two sub-folders to it; the sub-folders are named "Explorer" and
"WebViewPane". The following regions are shown in the following contexts:
- For the Explorer folder, the RightSubpane explorer region is shown.
- For the WebViewPane folder, the WebViewPane explorer region is shown.
- For a mail item, the subject of which contains a specified string, the
  BottomSubpane inspector region is shown.

Hello World Sample
------------------
An introductory-level sample project demonstrating how to create an
advanced Outlook region, which you can drag to specified positions.


2. Documentation Samples
========================
This folder contains sample projects described in the manual, see
{Add-in Express}\Docs\adxnet.pdf on your PC.
The projects are located in these subfolders:

- Your First Excel Automation Add-in
- Your First Excel RTD Server
- Your First Microsoft Office COM Add-in
- Your First Microsoft Outlook COM Add-in
- Your First Smart Tag
- Your First XLL Add-in

Your First Excel Automation Add-in
----------------------------------
This is a sample Excel Automation add-in (Excel UDF) project. A test
workbook is provided.

Your First Excel RTD Server
---------------------------
This is an introductory-level project demonstrating how to create an RTD
server that handles a topic (string or set of strings identifying a data
source). A test workbook is provided.

Your First Microsoft Office COM Add-in
--------------------------------------
This solution demonstrates how to:

- Develop an add-in supporting several Office applications (Excel, Word and
  PowerPoint 2000-2010).
- Create command bars and command bar controls in Office 2000-2003.
- Create Office 2007-2010 Ribbon controls.
- Create advanced task panes for Office 2000-2010.
- Handle application-level events.

Your First Microsoft Outlook COM Add-in
---------------------------------------
From this example you will learn how to:

- Develop an add-in that works in Outlook 2000-2010.
- Create Outlook-specific command bars and command bar controls.
- Create Ribbon controls in Outlook 2007-2010.
- Create advanced regions for Outlook 2000-2007.
- Add custom property pages to the Folder Properties dialog.
- Handle application-level events.
- Handle events of the Items collection (add, remove, or change Outlook Items).
- Handle keyboard shortcuts.

Your First Smart Tag
--------------------
This sample project creates a smart tag providing a custom action for
predefined words or phrases. The smart tag supports Office 2002-2010. See also
the CustomSmartTagRecognizer sample in the Other Samples folder.

Your First XLL Add-in
---------------------
This sample demonstrates how to:

- Create an XLL (Excel UDF).
- Add it to a custom function category.
- Check parameters passed to the UDF.
- Determine if the UDF is called from the Insert Function dialog.
- Return an Excel error value.

A test workbook is provided.


3. Other Samples
================
This folder contains a collection of samples that do not fit into any other
category.

CustomSmartTagRecognizer
------------------------
This is how you develop a smart tag recognizing a pattern, not a predefined
word or phrase.

ExcelAutomationAddin
--------------------
This sample demonstrates how to access an Excel Range object passed to your
Excel Automation add-in. A test workbook is provided.

ExcelTimesheet
--------------
Creating an Excel worksheet and controlling MS Forms COM components on it.
A test Excel template is provided.

OutlookContextMenus
-------------------
Adding custom items to CommandBar-based context menus in Outlook 2002-2007 and
Ribbon-based context menus in Outlook 2010.

OutlookFolderItemsEvents
------------------------
Connecting to the events of the Items collection in Outlook 2000-2010.

OutlookFoldersEvents
--------------------
Connecting to the events of the Folders collection in Outlook 2000-2010.

OutlookItemEvents
-----------------
Connecting to the events of an Outlook item in Outlook 2000-2010.

OutlookPropertyPage
-------------------
Creating a custom page for the Folder Properties dialog in Outlook 2000-2010.

RTDServerStock
--------------
An advanced RTD server sample project. A test workbook is provided.

WordFax
-------
Creating a Word document and controlling MS Forms COM components on it. A test
Word template is provided.


4. Outlook Security Manager
===========================
This is a set of samples demonstrating how to use the Outlook Security Manager
to switch off security warnings in Outlook 2000-2010.

COM Add-in
----------
An Add-in Express based Outlook COM add-in project adding a custom
CommandBar/Ribbon button that switches off or on security warnings in Outlook
2000-2010. Other buttons execute test functions that may or may not raise
security warnings depending on the state of the first button.

Send Mail Sample
----------------
This sample project is a Windows Forms application that sends e-mails using
the Outlook object model (Outlook versions 2000-2010). The application may or
may not raise security warnings depending on the state of the
DisableOOMWarnings property of the Outlook Security Manager.

Simple Application
------------------
This Windows Forms application gets some information about the first mail
item in the Inbox folder using:
a) Outlook object model (Outlook 2000-2010)
b) CDO (Outlook 2000-2007)
c) Simple MAPI (Outlook 2000-2010)

The application may or may not raise security warnings depending on the
state of the corresponding properties of the Outlook Security Manager.


5. Toolbar Controls for Microsoft Office
========================================

Excel Add-in Sample
-------------------
This sample add-in project shows several .NET controls on a toolbar in Excel
2000-2003.

Multiple Hosts Sample
---------------------
This sample add-in project shows a .NET UserControl on a toolbar in Word
2000-2003, Excel 2000-2003, PowerPoint 2000-2003 and Outlook 2000-2007.

Outlook Add-in Sample
---------------------
This sample project shows several .NET controls on a toolbar in Outlook
2000-2003.

-------------------------------------------------------------------------------
Copyright (C) Add-in Express Ltd. All rights reserved.
