Spread README    
FarPoint Technologies, Inc.     
August 19, 1996    
     
*********************************************************    
This README file is for Spread (TM) version 2.5, release date      
August 19, 1996     
     
The following topics are available:     
  Release Notes      
  Additional Information for C and C++ Users     
  Contacting FarPoint Technologies, Inc.    
  Comment Form            
     
     
*********************  Release Notes  *********************    
    
======  Installation  =====    
    
Refer to the INSTALL.WRI file for installation instructions and a list     
of installed files.    
    
=====   Redistribution  =====    
    
If you plan to redistribute Spread controls, please refer to the     
REDIST.TXT file for a list of files you must distribute with your     
application. The REDIST.TXT file is installed in your \SPREAD25     
directory.     
    
=====  Upgrade Instructions  =====    
     
If you are upgrading from existing projects that used either     
Spread/VBX version 2.1, Spread/VBX++ version 2.0, ToolBox version     
3.0, of Grid/VBX version 1.0 controls, refer to the UPGRAD.TXT file.     
The UPGRAD.TXT file is installed in your \SPREAD25 directory.    
    
If you are upgrading from an existing project using Spread/VBX or    
Grid/VBX to use one of the Spread version 2.5 OCX controls, you    
must update your event parameters. Refer to the EVENTS.TXT file    
installed in your \SPREAD25 directory.    
   
======  Documentation ======  
  
The Spread User's Guide is provided in print form and in the online   
help file. The Spread DLL Reference Guide and OCX/VBX Reference   
Guide are included in their respective online help file.  
   
======  Online Help  ======    
            
  If you are using Windows 95, contents files (SSDLL25.CNT and      
  SSOVX25.CNT) are included.           
     
  If you want online help for the Spread Designer, you must install     
  the OCX/VBX help file.     
     
  If you are using the 32-bit OCX control in Windows 95 in Visual      
  Basic, to access online help for events, type the event name in the      
  code window and press F1.     
     
======  Technical Issues Addressed -- Release Date 8/19/96  ====== 
 
The information in parentheses following each item is our company's     
tracking number, which you may have been given in response to a     
problem you reported.        
        
 - In Spread Designer, you can now display previously hidden columns 
   or rows by setting the column or row width to -1. (JAP5179)

 - When the ScrollBarMax property is set to False, the vertical scroll 
   bar range is 1 to either 20 or the DataRowCnt property value, 
   whichever is greater. The horizontal scroll bar range is 1 to either 8 
   or the DataColCnt property value, whichever is greater. (JIS338)

 - The Spread control no longer GPFs when you populate a combo box 
   cell with a string larger than 32767 characters. (KEM47)

 - The BackColorStyle property now works properly for static or owner 
   drawn cell types. (SCS3511)

 - Values entered in a float cell with the TypeFloatDecimalPlaces 
   property set to 2 and the TypeFloatSeparator property set to True no 
   longer display a 0. (SCS3591) 
  
====== Technical Issues Addressed -- Release Date 8/13/96  ======    
    
 - The Spread control DLL now shows in Borland's Resource   
   Workshop. (GRB1618)  
  
 - Customizing header text using the Spread Designer is now easier to      
   accomplish. (KEM3522)     
     
 - The ColHidden check box in the Spread Designer now works      
   correctly. (KEM3749)     
  
 - The DataAddNew event now fires correctly when updating an empty   
   database table. (KEM3762)  
     
 - When the Spread control has the focus, closing an application by      
   clicking on the icon in the upper left corner no longer causes a GPF.      
   (JAP4544)     
     
 - The GetCellFromScreenCoordinates now returns the correct      
   coordinates for a cell in column header row or the row header      
   column. (JAP4677)     
     
 - Float cells with zero decimal places no longer truncate the decimal      
   value. (KEM3755)     
     
 - Choosing any option from the Unit Type dialog box in Spread      
   Designer now works correctly.(SCS3144)     
     
 - When the EditModePermanent property is set to True and the last      
   column is locked, you can now scroll to the end of the spreadsheet.      
   (SCS3218)     
  
 - The 16-bit OCX Spread control now correctly writes data back to a   
   database. (SCS3319)  
  
 - The memory leak problem is now fixed. (SCS3383)  
     
 - Binding a column of numeric data to a column of static text cells no      
   longer causes a GPF when the form is loaded. (SEL5286)     
     
 - In Windows 3.x, unloading a Spread VBX control on an MDI child      
   form using a button in a Visual Basic picture box, a Visual      
   Basic frame, or MicroHelp frame control no longer causes a GPF.      
   (SEL5553)     
  
 - If you are using Visual C++ 4.x and you have a frame window   
   around the form view, buttons no longer require two key presses to   
   be activated. (SEL5713)  
  
 - When using a custom function inside the parameter list of another   
   custom function, characters are no longer lost from the formula   
   string. (SEL5720)  
  
 - The truncate function of the Formula property now calculates   
   floating point numbers with better precision. (SEL5799)  
     
 - Long (multiple line) formulas are no longer truncated. (TEL2937)     
    
=====  Technical Issues Addressed -- Release Date 7/10/96  =====    
    
 - The online help now correctly documents the ButtonDown  
   parameter for the ButtonClicked event as an integer value.  
   (GRB1793)     
     
 - Cells that overflow now print correctly. (GRB1964/1965)     
     
 - Float cells now round correctly to the specified number of decimal      
   places. (KEM3709)     
     
 - When the Spread control is bound to a Microsoft Access 7.0      
   database, the database is now updated when data is written to      
   it. (KEM3715)     
     
 - You can now set the font to Wingdings. (SCS2929)     
     
 - Changing the background or text color of cells in frozen rows or      
   columns no longer causes a painting problem. (SEL2)     
     
 - In Windows NT 3.5, the Spread control no longer causes a GPF      
   when you set the row height equal to the value of the      
   MaxTextRowHeight property while looping through more than nine      
   rows.(SEL5105)     
     
 - Picture cells now paint properly. (SEL5397)     
     
 - The Spread control no longer GPFs when loading a spreadsheet that      
   contains custom functions in the Form_Load event. (SEL5403)     
    
======  Technical Issues Addressed -- Release Date 6/10/96  ======    
    
 - When using the 32-bit OCX control in Visual C++ 4.0, all changes       
   made by the Spread Designer are now updated in the Resource       
   Compiler.     
     
 - You can now use the Visual Basic 4.0 Setup Wizard to register the     
   Spread control.      
       
 - When a cell is in edit mode and a message box is displayed by          
   clicking the right mouse button, the Spread control now receives the          
   focus after the message box is closed. (GRB1617)         
            
 - The Action property setting 8 (Load) now works correctly in Visual          
   Basic 4.0. (GRB1834)         
            
 - Rows and columns that frozen are no longer duplicated. (GBR1847)         
            
 - The horizontal and vertical scroll bars now work properly when the          
   ScrollBarMaxAlign property is set to True and columns are          
   hidden. (GRB1861)         
            
 - Data is now displayed in integer fields of bound spreadsheet          
   controls. (JAP3696)         
            
 - The GetClientArea and GetCellPos functions now return the height          
   and width values in twips. (JAP4045)         
            
 - When you display an icon in a picture cell whose background color          
   is set to vbWindowBackground, the background no longer turns          
   black. (JAP4167)         
            
 - In Visual Basic 3.0, the combo box cell now works correctly when          
   Spread version 2.1 is loaded before Spread version 2.5. (JAP4176)         
            
 - Picture cells no longer cause memory leaks in Visual Basic 3.0.          
   (JAP4201)         
            
 - For the Spread OCX control, the Back Color Displays drop-down list          
   box on the Display property page now includes the Over Horz Grid          
   Only and Over Vert Grid Only list items. (JIS3266)         
            
 - In a single select spreadsheet, invoking a drag method no longer          
   causes the spreadsheet to scroll with mouse movement. (KEM3548)         
            
 - When the Drag Mode property is set to 1 (Automatic), the scroll bars          
   are now correctly displayed. (KEM3548)         
           
 - When working with the Spread VBX control in Visual C++, using          
   the SetText function to set the header text no longer causes a          
   SmartHeap memory error. (KEM3636)         
            
 - Setting the Action property 0 (Activate Cell) now updates the record          
   pointer. (SCS2739)         
            
 - The Text property now works properly when used within the          
   KeyDown event. (SCS2798)         
            
 - Spreadsheets saved using the old Interface Designer (.SS files) are          
   now loaded correctly in the new Spread Designer. (SEL3850)         
            
 - When the DAutoSave property is set to False, the Action (15 - Data          
   Save) property now correctly saves inserted and modified rows in a          
   bound spreadsheet. (SEL3984)         
            
 - The MS Sans Serif 9.75 point font is now displayed correctly.          
   (SEL5080)        
       
 - When using Windows NT, the TypeIntegerMax and TypeIntegerMin        
   properties are now processed much faster. (SEL5100)        
         
 - Formulas are now updated correctly after sorting rows. (TEL1994)         
            
======  Technical Issues Addressed -- Release Date 5/10/96  ======    
    
 - Within the MouseDown event, the GetCellFromScreenCoord        
   function now returns the correct values for the Col and Row        
   parameters. (BOJ3070)       
       
 - The Spread VBX control no longer GPFs when you set the DataField        
   property for the row header column (column 0). (BOJ3703)       
       
 - The KeyPress event will now process the pressing of the Enter key.        
   (GRB814)       
       
 - Formulas that involve columns are now updated correctly when the        
   value of a cell is changed in code.(GRB1623)       
       
 - You can now assign the value of the Picture property for another        
   control (Image, for example) to the TypePictPicture,        
   TypeButtonPicture, TypeButtonPictureDown, or TypeCheckPicture        
   properties. (GRB1683)       
       
 - When you set the background or foreground color of a block of cells,        
   the color for that block now stays set when you change the        
   background or foreground color of the spreadsheet. (JAP3695)       
       
 - When you set the background color of a spreadsheet and you display        
   only the horizontal (vertical) lines, you can now set the back color        
   style to avoid spaces where the vertical (horizontal) lines cross the        
   horizontal (vertical) lines. This can be set through the Spread      
   Designer, the properties window, or in code. (JAP3757)       
       
 - The row height value no longer decreases every time you reset it.        
   (JAP3819)       
       
 - When you have a command button (Default property = True) and a        
   spreadsheet on a form, pressing the Enter key now fires the command        
   button's Click event. (JAP3853)       
       
 - Check box cells now display the correct data and picture when the        
   Spread control is bound to a Microsoft Access database table.        
   (JAP3883)       
       
 - The Button parameter in the MouseMove event no longer returns a 1        
   when the mouse buttons are up. (JAP3957)       
       
 - The Spread Designer now unlocks locked cells correctly. (JIS3201)      
      
 - The Spread Designer no longer GPFs intermittently when you apply       
   or exit the Spread Designer.(JIS3226)       
       
 - The ProcessTab property now works correctly when the        
   EditModePermanent property is True. (KEM3491)       
       
 - The PrintType property, Cell Range setting now works correctly.        
   (KEM3508)       
       
 - The scroll box now resizes correctly based on the maximum number        
   of rows in the spreadsheet. (KEM3520)       
       
 - In the Spread Designer, the User Defined Text portion of the Headers        
   dialog box has been redesigned and now works correctly.        
   (KEM3522)       
       
 - Invoking a Drag method for a Spread control in single select        
   operation mode no longer causes the spreadsheet to scroll with mouse        
   movement. (KEM3548)       
       
 - The NoBeep property now works correctly. (KEM3618)       
       
 - The SSSetFont and SSSetFontRange functions now work correctly        
   when the column number is greater than 32767. (SCS2506)       
       
 - When you set a color property to a system default color at design        
   time, the system default color is no longer converted to a hex        
   constant. (SCS2509)       
       
 - The TypeCheckCenter property now works correctly when set for a        
   range of cells. (SEL3466)       
       
 - When the Spread control is in row mode or single select mode, the        
   RetainSelBlock property is False, and you change the active cell,        
   the new row is now highlighted.(SEL3553)       
       
 - The drop-down list portion of multiple combobox cells on a        
   spreadsheet now displays correctly. (SEL3639)       
       
 - The SelBlockCol, SelBlockCol2, SelBlockRow, and SelBlockRow2,        
   and the SSGetSelectBlock and SSGetMultiSelBlock functions now        
   return a value of -1 when an entire row or an entire column is        
   selected. (SEL 3651).       
       
 - The popup calendar in a date cell is now modal when the Spread        
   control is in a modal dialog. (SEL3754)       
       
 - The ButtonClicked event now fires for a button cell and correctly        
   reflects the ButtonDown parameter as an integer type. (TEL1809)       
       
 - The correct number of Change, EditMode, and LeaveCell events are        
   fired now when EditModePermanent is set to True. (TEL1844)       
       
 - Integer fields in a spreadsheet bound to a Microsoft Access database        
   table now display correct data. (TEL1962)       
       
 - In the Spread Designer, the Multiline check box on the Edit Cell        
   Type dialog box now works correctly. (TEL1963)       
       
 - The MaxTextRowHeight property now returns the correct row height        
   for cells with short text lines and for static text cells. (TEL1977)      
      
 - The Lock property is now working properly for the VBX control.       
   (TEL2107)       
     
    
********* Additional Information for C and C++ Users **********    
     
If you are using the Spread control in either C or C++, review the      
following information about using the control.     
     
===== Using the Spread Version 2.5 VBX Control in C++ =====    
     
Due to the complexity of the Spread control, the C++ implementation     
files are large. To use the Spread version 2.5 VBX control in C++,     
you might have to break the SSVBXMFC.CPP or SSVBXOWL.CPP      
files into two files if your C++ compiler cannot handle large files.     
     
== Using the Spread Version 2.5 32-Bit DLL Control in Borland C ==    
     
To use the Spread version 2.5 32-bit DLL control in Borland C,      
you must use version 5.0 or greater of Borland C's IMPLIB.EXE      
to generate a valid TAB32D20.LIB or IMP32D20.LIB. The IMPLIB      
command does not generate a backward-compatible library for      
earlier versions of Borland C (that is, for versions previous      
to version 5.0).     
     
== Using the Spread Version 2.5 32-Bit OCX Control in MSVC 4.0 ==    
     
To use the Spread version 2.5 32-Bit OCX control in MSVC 4.0,     
you must generate wrapper classes for the controls.     
     
To generate a wrapper class, you must use an existing application     
that supports OLE controls, or you must create such an application.     
     
Complete the steps in the following section to create an application     
that supports OLE controls. If you already have such an application,      
proceed to the section "Generating a C++ Wrapper for the Spread      
Control in MSVC 4.0."     
     
==== Creating a New Application That Supports OLE Controls ====    
     
To generate a new application that supports OLE controls     
     
1. In MSVC 4.0, choose New from the File menu.     
2. In the New dialog box, select Project Workspace from the New list.     
3. Choose the OK button.     
4. In the New Project Workspace dialog box,     
   a. Select the default item (MFC AppWizard (exe)) from the Type list.     
   b. Type a project name into the Name box.     
   c. Type a location for your project files in the Location box     
      or choose the Browse button to select a drive and directory.     
   d. Choose the Create button.     
5. In the MFC AppWizard - Step 1 dialog box, choose the Next button.     
6. In the MFC AppWizard - Step 2 of 6 dialog box, choose the Next      
   button.     
7. In the MFC AppWizard - Step 3 of 6 dialog box, select the OLE      
   controls check box and choose the Finish button.      
   This accepts the defaults for all other choices.     
8. In the New Project Information dialog box, choose the OK button      
   and MSVC 4.0 generates an application that supports OLE controls.     
     
== Generating a C++ Wrapper for the Spread Control in MSVC 4.0 ==    
     
There are two methods for generating C++ wrappers in MSVC 4.0.     
You can place the control on a dialog, or you can add the control     
as an application component.     
     
     
To generate a C++ wrapper with the control on a dialog     
     
1.  Open an existing dialog or create a new dialog in MSVC 4.0.     
    To create a new dialog, select Resource from the Insert menu.      
    Then, in the Insert Resource dialog box, select Dialog from the      
    Resource type list, and then choose the OK button.     
2.  Click the right mouse button in the dialog to display the pop-up      
    menu.     
3.  Choose Insert OLE Control.     
4.  In the Insert OLE Control dialog box, select FarPoint Spread      
    Control from the OLE control list.     
5.  Choose the OK button.     
6.  Click the right mouse button in the dialog to display the pop-up      
    menu.     
7.  Choose ClassWizard.     
8.  If you are creating a new dialog, create a C++ class for the dialog      
    using the following steps:     
    a. In the Adding a Class dialog box, choose the OK button.     
    b. In the Create New Class dialog box under Class Information, type     
       the new name in the Name box.     
    c. Choose the Create button.     
9.  Create a member variable for the control you inserted into your     
    dialog using the following steps:     
    a. In the MFC ClassWizard window on the Member Variables tab,      
       select the control ID of the OLE control you added to the dialog      
       from the Control IDs list.     
    b. Choose the Add Variable button.     
    c. Choose the OK button in the Microsoft Developer Studio message      
       that explains that a class wrapper will be generated for you.     
10. In the Confirm Classes dialog box, you can modify the class name,      
    implementation filename, and the declaration header filename or      
    accept the defaults provided.     
11. Choose the OK button.     
12. In the Add Member Variable dialog box, type a variable name in      
    the Member variable name box.     
13. Choose the OK button.     
    ClassWizard generates a wrapper class to support the Spread control.     
     
      
To generate a C++ wrapper with the control as an application      
component     
     
1. In MSVC 4.0, choose Component from the Insert menu.     
2. In the Component Gallery dialog box, choose the OLE Controls tab.     
3. Select the FarPoint Spread Control from the list.     
4. Choose the Insert button.     
5. In the Confirm Classes dialog box, you can modify the class name,      
   implementation filename, and the declaration header filename or      
   accept the defaults provided.     
6. Choose the OK button.     
7. In the Component Gallery dialog box, choose the Close button.     
   ClassWizard generates a wrapper class to support the Spread control.     
     
     
************ Contacting FarPoint Technologies, Inc. ************    
       
If you discover a problem with either the software or the accompanying       
documentation, or if you would like to share your thoughts about this       
product with FarPoint Technologies, please send us a completed       
Comment Form.       
       
Send us the Comment Form using one of the following methods, or if      
you call, please provide the information requested on the form to our       
technical support representative. For additional assistance, contact our       
Technical Support department.       
       
To contact FarPoint Technologies, Inc.       
       
- Call our Technical Support department at (919) 460-1887.        
  FarPoint Technologies' Technical Support department is available       
  between the hours of 9:00 a.m. and 5:30 p.m. EST.       
       
- Fax us at (919) 460-7606.       
       
- Post comments or questions to the Technical Support forum on our       
  Bulletin Board Service at (919) 460-6712 (28.8K bps) or       
  (919) 460-5771 (1200 to 14.4K bps).       
       
- Contact the FarPoint forum on CompuServe (type GO FARPOINT).       
       
- Write us at:       
   FarPoint Technologies, Inc.       
   133 Southcenter Court       
   Suite 1000       
   Morrisville, NC 27560       
       
- Send e-mail to us at farpoint@fpoint.com (general mail) or      
  fpsupport@fpoint.com (technical support).       
       
You can also contact FarPoint Technologies at       
       
- Our World Wide Web site (http://www.fpoint.com/fpoint)       
       
- Our FTP site (ftp.fpoint.com and change to the directory /fpoint.com)       
     
     
******************  Comment Form  *********************    
            
If you would like to send comments about this release, complete             
and fax this form to FarPoint Technologies at (919) 460-7606. Please             
include a detailed description of the problem along with the steps             
necessary to reproduce it. The preferred method is a sample application             
that reproduces the problem. If you have a sample, please upload it to             
our BBS at (919) 460-6712 (28.8K bps) or (919) 460-5771 (1200 to             
14.4K bps), to our CompuServe forum, or to our Internet address. Be             
sure to include the name of the ZIP file* that you uploaded.            
            
Today's date: _______________             
            
Product: Spread		Version: 2.5.010            
            
Customer name: ________________            
     
Company name: ___________            
            
Phone number: (_____)____________            
     
Fax number: (_____)___________            
     
E-mail address: _________________            
            
*Sample application file name:_____________________            
                   
Control Type (circle one or more):             
16-bit DLL      32-bit DLL      VBX      16-bit OCX      32-bit OCX            
            
Development environment (i.e. Visual Basic, Visual C++,            
Borland C++): ____________	Version: ______            
            
Description of the problem:            
            
            
            
Steps necessary to reproduce the problem:            
            
            
            
Comments:            
            
            
            
Documentation suggestions:
