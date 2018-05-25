# Zebra-TLP2824-Printing-from-Excel
This Project was aimed at printing barcodes from a Bill of Material using an old TLP2824. There were some challenges, but a use of the EPL language and straight out prining to the device was essential. 
The project used a DLL created in VS 2017 Community to run the code that sends EPL tot the printer. The DLL is accessed by the Excel Spreadsheet via VBA code, importing the DLL in the usual way. 
The most difficult part was the simplest. Using the test program from Zebra provided the clue. In order to treat the printer like a file and use the Windows CreateFile subroutine on it, the proper "name" of the printer had to be known. The printer was connected via USB and identified as "USB001". But that was not good enough, and so this was more than just sending "raw" to a COM port. 
The clue was provided by Developer-Demo_Windows.exe in D:\Program Files (x86)\Zebra Technologies\link_os_sdk\PC-.NET\v2.13.898\demos\Release. Notably, the "magic string" was
     LPTPORT = "\\?\usb#vid_0a5f&pid_0016#21j113600300#{28d78fad-5a12-11d1-ae5b-0000f803a8c2}"
And this was used in the variable LPTPORT in the EPL demo from https://support.zebra.com/cpws/docs/zpl/zpl_elp_vbnet.htm
For a C++ program, the string needed to be different. This is the sort of thing to be wary of in the world of "raw printing fu". 
     char printer[] = "\\\\.\\usb#vid_0a5f&pid_0016#21j113600300#{28d78fad-5a12-11d1-ae5b-0000f803a8c2}";
