             
'This program will generate a serial / registration number for programs that you
'develop.  It also maintains a database of users and registration numbers.  In
'addition, it writes and sends 'reg' files to your users via MS Outlook Express.

'Currently the database, PrgmRegistration.mdb,contains two tables;
'tblProgram1Reg and tblProgram2Reg.  You can, and should change the names
'of these tables to the names of various software that you want to have your
'users register.

'NB! in the Subroutine cmdOpenMAPI_Click, you must change the last lines in the
'MAPIMessage1.MsgNoteText statement to reflect your own name and email address.
'**************************************************************************************************************************
'& "YourName" & vbCrLf & vbCrLf & "youremail@any.com"
'**************************************************************************************************************************
'In the Subroutine BuildRegFile:
'   The Select Case cboProgram.ListIndex is currently set to register x number of programs,
'   change as necessary to accomodate the number of programs you wish to register.
' You must also change the MainKey and SubKey variables to reflect where their values are
' to be stored in the Registry.
'**************************************************************************************************************************

'If you have any questions, you can contact me by email at: jpcunningham@cox.net



Web: http://members.cox.net/johnpc7/
 
