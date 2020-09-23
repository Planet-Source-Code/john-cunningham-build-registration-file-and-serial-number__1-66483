VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmBuildRegFile 
   Caption         =   "Build Registry File"
   ClientHeight    =   9165
   ClientLeft      =   3105
   ClientTop       =   2385
   ClientWidth     =   9630
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdKillAllRegnTextFiles 
      Caption         =   "Kill All Reg and Text Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3368
      TabIndex        =   20
      Top             =   5640
      Width           =   2895
   End
   Begin VB.ComboBox cboProgram 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   3900
      TabIndex        =   18
      Text            =   "        Select Program"
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear Input"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   17
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdRunRegFile 
      Caption         =   "&Run Reg File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4140
      TabIndex        =   16
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton cmdShowRegistry 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   3900
      Picture         =   "frmBuildRegFile.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "View Registry"
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdShowDataBase 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   1860
      Picture         =   "frmBuildRegFile.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Launch MS Access"
      Top             =   4680
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2355
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4154
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1635
   End
   Begin VB.CommandButton cmdOpenOLExpress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   5580
      Picture         =   "frmBuildRegFile.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Open Outlook Express"
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton cmdOpenMAPI 
      Caption         =   "&Email Registration"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4141
      TabIndex        =   3
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox txtRFName 
      Height          =   285
      Left            =   720
      TabIndex        =   10
      Top             =   4440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtEmailAddress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3735
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.CommandButton cmdGenerateSN 
      Caption         =   "&Generate Registration Number - Build Reg File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3301
      TabIndex        =   2
      Top             =   1680
      Width           =   4095
   End
   Begin VB.CommandButton cmdOpenRegFiles 
      Caption         =   "&View Reg Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1981
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdKillRegFile 
      Caption         =   " &Kill *.Reg File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox txtKey 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3736
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3735
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin MSMAPI.MAPIMessages MAPIMessage1 
      Left            =   8040
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   8040
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Label lblPrgSelected 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3840
      TabIndex        =   19
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      Height          =   2895
      Left            =   1740
      Top             =   120
      Width           =   6135
   End
   Begin VB.Shape Shape4 
      Height          =   972
      Left            =   1752
      Top             =   4560
      Width           =   6132
   End
   Begin VB.Shape Shape3 
      Height          =   615
      Left            =   1755
      Top             =   3840
      Width           =   6135
   End
   Begin VB.Shape Shape2 
      Height          =   615
      Left            =   1755
      Top             =   3120
      Width           =   6135
   End
   Begin VB.Label Label3 
      Caption         =   "Email Address:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2175
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Registration Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1980
      TabIndex        =   6
      Top             =   2160
      Width           =   1572
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2220
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmBuildRegFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
            '**********************************************************
            '**       Build Program Registration Files               **
            '**       John P. Cunningham - 09/6/2006                 **
            '**       Web: http://members.cox.net/johnpc7            **
            '**       Email:  jpcunningham@cox.net                   **
            '**********************************************************
            
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

'*********************************************************************************************************************************
'The following API serves to open the Help File and the Access Database
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                (ByVal hwnd As Long, ByVal lpszOp As String, _
                 ByVal lpszFile As String, ByVal lpszParams As String, _
                 ByVal LpszDir As String, ByVal FsShowCmd As Long) _
                 As Long
'*********************************************************************************************************************************

                          '**** Form Level Declarations ****

'*********************************************************************************************************************************
'                Make sure to open Project-References and select
'                Microsoft ActiveX Data Object Library 2.x or higher.
'*********************************************************************************************************************************
Dim DbFile As String                        'Name of DataBase

Dim cn As ADODB.Connection                  'Connect to the ADO Data Type
Dim rs As ADODB.Recordset                   'Record Source Name
Dim SQLstmt As String                       'SQL Statement String(s)
Dim RcCount As Integer                      'Record Counter

Dim intID As Integer                        'Integer for the Record Counter
Dim strName As String                       'Variable to hold Customer's Name
Dim strEmailAddress As String               'Variable to hold Customer's Email Address
Dim strRegistration As String               'Variable to hold Customer's Registration Number
Dim StrTDate As String                      'Variable to hold the date
Dim PrgmName As String                      'Variable to hold the Program's Name
Dim RetVal As Variant                       'Return Variable for Message Boxes

Dim MainKey As String                       'Main Key of Registry Storage
Dim SubKey As String                        'Sub Key of Registry Storage
Dim FileToAttach As String                  'File Attachment Name to Email to Customer


Const AppFile As String = "\BuildReg Help File.chm"  'Name of the Help File here ONLY!
Dim ExecPath As String                               'Path to Help File

Public Function GenKey(Username As String) As String
'This Function generates a program serial number by converting the Username (txtName)
'into a string that is numeric
Dim TVal As Long
Dim i As Integer
Dim TText As String
Dim TString As String
Dim kk As Integer

'**********************************************************************************************
'Set up values to generate the Serial Numbers
'For each one of your programs change kk accordingly
'Use any value for kk except 0 and change it's value for each program in cboProgram
    Select Case cboProgram.ListIndex
        Case 0
            kk = 16
        Case 1
            kk = 18
        Case 2
            'kk =
        Case 3
            'kk =
    End Select
'**********************************************************************************************

TString = "" 'Reset the variable
        
For i = 1 To Len(Username) ' Start the loop using the length of the username
'**********************************************************************************************
                        
                TVal = Asc(Mid(Username, i, 1)) + kk 'Converts the next letter of username to
                                                     'it's ASCII value, then adds kk
                TVal = TVal + Fix((TVal * (127 + Len(Username)))) 'This adds the
                                'last result with kk * the length of the username
                TVal = TVal + Len(Username) 'It adds to the last result the length of
                                            'the Username
                TString = TString & Trim(StrReverse(Str(TVal)))  'This reverses
                'the last result and appends it to the last result in TString

'*********************************************************************************************
                
Next i 'Continue getting the next letter in Username
    
TText = TString 'This puts the generated key into TText

'*********************************************************************************************
If Len(TText) >= 8 Then 'This tests to see if the length of the key is 8 or greater
    Mid(TText, 4, 1) = "-" 'If so then place a hyphen in the key
    Mid(TText, 12, 1) = "-" 'Place another hypen in the key
End If
'*********************************************************************************************

TText = Left(TText, 16) 'This trims the key down making it look nice
   
GenKey = TText
    
End Function

Private Sub BuildRegFile()
 On Error GoTo Build_Reg_Key_Error
  
' ---------------------------------------------------------------------------
' define local variables
' ---------------------------------------------------------------------------
  Dim hFile          As Integer
  Dim strfilename    As String
  Dim strFilename1   As String
  Dim RegLine1       As String
  Dim RegLine2       As String
  Dim RegLine3       As String
  Dim RegLine4       As String
  Dim RegLine5       As String
  Dim RegLine6       As String
  Dim Username       As String
  Dim SKey As String
  Dim RegistrationNumber As String
  
  strfilename = ""
  strFilename1 = vbNull
 
'**********************************************************************
'Currently set to register x number of programs, change as necessary
  Select Case cboProgram.ListIndex
        Case 0
            SKey = "Program 1"
        Case 1
            SKey = "Program 2"
        Case 2
            SKey = "Program 3"
        Case 3
            SKey = "Program 4"
   End Select
 '**********************************************************************
        Username = "User"
        RegistrationNumber = "Registration Number"
        '***************************************************************
        'Change the next two lines to suit your program
        MainKey = "Your Name"
        SubKey = SKey
         
        strfilename = "C:\Build Reg n SN Files\" & txtName & ".txt"
        strFilename1 = "C:\Build Reg n SN Files\" & txtName & ".reg"
        txtRFName = strfilename
        
        RegLine1 = "REGEDIT4"
        RegLine2 = ";Do not modify this file"
        RegLine3 = "\\Remember to change this files extension to ''.reg''"
        RegLine4 = "[HKEY_CURRENT_USER\Software\" & MainKey & "]"
       '******************************************************************************
        'Change the next line to suit your program
        RegLine5 = "[HKEY_CURRENT_USER\Software\" & MainKey & "\" & SKey & "]"
' ---------------------------------------------------------------------------
' Create txt file
' ---------------------------------------------------------------------------
        hFile = FreeFile
        Open strfilename For Output As #hFile
' ---------------------------------------------------------------------------
' Write data to txt file
' ---------------------------------------------------------------------------
        Print #hFile, RegLine1
        Print #hFile, " "
        Print #hFile, RegLine2
        Print #hFile, RegLine3
        Print #hFile, " "
        Print #hFile, RegLine4
        Print #hFile, " "
        Print #hFile, RegLine5
        Print #hFile, Chr$(34) & Username & Chr$(34) & "=" & Chr$(34) & txtName & Chr$(34)
        Print #hFile, Chr$(34) & RegistrationNumber & Chr$(34) & "=" & Chr$(34) & txtKey & Chr$(34)
        Print #hFile, " "
        Close #hFile
' ---------------------------------------------------------------------------
' Create REG file
' ---------------------------------------------------------------------------
       hFile = FreeFile
       Open strFilename1 For Output As #hFile
' ---------------------------------------------------------------------------
' Write data to REG file
' ---------------------------------------------------------------------------
        Print #hFile, RegLine1
        Print #hFile, " "
        Print #hFile, RegLine2
        Print #hFile, RegLine3
        Print #hFile, " "
        Print #hFile, RegLine4
        Print #hFile, " "
        Print #hFile, RegLine5
        Print #hFile, Chr$(34) & Username & Chr$(34) & "=" & Chr$(34) & txtName & Chr$(34)
        Print #hFile, Chr$(34) & RegistrationNumber & Chr$(34) & "=" & Chr$(34) & txtKey & Chr$(34)
        Print #hFile, " "
        Close #hFile

Exit Sub
 
' ---------------------------------------------------------------------------
' Go here only if there is an error
' ---------------------------------------------------------------------------
Build_Reg_Key_Error:
  MsgBox "Error:  " & CStr(Err.Number) & "  " & Err.Description & vbCrLf & vbCrLf & _
         "Error occured creating file:" & vbCrLf & vbCrLf & strfilename, _
         vbOKOnly, "ERROR  ERROR"
     Close #hFile
End Sub

Private Sub cboProgram_Click()

'If txtName = "" Or txtEmailAddress = "" Then Exit Sub

  Me.Height = 9000
  'Center Form on the screen
  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = (Screen.Height - Me.Height) / 2
 
'********************************************************************************

    'Set the Database Applicable Path
      DbFile = App.Path & "\PrgmRegistration.mdb"

       'Establish the Connection
       Set cn = New ADODB.Connection
       cn.CursorLocation = adUseClient
       cn.ConnectionString = _
             "Provider=Microsoft.Jet.OLEDB.4.0;" & _
             "Data Source=" & DbFile & ";" & _
             "Persist Security Info=False"

      'Open the Connection
      cn.Open
'*********************************************************************
Select Case cboProgram.ListIndex

    Case 0
    
        SQLstmt = "SELECT * FROM [tblProgram1Reg]Order by ID"
        lblPrgSelected.Caption = "Program 1 Selected"
        
    Case 1
        SQLstmt = "SELECT * FROM [tblProgram2Reg]Order by ID"
        lblPrgSelected.Caption = "Program 2 Selected"
        
    Case 2
        'Add SQL statement for Program 3
        'lblPrgSelected.Caption = "Program 3 Selected"
        
    Case 3
        'Add SQL statement for Program 4
        lblPrgSelected.Caption = "Program 4 Selected"
        
End Select
'*********************************************************************
Set rs = New ADODB.Recordset
rs.Open SQLstmt, cn, adOpenStatic, adLockOptimistic, adCmdText
        
Set DataGrid1.DataSource = rs

End Sub

Private Sub cmdClear_Click()

    txtName = ""
    txtEmailAddress = ""
    txtKey = ""
    txtName.SetFocus
    Set DataGrid1.DataSource = Nothing
    cmdRunRegFile.Enabled = True
    
End Sub

Private Sub cmdGenerateSN_Click()

If cboProgram.ListIndex = -1 Then
    RetVal = MsgBox("Select a Program", 48, "Build Registration File")
    Exit Sub
End If
    
'**************************************************************************************************************************
If txtName.Text = vbNullString Or txtEmailAddress = vbNullString Or cboProgram.Text = "Program" Then
    RetVal = MsgBox("Name and Email data must be filled in and Program must be selected.", 64, "Build Registry File")
    txtName.SetFocus 'Sets the focus on txtname
    Exit Sub 'Exits this sub
End If

'**************************************************************************************************************************
If Len(Trim(txtName.Text)) < 6 Then 'This checks the length of the text and makes sure it's not less than 6
    RetVal = MsgBox("Name must be at least 6 characters long!", 64, "Build Registry File")
    txtName.SetFocus 'Return to the form and set the focus on txtname
    Exit Sub 'Exits this sub
Else
    txtKey.Text = GenKey(Trim(txtName.Text)) 'If all is well, then call our function and place the result in txtkey
End If

'**************************************************************************************************************************
CreateNewDirectory "C:\Build Reg n SN Files"
BuildRegFile
cmdOpenMAPI.Enabled = True
AddNewRecord
'cmdKillRegFile.Enabled = True

End Sub

Private Sub cmdKillRegFile_Click()

RetVal = MsgBox("Are you sure you want to delete this file?  " & vbCrLf _
    & txtName.Text, 4, "Build Reg File")

Select Case RetVal

     Case 6     'Yes
          Kill "C:\Reg Files\" & txtName & ".txt"
          Kill "C:\Reg Files\" & txtName & ".reg"
   
     Case 7     'No
         Exit Sub
End Select
           
End Sub

Private Sub cmdOpenRegFiles_Click()
  
 Dim strCheckForFile As String
 Dim MstrFilePath As String
 Dim CompleteMstrFilePath As String
 
    'Give the file selection window a title.
    CommonDialog1.DialogTitle = "Select a Reg File"
    'The file selection window will start in the
    'applications directory.
    CommonDialog1.InitDir = "C:\Build Reg n SN Files"
      
     
    'Allow user to view only pertinent files.
    CommonDialog1.Filter = "All Files (*.*)|*.*"
                       
    'Open the file selection window.
    CommonDialog1.ShowOpen
    
    'Select the last four letters of the file selected.
    strCheckForFile = Right(CommonDialog1.FileName, 4)
   
    Select Case strCheckForFile
           
        Case vbNullString
            'Do not allow empty strings.
            Exit Sub
            
        Case ".txt"
        'The following will open notepad with the file right here
        Shell ("c:\windows\notepad.exe " & CommonDialog1.FileName), vbNormalFocus
 
             'Assign the chosen file to the path string.
             MstrFilePath = CommonDialog1.FileName
             
     End Select


End Sub

Private Sub cmdOpenOLExpress_Click()
'First clear the clipboard
  Clipboard.Clear
  
    Shell "C:\Program Files\Outlook Express\msimn.exe", 1
    
End Sub

Private Sub cmdOpenMAPI_Click()
FileToAttach = txtRFName
 
'Select the program you are registering
Select Case cboProgram.ListIndex
    Case 0
        PrgmName = "Program 1"
    Case 1
        PrgmName = "Program 2"
    Case 2
        PrgmName = "Program 3"
    Case 3
        PrgmName = "Program 4"
 End Select
 
'**********************************************************************
'Add the MAPI components (MSMAPI32.OCX)

MAPISession1.SignOn
MAPISession1.DownLoadMail = False

DoEvents
    MAPIMessage1.SessionID = MAPISession1.SessionID
    MAPIMessage1.Compose
    
    MAPIMessage1.RecipAddress = txtEmailAddress
    MAPIMessage1.ResolveName
    MAPIMessage1.MsgSubject = "Registration File Instructions"
    MAPIMessage1.AttachmentPathName = FileToAttach
    MAPIMessage1.AttachmentName = txtName & ".txt"

'*******Supply your name and email address in the last line of this Select Stmt************
    MAPIMessage1.MsgNoteText = "Attached File:  - " & txtName & ".txt" & "    - " & PrgmName _
        & vbCrLf & vbCrLf _
        & "This file was intentionally sent as a ''.txt'' File  in order to preclude your email server from treating the attached file as if it were infected with a virus. " _
        & "(Some Email Servers will not allow files with a ''*.reg'' file extension to be opened.)" _
        & vbCrLf & vbCrLf & "You should copy the attached file to your root (C:\) directory, then rename it with a reg extension," _
        & vbCrLf & vbCrLf & "    i.e.,                ''" & txtName & ".reg" & "''" _
        & vbCrLf & vbCrLf & "Next open Windows Explorer, navigate to the file (C:\" & txtName & ".reg) and click on the file to add its contents to the registry." _
        & vbCrLf & vbCrLf & vbCrLf & "Regards," & vbCrLf & vbCrLf _
                        & "YourName" & vbCrLf & vbCrLf & "youremail@any.com"
'*******Supply your name  & email address in the above line************

'**********************************************************************
MAPIMessage1.Send False
'**********************************************************************
MAPISession1.SignOff
'**********************************************************************
End Sub

Public Sub AddNewRecord()

Close_cn
'     Set the Database Applicable Path
      DbFile = App.Path & "\PrgmRegistration.mdb"

   ' Establish the Connection
       Set cn = New ADODB.Connection
       cn.CursorLocation = adUseClient
       cn.ConnectionString = _
             "Provider=Microsoft.Jet.OLEDB.4.0;" & _
             "Data Source=" & DbFile & ";" & _
             "Persist Security Info=False"

      'Open the Connection
      cn.Open
'*********************************************************************
'If you add to the combobox you will have to edit the database table names
Select Case cboProgram.ListIndex
    Case 0
        SQLstmt = "SELECT * FROM [tblProgram1Reg]Order by ID"
    Case 1
        SQLstmt = "SELECT * FROM [tblProgram2Reg]Order by ID"
    Case 2
         'Add your SQL statement here for Program 3
    Case 3
         'Add your SQL statement here for Program 4
End Select
'*********************************************************************
   Set DataGrid1.DataSource = rs
   
   Set rs = New ADODB.Recordset
   rs.Open SQLstmt, cn, adOpenStatic, adLockOptimistic, adCmdText
           
'How many records in the Database
     RcCount = rs.RecordCount
  
     Dim i As Integer

   ' Get data from the TextBoxes.
    intID = RcCount + 1
    strName = txtName
    strEmailAddress = txtEmailAddress
    strRegistration = txtKey
    StrTDate = Date

      rs.AddNew
      rs!ID = intID
      rs!Name = strName
      rs!EmailAddress = strEmailAddress
      rs!RegistrationNumber = strRegistration
      rs!RegistrationDate = StrTDate
       
      
      rs.Update
      Set DataGrid1.DataSource = rs
   Close_cn
    
End Sub

Private Sub cmdShowDataBase_Click()
    'The following demonstrates the ShellExecute API method.
    ShellExecute Me.hwnd, "Open", App.Path & _
        "\PrgmRegistration.mdb", "", "C:\", vbNormalFocus

End Sub

Private Sub cmdShowRegistry_Click()

    Shell "c:\windows\regedit.exe ", vbNormalFocus

End Sub

Private Sub cmdRunRegFile_Click()

        ShellExecute Me.hwnd, "Open", _
        "C:\Build Reg n SN Files\" & txtName & ".reg", "", "C:\", vbNormalFocus
   
End Sub

Private Sub cmdKillAllRegnTextFiles_Click()

 Kill "C:\Build Reg n SN Files\" & "*.*"
   
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 'Allow the <Enter> Key to  be used as a <Tab> Key
     'Set the Form's KeyPreview to True.

     If KeyAscii = vbKeyReturn Then
          KeyAscii = 0
          SendKeys "{TAB}"
     End If
     
End Sub

Private Sub Form_Load()

    'The following allows the <F1> key to execute the Help File
    'NB! Remember to set the Form's HelpContextID to 0
    App.HelpFile = App.Path & AppFile
    
    Me.Height = 6300
   
    cboProgram.AddItem "    Program 1"
    cboProgram.AddItem "    Program 2"
    cboProgram.AddItem "    Program 3"
    cboProgram.AddItem "    Program 4"
    

 
End Sub

Function ParseStr(ByVal Text, ByVal separator, ByVal start As Integer, _
ByVal toEnd As Integer) As String
    
    Dim i As Integer, Temp As String, result As String
    Dim ParseStrBegin As Integer, t As Integer, Count As Integer
    Dim ParseStrEnd As Integer, Found As Integer
    
    ParseStr = ""
    If Text = "" Then Exit Function
    If separator = "" Then Exit Function
    If Not (start > 0) Then start = 1
    If toEnd < start Then toEnd = start
    'Find first instance of the separator
     t = InStr(1, Text, separator)
    
    'If no occurence return original string and exit
    If t = 0 Then
        ParseStr = Text
        Exit Function
    End If

    'If first ParseStr, return left most data and exit

    If (start = 1) And (start = toEnd) Then
        If t = 1 Then
            ParseStr = ""
            Exit Function
        Else
            ParseStr = Left(Text, t - 1)
            Exit Function
        End If
    End If
    
    ParseStrBegin = 1
    For i = 1 To start - 1
       t = InStr(ParseStrBegin, Text, separator)
       If t = 0 Then Exit For
       ParseStrBegin = t + 1
    Next i
    
    ' If there is no separator exit function with "" result
    If t = 0 Then Exit Function
    
    'If only one ParseStr to return, find it and exit
    If start = toEnd Then
        t = InStr(ParseStrBegin, Text, separator)
    If t = 0 Then t = Len(Text) + 1
    result = Left(Text, t - 1)
    ParseStr = Right(result, t - ParseStrBegin)
    Exit Function
    End If
    
    'Find last ParseStr then exit

    ParseStrEnd = t + 1
    If start = 1 Then start = 2
    For i = start To toEnd

    t = InStr(ParseStrEnd, Text, separator)
    If t = 0 Then
        t = Len(Text) + 1
        Exit For
    End If
    
    ParseStrEnd = t + 1
    Next i
    
    If t = 0 Then t = Len(Text) + 1
    result = Left(Text, t - 1)
    ParseStr = Right(result, t - ParseStrBegin)

End Function

Private Sub Close_cn()
     
     Set cn = Nothing
     Set rs = Nothing
     
End Sub

Private Sub cmdDelete_Click()
Dim intResponse As Integer

    'Remember to add the cmdDelete Command Button to your Form
    Beep
    intResponse = MsgBox("Delete the Current Record", _
                  vbYesNo + vbQuestion, "Delete Record")

'****************************************************

     If intResponse = vbYes Then
     
          If Not rs.EditMode = adEditAdd Then
               rs.Delete
          End If
          
     End If
     rs.MoveFirst
     
        With rs
            txtName.Text = !Name
            txtEmailAddress = !EmailAddress
            txtKey = !RegistrationNumber
            Set DataGrid1.DataSource = rs
        End With
  
'****************************************************

 'txtName = vbNullString
 'txtEmailAddress = vbNullString
 'txtKey = vbNullString
    
End Sub

Private Sub DataGrid1_Click()
'When user clicks on a DataGrid Item, show the values in the TextBoxes

'*********************************************************************

        With rs
            txtName.Text = !Name
            txtEmailAddress = !EmailAddress
            txtKey = !RegistrationNumber
            Set DataGrid1.DataSource = rs
        End With
   
'*********************************************************************
cmdRunRegFile.Enabled = False

End Sub



Private Sub mnuAbout_Click()

    frmAbout.Show

End Sub

Private Sub mnuExit_Click()

    Unload Me
    
End Sub

Private Sub mnuHelp_Click()

    ExecPath = IIf("/" = Mid(App.Path, Len(App.Path)), App.Path, App.Path) & AppFile
    LaunchApp (App.Path & AppFile)

End Sub

Private Sub txtEmailAddress_GotFocus()

    txtEmailAddress.SelStart = 0
    txtEmailAddress.SelLength = Len(txtEmailAddress)
    
End Sub

Private Sub txtKey_GotFocus()

    txtKey.SelStart = 0
    txtKey.SelLength = Len(txtKey)
    
End Sub

Private Sub txtName_GotFocus()

    txtName.SelStart = 0
    txtName.SelLength = Len(txtName)
    
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

If KeyAscii > 95 And KeyAscii < 123 Then
'Capitalize the first character
        If txtName.SelStart = 0 Then
            KeyAscii = KeyAscii - 32
        'if more than one word capitalize
        ElseIf Mid(txtName.Text, txtName.SelStart, 1) < "!" Then
            KeyAscii = KeyAscii - 32
         End If
    End If
    
End Sub
Sub CreateNewDirectory(DirName As String)

    Dim NewLen As Integer
    Dim DirLen As Integer
    Dim MaxLen As Integer
    
    NewLen = 4
    MaxLen = Len(DirName)
    
    If Right$(DirName, 1) <> "\" Then
        DirName = DirName + "\"
        MaxLen = MaxLen + 1
    End If
    
    On Error GoTo DirError
    
MakeNext:
    DirLen = InStr(NewLen, DirName, "\")
    MkDir Left$(DirName, DirLen - 1)
    NewLen = DirLen + 1
    If NewLen >= MaxLen Then
        Exit Sub
    End If
    GoTo MakeNext

DirError:
    Resume Next
End Sub

Public Function LaunchApp(ByVal URL As String) As Long
Dim strFile As String
    
    On Error Resume Next
    strFile = ShellExecute(0&, vbNullString, URL, vbNullString, _
    vbNullString, vbNormalFocus)
    
End Function


