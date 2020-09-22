VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "EIT - Virus Removal"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6990
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Get Updates"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "About"
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   6735
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Scanning"
         Height          =   2775
         Left            =   0
         TabIndex        =   2
         Top             =   1800
         Width           =   6735
         Begin VB.CommandButton cmdInfo 
            Caption         =   "&Get Info On Selected"
            Height          =   375
            Left            =   3600
            TabIndex        =   6
            Top             =   2280
            Width           =   1935
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete Selected"
            Height          =   375
            Left            =   1920
            TabIndex        =   5
            Top             =   2280
            Width           =   1575
         End
         Begin VB.CommandButton cmdScan 
            Caption         =   "&Scan For Selected"
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   2280
            Width           =   1695
         End
         Begin VB.ListBox lstVirus 
            BackColor       =   &H00E0E0E0&
            Height          =   1815
            ItemData        =   "frmMain.frx":0442
            Left            =   120
            List            =   "frmMain.frx":0444
            TabIndex        =   3
            Top             =   360
            Width           =   6495
         End
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"frmMain.frx":0446
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   6495
      End
   End
   Begin VB.Image Image1 
      Height          =   1170
      Left            =   0
      Picture         =   "frmMain.frx":065E
      Top             =   0
      Width           =   7020
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
MsgBox "Not yet implemented. Just scan for the virus, if it's found then you can delete it.", vbInformation, "Nope"
End Sub

Private Sub cmdInfo_Click()
MsgBox "Not yet implemented.", vbInformation, "Nope"
End Sub

Private Sub cmdScan_Click()
Dim Delete As String
Delete = ""

'This may look a little messy, but it's all I know how to do..
'If you can approve upon this, then email me at kold@localnet.com with your suggestions and/or new code
'and I will (maybe) put it in the next release.
Select Case lstVirus.Text
    Case "W32/Sobig.f@MM - A High Risk Worm"
        If Dir("C:\WINNT\WINPPR32.exe") > "C:\WINNT\WINPPR32.exe" Then
            Delete = MsgBox("Warning: This virus has been found on your system. Would you like to delete it now?", vbYesNo + vbCritical, "Virus Alert")
            If Delete = vbYes Then
                Kill "C:\WINNT\WINPPR32.exe"
                SaveSetting "EIT - Virus Removal", "Settings", "Sobig", "1"
            End If
        Else
            Delete = MsgBox("This virus was not found on your system. Would you like to remove it from the list?", vbYesNo + vbInformation, "Virus Not Found")
            If Delete = vbYes Then
                SaveSetting "EIT - Virus Removal", "Settings", "Sobig", "1"
            End If
        End If
        
    Case "W32/Dumaru.a@MM - A Medium Risk Worm"
        If Dir("C:\WINNT\dllreg.exe") > "C:\WINNT\dllreg.exe" Then
            Delete = MsgBox("Warning: This virus has been found on your system. Would you like to delete it now?", vbYesNo + vbCritical, "Virus Alert")
            If Delete = vbYes Then
                Kill "C:\WINNT\dllreg.exe"
                Kill "C:\WINNT\SYSTEM\load32.exe"
                Kill "C:\WINNT\SYSTEM\vxdmgr32.exe"
                SaveSetting "EIT - Virus Removal", "Settings", "Dumaru", "1"
            End If
        Else
            Delete = MsgBox("This virus was not found on your system. Would you like to remove it from the list?", vbYesNo + vbInformation, "Virus Not Found")
            If Delete = vbYes Then
                SaveSetting "EIT - Virus Removal", "Settings", "Dumaru", "1"
            End If
        End If
        
    Case "W32/Lovsan.worm.a - A Medium Risk Worm"
        If Dir("C:\Windows\System32\msblast.exe") > "C:\Windows\System32\msblast.exe" Then
            Delete = MsgBox("Warning: This virus has been found on your system. Would you like to delete it now?", vbYesNo + vbCritical, "Virus Alert")
            If Delete = vbYes Then
                Kill "C:\Windows\System32\msblast.exe"
                SaveSetting "EIT - Virus Removal", "Settings", "Lovsan", "1"
            End If
        Else
            Delete = MsgBox("This virus was not found on your system. Would you like to remove it from the list?", vbYesNo + vbInformation, "Virus Not Found")
            If Delete = vbYes Then
                SaveSetting "EIT - Virus Removal", "Settings", "Lovsan", "1"
            End If
        End If
End Select
End Sub

Private Sub Command1_Click()
Unload Me
End
End Sub

Private Sub Command2_Click()
MsgBox "Not yet implemented. Check on www.pscode.com for any updates to this program.", vbInformation, "Nope"
End Sub

Private Sub Form_Load()
LoadViruses
End Sub

Public Sub LoadViruses()
Dim Sobig As String
Dim Dumaru As String
Dim Lovsan As String

SaveSetting "EIT - Virus Removal", "Settings", "Sobig", "0"
SaveSetting "EIT - Virus Removal", "Settings", "Dumaru", "0"
SaveSetting "EIT - Virus Removal", "Settings", "Lovsan", "0"

'If the virus setting is 0 then it hasn't been scanned for yet
'Once the virus is deleted or not found in the scan, its setting is set to 1 and it is removed from the virus list
Sobig = GetSetting("EIT - Virus Removal", "Settings", "Sobig", "0")
Dumaru = GetSetting("EIT - Virus Removal", "Settings", "Dumaru", "0")
Lovsan = GetSetting("EIT - Virus Removal", "Settings", "Lovsan", "0")

If Sobig = "0" Then lstVirus.AddItem "W32/Sobig.f@MM - A High Risk Worm"
If Dumaru = "0" Then lstVirus.AddItem "W32/Dumaru.a@MM - A Medium Risk Worm"
If Lovsan = "0" Then lstVirus.AddItem "W32/Lovsan.worm.a - A Medium Risk Worm"

'If they've been removed already, tell the user that they have been
If Sobig And Dumaru And Lovsan = "1" Then lstVirus.AddItem "You have removed all available viruses to scan for."
End Sub
