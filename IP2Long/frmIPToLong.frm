VERSION 5.00
Begin VB.Form frmIPToLong 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dotted IP To Long"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   3420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton OptUseCustom 
      Caption         =   "Use Custom Function"
      Height          =   225
      Left            =   105
      TabIndex        =   3
      Top             =   525
      Value           =   -1  'True
      Width           =   1905
   End
   Begin VB.OptionButton optUseAPI 
      Caption         =   "Use API function"
      Height          =   225
      Left            =   105
      TabIndex        =   2
      Top             =   840
      Width           =   1905
   End
   Begin VB.TextBox txtIP 
      Height          =   330
      Left            =   105
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   105
      Width           =   3270
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   540
      Left            =   2100
      TabIndex        =   0
      Top             =   525
      Width           =   1275
   End
End
Attribute VB_Name = "frmIPToLong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long


Private Sub cmdConvert_Click()
  
  Dim LongIP As Long
    
    'We can use the APi function
    If optUseAPI.Value Then
        LongIP = inet_addr(txtIP.Text)
    'Or our custom built function
    Else
        LongIP = modIPToLong.IPToLong(txtIP.Text)
    End If
    
    'Was it an in valid IP?
    If LongIP = modIPToLong.ERR_INVALIDIP Then
        txtIP.Text = "*** Error ***"
    'It was valid
    Else
        txtIP.Text = Str(LongIP)
    End If
    
End Sub
