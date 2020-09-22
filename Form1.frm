VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "TabStrip"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   7
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Item 1"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Item 2"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Item 3"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Item 4"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Item 5"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Item 6"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Item 7"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4200
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0CDC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Ubicación"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Item"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'* Autor : Carlos Olaya Boulangger *'

Dim nPosicion As Byte
Dim Cambio As Boolean
Dim Ultimo As Byte

Private Sub TabStrip1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Integer, i As Integer

    Text2 = x
    nCol = 0
    For i = 1 To TabStrip1.Tabs.Count
        nCol = nCol + TabStrip1.Tabs(i).Width
        If x > TabStrip1.Tabs(i).Left - (TabStrip1.Left + 30) And x < nCol Then
            Text1 = i
            Ultimo = i
            If Ultimo <> nPosicion Then
                Ultimo = nPosicion
                nPosicion = i
                Cambio = True
            End If
        End If
    Next
    If Cambio Then
        For i = 1 To TabStrip1.Tabs.Count
            If TabStrip1.Tabs(i).Image <> 1 And i <> nPosicion Then
                TabStrip1.Tabs(i).Image = 1
            End If
        Next
        If nPosicion <> 0 Then
            If TabStrip1.Tabs(nPosicion).Image <> 2 Then
                TabStrip1.Tabs(nPosicion).Image = 2
            End If
        End If
        Cambio = False
    End If
End Sub



