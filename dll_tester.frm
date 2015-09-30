VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   LinkTopic       =   "Sistema de testes de component"
   ScaleHeight     =   1875
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Execute 
      Caption         =   "Iniciar teste"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Execute_Click()
    
    Dim obj As Object
    Dim rs As Object
    Dim Parameter1, Parameter2
    Dim erro, dll As String
    
    dll = "MinhaDLL.cls_class_module" ' ex: MinhaDLL.dll

    Set obj = CreateObject(dll)

    ' This is for a string function
    Debug.Print (obj.GetValue(Parameter1))

    ' This example uses a RecordSet function
    Set rs = obj.Function(Parameter1, Parameter2, erro) 'the erro is a string
                                                        'for returning a value for returning a internal error
'
    If Len(erro) > 0 Then
        Debug.Print erro
    Else
        Debug.Print (rs.Fields(0).Name)
        Debug.Print (rs.Fields(0).Value)
        Debug.Print (rs.Fields(1).Name)
        Debug.Print (rs.Fields(1).Value)
        Debug.Print (rs.Fields(2).Name)
        Debug.Print (rs.Fields(2).Value)
        ' You can add as many fields your RecordSet have
                
        Debug.Print (rs.RecordCount)
    End If

End Sub
