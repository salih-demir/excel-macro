VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Macro"
   ClientHeight    =   1452
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   1884
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim b As Boolean
Private Sub CommandButton1_Click()
Dim a As Boolean
a = False
  Rows.RowHeight = 50
    Rows.ColumnWidth = 10
    Cells.Orientation = 90
    Cells.Font.Size = 18
    Cells.Font.Name = "Buxton Sketch"
    Cells.HorizontalAlignment = xlCenter
    Cells.VerticalAlignment = xlCenter
    If b = False Then
    For k = 1 To 6
    For i = 1 To 6
        If k Mod 2 = 0 Then
         If a = False Then
        Cells(i, k).Interior.ColorIndex = 2
        Cells(i, k).Font.ColorIndex = 1
        Cells(i, k).Value = "):"
        a = True
        Else
        Cells(i, k).Interior.ColorIndex = 1
        Cells(i, k).Font.ColorIndex = 2
        Cells(i, k).Value = "(:"
        a = False
        End If
        Else
         If a = False Then
        Cells(i, k).Interior.ColorIndex = 1
        Cells(i, k).Font.ColorIndex = 2
        Cells(i, k).Value = "(:"
        a = True
        Else
        Cells(i, k).Interior.ColorIndex = 2
        Cells(i, k).Font.ColorIndex = 1
        Cells(i, k).Value = "):"
        a = False
        End If
        End If
    Next i
Next k
 b = True
    Else
    b = False
       For k = 1 To 6
    For i = 1 To 6
        If k Mod 2 = 0 Then
         If a = False Then
          Cells(i, k).Interior.ColorIndex = 1
        Cells(i, k).Font.ColorIndex = 2
        Cells(i, k).Value = "(:"
        a = True
        Else
        Cells(i, k).Interior.ColorIndex = 2
        Cells(i, k).Font.ColorIndex = 1
        Cells(i, k).Value = "(:"
        a = False
        End If
        Else
         If a = False Then
        Cells(i, k).Interior.ColorIndex = 2
        Cells(i, k).Font.ColorIndex = 1
        Cells(i, k).Value = "(:"
        a = True
        Else
          Cells(i, k).Interior.ColorIndex = 1
        Cells(i, k).Font.ColorIndex = 2
        Cells(i, k).Value = "(:"
        a = False
        End If
        End If
    Next i
Next k
    End If
End Sub
Private Sub CommandButton2_Click()
MsgBox "Hello Office"
End Sub
