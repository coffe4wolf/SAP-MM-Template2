VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MultiSelectionForm 
   ClientHeight    =   6570
   ClientLeft      =   -315
   ClientTop       =   -1110
   ClientWidth     =   5655
   OleObjectBlob   =   "MultiSelectionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MultiSelectionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Public Sub ClearButtonActive_ImageButton_Click()
    Call Populate2DimListBox(wsContractorsMaster.Name, Wb) ' Populate with values
End Sub


Public Sub FindButtonActive_ImageButton_Click()

    Dim sht As Worksheet
    Dim str As String
    Dim i As Long, LR As Long
    
    Set sht = wsContractorsMaster
    str = FindBox.value
    LR = GetBorders("LR", sht.Name, ThisWorkbook)
    
    If str <> vbNullString Then
    
        ' Clear all items except selected
        For i = ListBox1.ListCount - 1 To 0 Step -1
            If Not ListBox1.Selected(i) Then ListBox1.RemoveItem i
        Next i
        
        ' Perform search
        For i = 2 To LR
            If InStr(1, sht.Cells(i, 1).value, str, vbTextCompare) > 0 Or InStr(1, sht.Cells(i, 2).value, str, vbTextCompare) > 0 Then
                ListBox1.AddItem
                ListBox1.List(ListBox1.ListCount - 1, 0) = sht.Cells(i, 1).value
                ListBox1.List(ListBox1.ListCount - 1, 1) = sht.Cells(i, 2).value
            End If
        Next i
    Else
        Call Populate2DimListBox(sht.Name) ' Populate with values
    End If

End Sub

Private Sub GoButtonActive_ImageButton_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Public Sub GoButtonActive_ImageButton_Click()

Dim i As Long, J As Long, Count As Long

J = 0
Count = 0

For i = 0 To ListBox1.ListCount - 1
    'Check if the row is selected and add to count
    If ListBox1.Selected(i) Then
        wsCreateMM.chooseSupplier_TextBox.value = ListBox1.List(i)
    End If
Next i

'ProductCode = ListBox1.List(0)
'RequestForm_UserForm.ProductCode_TextBox.Value = ""
'RequestForm_UserForm.ProductCode_TextBox.Value = ProductCode

Unload Me

Exit Sub

End Sub


Private Sub GoButtonInactive_ImageButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    GoButtonInactive_ImageButton.Visible = False
    
End Sub

Private Sub FindButtonInactive_ImageButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    FindButtonInactive_ImageButton.Visible = False
    ClearButtonInactive_ImageButton.Visible = True
    
End Sub


Private Sub ClearButtonInactive_ImageButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    ClearButtonInactive_ImageButton.Visible = False
    FindButtonInactive_ImageButton.Visible = True
    
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    GoButtonInactive_ImageButton.Visible = True
    FindButtonInactive_ImageButton.Visible = True
    ClearButtonInactive_ImageButton.Visible = True
End Sub

Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    Unload MultiSelectionForm

End Sub

