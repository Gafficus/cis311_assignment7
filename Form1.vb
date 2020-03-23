'------------------------------------------------------------
'-                File Name : Form1.frm                     - 
'-                Part of Project: Assign7                  -
'------------------------------------------------------------
'-                Written By: Nathan Gaffney                -
'-                Written On: 22 Mar 2020                   -
'------------------------------------------------------------
'- File Purpose:                                            -
'- This file contains the main application form where the   -
'- user will input use drag and drop functionality to play  -
'- a modified version of connect four. The wind condition   -
'- will be constricted to vertical or horizantal series of  -
'- four peices.                                             -
'------------------------------------------------------------
'- Program Purpose:                                         -
'-                                                          -
'- This program allows for user/s to play connect four      -
'- locally                                                  -
'------------------------------------------------------------
'- Global Variable Dictionary (alphabetically):             -
'- (None)                                                   -
'------------------------------------------------------------
Public Class Form1
    '------------------------------------------------------------
    '-                Subprogram Name: Form1_Load               -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 22 Mar 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '- This subroutine is called when the form loads. It will   -
    '- enable drag and drop on the column headers               -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim columns = From buttons In Me.Controls
                      Where buttons.GetType = GetType(Button) And buttons.Name.StartsWith("btnCol")
                      Select buttons
        For Each c In columns
            c.AllowDrop = True
        Next
    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: dragPlayerButtons        -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 22 Mar 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '- This subroutine is called-
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub DragPlayerButtons(sender As Object, e As MouseEventArgs) Handles btnPlayer1.MouseMove,
                                                                                 btnPlayer2.MouseMove
        If e.Button = 0 Then
            Exit Sub
        End If
        Dim DragDropResult As DragDropEffects
        Dim data As New DataObject()
        data.SetData(sender)
        DragDropResult = sender.DoDragDrop(data, DragDropEffects.Copy)
    End Sub

    Private Sub ColumnsDragEnter(sender As Object, e As DragEventArgs) _
                                            Handles btnCol1.DragEnter,
                                                    btnCol2.DragEnter,
                                                    btnCol3.DragEnter,
                                                    btnCol4.DragEnter,
                                                    btnCol5.DragEnter,
                                                    btnCol6.DragEnter,
                                                    btnCol7.DragEnter
        If ColumnFull(sender) Then
            sender.AllowDrop = False
            sender.BackColor = Color.Red
        Else
            sender.BackColor = Color.Green
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub ColumnsDragLeave(sender As Object, e As EventArgs) _
                                            Handles btnCol1.DragLeave,
                                                    btnCol2.DragLeave,
                                                    btnCol3.DragLeave,
                                                    btnCol4.DragLeave,
                                                    btnCol5.DragLeave,
                                                    btnCol6.DragLeave,
                                                    btnCol7.DragLeave

        sender.BackColor = btnReset.BackColor
        sender.UseVisualStyleBackColor = True

    End Sub
    Private Function ColumnFull(sender As Object) As Boolean
        Return False
    End Function

    Private Sub ColumnsDragDrop(sender As Object, e As DragEventArgs) _
                                            Handles btnCol1.DragDrop,
                                                    btnCol2.DragDrop,
                                                    btnCol3.DragDrop,
                                                    btnCol4.DragDrop,
                                                    btnCol5.DragDrop,
                                                    btnCol6.DragDrop,
                                                    btnCol7.DragDrop
        Dim panel = From controls In Me.Controls
                    Where controls.Name = sender.Tag
                    Select controls
        For Each aButton As Control In panel(0).Controls
            MessageBox.Show(aButton.Name)
        Next
    End Sub
End Class
