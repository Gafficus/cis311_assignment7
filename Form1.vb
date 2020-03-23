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
        data.SetData(sender.Name)
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
        Dim flag As Boolean = True
        Dim intNumberFilled As Integer = 0
        Dim panel = From controls In Me.Controls
                    Where controls.Name = sender.Tag
                    Select controls
        For Each aButton As Control In panel(0).Controls
            If Not aButton.Text = "" Then
                intNumberFilled += 1
            End If
        Next
        If intNumberFilled < panel(0).Controls.Count Then
            flag = False
        End If
        Return flag
    End Function

    Private Sub ColumnsDragDrop(sender As Object, e As DragEventArgs) _
                                            Handles btnCol1.DragDrop,
                                                    btnCol2.DragDrop,
                                                    btnCol3.DragDrop,
                                                    btnCol4.DragDrop,
                                                    btnCol5.DragDrop,
                                                    btnCol6.DragDrop,
                                                    btnCol7.DragDrop
        If ColumnFull(sender) Then
            Exit Sub
        End If

        sender.BackColor = btnReset.BackColor
        sender.UseVisualStyleBackColor = True
        Dim intControl As Integer = 5
        Dim panel = From controls In Me.Controls
                    Where controls.Name = sender.Tag
                    Select controls

        Dim parentButton = From controls In Me.Controls
                           Where controls.Name = e.Data.GetData("Text")
                           Select controls
        Dim lastButton As Button
        With panel(0)

            While intControl >= 0
                If .Controls(intControl).Text = "" Then
                    .Controls(intControl).BackColor = parentButton(0).BackColor
                    .Controls(intControl).Text = parentButton(0).Text
                    System.Threading.Thread.Sleep(25)
                    Application.DoEvents()
                    lastButton = .Controls(intControl)
                    .Controls(intControl).BackColor = btnReset.BackColor
                    .Controls(intControl).UseVisualStyleBackColor = True
                    .Controls(intControl).Text = ""
                End If
                intControl -= 1
            End While
        End With
        lastButton.BackColor = parentButton(0).BackColor
        lastButton.Text = parentButton(0).Text
        If checkForWin(lastButton) Then

        End If
    End Sub

    Private Function checkForWin(lastButton As Button) As Boolean
        Return False
    End Function

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Dim panels = From panel In Me.Controls
                     Where panel.GetType = GetType(Panel) And panel.Name.startsWith("pnl")
                     Select panel
        For Each panel In panels
            For Each button In panel.Controls
                button.BackColor = btnReset.BackColor
                button.UseVisualStyleBackColor = True
                button.Text = ""
            Next
        Next
    End Sub
End Class
