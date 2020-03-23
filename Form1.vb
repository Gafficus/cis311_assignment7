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
    Dim arrBoard(5, 6) As Integer
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
        If MessageBox.Show("Does Blue Player play first?", "Player Order", MessageBoxButtons.YesNo) _
                        = DialogResult.Yes Then
            btnPlayer1.Enabled = True
            btnPlayer2.Enabled = False
        Else
            btnPlayer1.Enabled = False
            btnPlayer2.Enabled = True
        End If
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
    '------------------------------------------------------------
    '-                Subprogram Name: ColumnsDragEnter         -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 23 Mar 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '- This subroutine is called when the mouse drags into the  -
    '- control, it will change the color of the button          -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e As DragEventArgs - holds the event args                -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
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
    '------------------------------------------------------------
    '-                Subprogram Name: ColumnsDragLeave         -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 23 Mar 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '- This subroutine is called when the mouse drags out of the-
    '- control, it will reset the color of the button           -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e As EventArgs - holds the event args                    -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
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
    '------------------------------------------------------------
    '-                Subprogram Name: ColumnFull           -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 23 Mar 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called to determine if a column is full-
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- flag - holds whether the columns are considered full     -
    '- intNumberFilled - the number of rows that have pieces    -
    '- panel - the column that contains the pieces              -
    '------------------------------------------------------------
    '- Returns:                                                 -
    '- flag – telling if the column is full                     -
    '------------------------------------------------------------
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
    '------------------------------------------------------------
    '-                Subprogram Name: ColumnsDragDrop          -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 23 Mar 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when the reset button is       -
    '- pressed, it will iterate through the board and reset the -
    '- board to a fresh state                                   -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the DragEventArgs object sent to the routine   -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- intControl - this is the index iterator for which control-
    '- will have its' contents changed                          -
    '- lastButton - this is the last button that waas actually  -
    '- changed, this allows the program to know which buttons   -
    '- have not been changed                                    -
    '- panel - this is the column that the piece is being       -
    '- dropped into                                             -
    '- parentButton - this is the player that is dropping the   -
    '- piece, it will have the coloring and o indicator         -
    '------------------------------------------------------------
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
        Dim intPlayerNumber As Integer = 1
        If btnPlayer1.Enabled = False Then
            intPlayerNumber = 2
        End If
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
        arrBoard(CInt(lastButton.Tag.Substring(0, 1)), CInt(lastButton.Tag.Substring(1, 1))) =
            intPlayerNumber
        If checkForWin(lastButton) Then
            Dim strWinner As String
            If btnPlayer1.Enabled = True Then
                strWinner = "One"
            Else
                strWinner = "Two"
            End If
            lblWin.Text = "Player " & strWinner & " wins!"
        End If
        swapPlayers()
    End Sub
    '------------------------------------------------------------
    '-                Subprogram Name: swapPlayers           -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 23 Mar 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when the dragdrop completes    -
    '- this will swap which player is available to play         -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- (None)                                                  --
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- (None)                                                   -
    '------------------------------------------------------------
    Private Sub swapPlayers()
        btnPlayer1.Enabled = Not btnPlayer1.Enabled
        btnPlayer2.Enabled = Not btnPlayer2.Enabled
    End Sub

    Private Function checkForWin(placedButton As Button) As Boolean
        'Dim right As Integer = rightCheck(placedButton)
        'Dim left As Integer = leftCheck(placedButton)
        'Dim up As Integer = upCheck(placedButton)
        'Dim down As Integer = downCheck(placedButton)
        'If (right + left) >= 4 Or (up + down) >= 4 Then
        '    Return True
        'Else
        '    Return False
        'End If
        Return False
    End Function

    'Private Function upCheck(location As Button) As Integer
    '    Dim numCont As Integer = 0
    '    Dim intLocationY As Integer = CInt(location.Tag.Substring(0, 1))
    '    Dim intLocationX As Integer = CInt(location.Tag.Substring(1, 1))

    '    Return numCont
    '    Throw New NotImplementedException
    'End Function

    'Private Function downCheck(location As Button) As Integer
    '    Throw New NotImplementedException
    'End Function

    'Private Function leftCheck(location As Button) As Integer
    '    Throw New NotImplementedException
    'End Function

    'Private Function rightCheck(location As Button) As Integer
    '    Throw New NotImplementedException
    'End Function
    '------------------------------------------------------------
    '-                Subprogram Name: btnReset_Click           -
    '------------------------------------------------------------
    '-                Written By: Nathan Gaffney                -
    '-                Written On: 23 Mar 2020                   -
    '------------------------------------------------------------
    '- Subprogram Purpose:                                      -
    '-                                                          -
    '- This subroutine is called when the reset button is       -
    '- pressed, it will iterate through the board and reset the -
    '- board to a fresh state                                   -
    '------------------------------------------------------------
    '- Parameter Dictionary (in parameter order):               -
    '- sender – Identifies which particular control raised the  –
    '-          click event                                     - 
    '- e – Holds the EventArgs object sent to the routine       -
    '------------------------------------------------------------
    '- Local Variable Dictionary (alphabetically):              -
    '- panels - this holds an list of all the panels that create-
    '- the board.                                               -
    '------------------------------------------------------------
    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        'restricting to starts with is redundant but prevents accidents
        Dim panels = From panel In Me.Controls
                     Where panel.GetType = GetType(Panel) And panel.Name.startsWith("pnl")
                     Select panel
        For Each panel In panels
            'Iterate trhough all the buttons contained in the 
            'panel and then change it's color to gray
            For Each button In panel.Controls
                button.BackColor = btnReset.BackColor
                button.UseVisualStyleBackColor = True
                button.Text = ""
            Next
        Next
    End Sub
End Class
