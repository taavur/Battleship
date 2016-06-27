' FAILINIMI:    formBattleshipMain.vb
' AUTOR:        Taavi Kappak
' LOODUD:       5.11.2014 
' MUUDETUD:     1.12.2014 
' 
' KIRJELDUS:    Programm, mis realiseerib laevade pommitamise mängu kahe mängija jaoks.

Public Class formBattleshipMain

    ' Enumeratioon on koodi loetavuse parandamiseks, et ei tekiks mingeid maagilisi numbreid.
    Private Enum Ships
        aircraftCarrier = 1
        Battleship = 2
        Submarine = 3
        Destroyer = 4
        patrolBoat = 5
    End Enum

    Private Enum direction
        horizontal = 0
        vertical = 1
    End Enum

    ' Muutujad
    Dim playerOneHits As Integer = 0
    Dim playerTwoHits As Integer = 0
    Dim playerOneTurn As Boolean = True
    Dim gameStarted As Boolean = False
    Dim fieldGenerated As Boolean = False

    ' Programmi käivitades tehtavad operatsioonid ehk peamiselt asjade peitmine
    Private Sub formBattleshipMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblComputer.Visible = False
        lblPlayer.Visible = False
        tlpComputer.Visible = False
        tlpPlayer.Visible = False
        btnGenerateFields.Visible = False
        btnStart.Visible = False
        lblplayerone.Visible = False
        lblplayertwo.Visible = False
        lblScore.Visible = False
        lblPlayerOneScore.Visible = False
        lblPlayerTwoScore.Visible = False
    End Sub

    ' Exit nuppu vajutades sulgub programm
    Private Sub btnMainExit_Click(sender As Object, e As EventArgs) Handles btnMainExit.Click
        Me.Close()
    End Sub

    ' Klikkides menüüst Exit, sulgub programm
    Private Sub ExitToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    ' Menüüst saab lahti võtta abi akna, mis seletab, kuidas mängida
    Private Sub HowToPlayToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HowToPlayToolStripMenuItem.Click
        formHowToPlay.Show()
    End Sub

    ' 'Enter game' nuppu vajutades kaob see nupp ja ilmuvad mänguväljad ning lisanupud
    Private Sub btnMainStart_Click(sender As Object, e As EventArgs) Handles btnMainStart.Click
        btnMainStart.Visible = False
        lblComputer.Visible = True
        lblPlayer.Visible = True
        tlpComputer.Visible = True
        tlpPlayer.Visible = True
        btnGenerateFields.Visible = True
        btnStart.Visible = True
        lblplayerone.Visible = True
        lblplayertwo.Visible = True
        lblScore.Visible = True
        lblPlayerOneScore.Visible = True
        lblPlayerTwoScore.Visible = True
    End Sub

    ' Start nupu tegevused. Kui väljad ei ole genereeritud, annab programm vastava teate.
    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        If fieldGenerated = False Then
            MsgBox("Please generate field before starting new game.", MsgBoxStyle.Information, "Info")
            Exit Sub
        End If
        gameStarted = True
        btnStart.Enabled = False
        Label201.Text = "Turn"
    End Sub

    ' Mänguväljaku, mida hakkab ründama esimene mängija, klikkimise mõju realisatsioon. Kui kasutaja klikib ühele ruudule,
    ' siis värvitakse see mustaks ja kiri punaseks (juhul kui seal on miskit). Kui vajutatud ruudus on X, siis lisatakse skoorile +1
    Private Sub Player_Two_Field_Label_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label99.Click, Label98.Click, Label97.Click, Label96.Click, Label95.Click, Label94.Click, Label93.Click, Label92.Click, Label91.Click, Label90.Click, Label9.Click, Label89.Click, Label88.Click, Label87.Click, Label86.Click, Label85.Click, Label84.Click, Label83.Click, Label82.Click, Label81.Click, Label80.Click, Label8.Click, Label79.Click, Label78.Click, Label77.Click, Label76.Click, Label75.Click, Label74.Click, Label73.Click, Label72.Click, Label71.Click, Label70.Click, Label7.Click, Label69.Click, Label68.Click, Label67.Click, Label66.Click, Label65.Click, Label64.Click, Label63.Click, Label62.Click, Label61.Click, Label60.Click, Label6.Click, Label59.Click, Label58.Click, Label57.Click, Label56.Click, Label55.Click, Label54.Click, Label53.Click, Label52.Click, Label51.Click, Label50.Click, Label5.Click, Label49.Click, Label48.Click, Label47.Click, Label46.Click, Label45.Click, Label44.Click, Label43.Click, Label42.Click, Label41.Click, Label40.Click, Label4.Click, Label39.Click, Label38.Click, Label37.Click, Label36.Click, Label35.Click, Label34.Click, Label33.Click, Label32.Click, Label31.Click, Label30.Click, Label3.Click, Label29.Click, Label28.Click, Label27.Click, Label26.Click, Label25.Click, Label24.Click, Label23.Click, Label22.Click, Label21.Click, Label20.Click, Label2.Click, Label19.Click, Label18.Click, Label17.Click, Label16.Click, Label15.Click, Label14.Click, Label13.Click, Label12.Click, Label11.Click, Label100.Click, Label10.Click, Label1.Click
        If playerOneTurn = True And gameStarted = True Then
            Dim clickedLabel = TryCast(sender, Label)
            If clickedLabel IsNot Nothing Then

                If clickedLabel.BackColor = Color.Black Then Exit Sub

                clickedLabel.BackColor = Color.Black
                clickedLabel.ForeColor = Color.Red
                If clickedLabel.Text = "X" Then
                    playerTwoHits = playerTwoHits + 1
                    lblPlayerTwoScore.Text = playerTwoHits
                End If
                If isGameOver() Then
                    gameStarted = False
                    MsgBox("Player One won.", MsgBoxStyle.Information, "Info")
                    Exit Sub
                End If
                playerOneTurn = False
                Label201.Text = ""
                Label202.Text = "Turn"
            End If
        End If
    End Sub

    '' See lõik on praegu tegemisel, võimalus mängida User vs AI.
    'Private Sub Computer_Fires(ByVal field As TableLayoutPanel)
    '    Dim random As New Random
    '    Dim generatedCellNumber = random.Next(101, 201)
    '    Dim cellLabelName = "Label" & generatedCellNumber

    '    For Each control In field.Controls
    '        Dim cellLabel = TryCast(control, Label)
    '        If cellLabel.Name = cellLabelName Then
    '            If cellLabel.BackColor = Color.Black Then

    '            End If
    '            cellLabel.BackColor = Color.Black
    '            cellLabel.ForeColor = Color.Red
    '            If cellLabel.Text = "X" Then
    '                computerHits = computerHits + 1
    '                playerTurn = True
    '                Exit Sub
    '            End If
    '            Exit For
    '        End If
    '    Next
    'End Sub

    ' kontrollib, kas mäng on läbi (kas kummalgil on kõik laevad hävitatud või ei)
    Private Function isGameOver() As Boolean
        If playerOneHits = 17 Or playerTwoHits = 17 Then
            Return True
        End If
        Return False
    End Function

    ' Mänguväljaku, mida hakkab ründama teine mängija, klikkimise mõju realisatsioon. Kui kasutaja klikib ühele ruudule,
    ' siis värvitakse see mustaks ja kiri punaseks (juhul kui seal on miskit). Kui vajutatud ruudus on X, siis lisatakse skoorile +1
    Private Sub Player_One_Field_Label_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label200.Click, Label199.Click, Label198.Click, Label197.Click, Label196.Click, Label195.Click, Label194.Click, Label193.Click, Label192.Click, Label191.Click, Label190.Click, Label189.Click, Label188.Click, Label187.Click, Label186.Click, Label185.Click, Label184.Click, Label183.Click, Label182.Click, Label181.Click, Label180.Click, Label179.Click, Label178.Click, Label177.Click, Label176.Click, Label175.Click, Label174.Click, Label173.Click, Label172.Click, Label171.Click, Label170.Click, Label169.Click, Label168.Click, Label167.Click, Label166.Click, Label165.Click, Label164.Click, Label163.Click, Label162.Click, Label161.Click, Label160.Click, Label159.Click, Label158.Click, Label157.Click, Label156.Click, Label155.Click, Label154.Click, Label153.Click, Label152.Click, Label151.Click, Label150.Click, Label149.Click, Label148.Click, Label147.Click, Label146.Click, Label145.Click, Label144.Click, Label143.Click, Label142.Click, Label141.Click, Label140.Click, Label139.Click, Label138.Click, Label137.Click, Label136.Click, Label135.Click, Label134.Click, Label133.Click, Label132.Click, Label131.Click, Label130.Click, Label129.Click, Label128.Click, Label127.Click, Label126.Click, Label125.Click, Label124.Click, Label123.Click, Label122.Click, Label121.Click, Label120.Click, Label119.Click, Label118.Click, Label117.Click, Label116.Click, Label115.Click, Label114.Click, Label113.Click, Label112.Click, Label111.Click, Label110.Click, Label109.Click, Label108.Click, Label107.Click, Label106.Click, Label105.Click, Label104.Click, Label103.Click, Label102.Click, Label101.Click
        If playerOneTurn = False And gameStarted = True Then
            Dim clickedLabel = TryCast(sender, Label)

            If clickedLabel IsNot Nothing Then

                If clickedLabel.BackColor = Color.Black Then Exit Sub

                clickedLabel.BackColor = Color.Black
                clickedLabel.ForeColor = Color.Red
                If clickedLabel.Text = "X" Then
                    playerOneHits = playerOneHits + 1
                    lblPlayerOneScore.Text = playerOneHits
                End If
                If isGameOver() Then
                    gameStarted = False
                    MsgBox("Player Two won.", MsgBoxStyle.Information, "Info")
                    Exit Sub
                End If
                playerOneTurn = True
                Label201.Text = "Turn"
                Label202.Text = ""
            End If
        End If
    End Sub

    ' Nupu "Generate fields" klikkimise realisatsioon. Alguses tehakse olemasolevad väljad puhtaks (kutsudes vastav funktsioon välja), 
    ' seejärel genereeritakse uued laevade asukohad, kutsudes vastav funktsioon välja. Seis nullitakse samuti.
    Private Sub btnGenerateFields_Click(sender As Object, e As EventArgs) Handles btnGenerateFields.Click
        Clear_Fields(tlpComputer)
        Clear_Fields(tlpPlayer)
        Field_Generator(tlpComputer, "computer")
        Field_Generator(tlpPlayer, "player")
        fieldGenerated = True
        btnStart.Enabled = True
        playerOneTurn = True
        Label201.Text = "Turn"
        Label202.Text = ""
        playerOneHits = 0
        playerTwoHits = 0
    End Sub

    ' Väljade puhastamine. Väljad värvitakse üle ja vanad laevade asukohad nullitakse.
    Private Sub Clear_Fields(ByVal field As TableLayoutPanel)
        For Each control In field.Controls
            Dim cellLabel = TryCast(control, Label)
            cellLabel.Text = ""
            cellLabel.BackColor = Color.SteelBlue
            cellLabel.ForeColor = Color.SteelBlue
        Next
    End Sub

    ' Väljade generaator. Siin genereeritakse kõikide laevade asukohad arvestades asjaolu, et need ei tohi üksteisega kattuda ja väljaku
    ' piiridest välja minna. Mõned read on realiseeritud ainult siis, kui programmi jooksutatakse Debug mode's, et oleks näha, kuhu on laevad pandud
    Private Sub Field_Generator(ByVal field As TableLayoutPanel, ByVal whoseField As String)

        Dim random As New Random
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim firstCellNumber As Integer
        Dim shipDirection As Integer
        Dim cellArray(4) As Integer

        For i = 1 To 5
            firstCellNumber = random.Next(1, 101)
            shipDirection = random.Next(0, 2)

            Select Case i

                ' Aircraft Carrier'i paigutamine väljakule
                Case Ships.aircraftCarrier
                    Select Case shipDirection
                        Case direction.vertical
                            If firstCellNumber > 60 Then
                                firstCellNumber = random.Next(1, 61)
                            End If

                            If whoseField = "player" Then
                                cellArray(0) = firstCellNumber + 100
                            Else
                                cellArray(0) = firstCellNumber
                            End If

                            For k = 1 To 4
                                cellArray(k) = cellArray(k - 1) + 10
                            Next
                        Case direction.horizontal
                            While (firstCellNumber > 6 And firstCellNumber < 11) Or (firstCellNumber > 16 And firstCellNumber < 21) _
                                Or (firstCellNumber > 26 And firstCellNumber < 31) Or (firstCellNumber > 36 And firstCellNumber < 41) _
                                Or (firstCellNumber > 46 And firstCellNumber < 51) Or (firstCellNumber > 56 And firstCellNumber < 61) _
                                Or (firstCellNumber > 66 And firstCellNumber < 71) Or (firstCellNumber > 76 And firstCellNumber < 81) _
                                Or (firstCellNumber > 86 And firstCellNumber < 91) Or (firstCellNumber > 96 And firstCellNumber < 101)

                                firstCellNumber = random.Next(1, 101)
                            End While

                            If whoseField = "player" Then
                                cellArray(0) = firstCellNumber + 100
                            Else
                                cellArray(0) = firstCellNumber
                            End If

                            For k = 1 To 4
                                cellArray(k) = cellArray(k - 1) + 1
                            Next

                    End Select

                    For j = 0 To 4
                        For Each control In field.Controls
                            Dim cellLabel = TryCast(control, Label)
                            Dim cellLabelName = "Label" & cellArray(j)
                            If cellLabel.Name = cellLabelName Then
#If DEBUG Then
                                cellLabel.BackColor = Color.Red
#End If
                                cellLabel.Text = "X"
                                Exit For
                            End If
                        Next
                    Next

                    ' Battleship'i paigutamine väljakule
                Case Ships.Battleship
                    Select Case shipDirection
                        Case direction.vertical
                            Dim err As Boolean = True
                            Dim tempArray(3) As Integer
                            While err = True
                                If firstCellNumber > 70 Then
                                    firstCellNumber = random.Next(0, 71)
                                End If

                                If whoseField = "player" Then
                                    cellArray(0) = firstCellNumber + 100
                                Else
                                    cellArray(0) = firstCellNumber
                                End If

                                For k = 0 To 3
                                    If k <> 0 Then
                                        cellArray(k) = cellArray(k - 1) + 10
                                    End If

                                    For Each control In field.Controls
                                        Dim cellLabel = TryCast(control, Label)
                                        Dim tempLabelName = "Label" & cellArray(k)
                                        If cellLabel.Name = tempLabelName Then
                                            If cellLabel.Text = "X" Then
                                                tempArray(k) = 1
                                            Else
                                                tempArray(k) = 0
                                            End If
                                        End If
                                    Next
                                Next

                                If tempArray(0) = 1 Or tempArray(1) = 1 Or tempArray(2) = 1 Or tempArray(3) = 1 Then
                                    firstCellNumber = random.Next(1, 71)
                                Else
                                    err = False
                                End If
                            End While

                        Case direction.horizontal
                            Dim err = True
                            Dim tempArray(3) As Integer
                            While err = True
                                While (firstCellNumber > 7 And firstCellNumber < 11) Or (firstCellNumber > 17 And firstCellNumber < 21) _
                                    Or (firstCellNumber > 27 And firstCellNumber < 31) Or (firstCellNumber > 37 And firstCellNumber < 41) _
                                    Or (firstCellNumber > 47 And firstCellNumber < 51) Or (firstCellNumber > 57 And firstCellNumber < 61) _
                                    Or (firstCellNumber > 67 And firstCellNumber < 71) Or (firstCellNumber > 77 And firstCellNumber < 81) _
                                    Or (firstCellNumber > 87 And firstCellNumber < 91) Or (firstCellNumber > 97 And firstCellNumber < 101)

                                    firstCellNumber = random.Next(1, 101)
                                End While

                                If whoseField = "player" Then
                                    cellArray(0) = firstCellNumber + 100
                                Else
                                    cellArray(0) = firstCellNumber
                                End If

                                For k = 0 To 3
                                    If k <> 0 Then
                                        cellArray(k) = cellArray(k - 1) + 1
                                    End If

                                    For Each control In field.Controls
                                        Dim cellLabel = TryCast(control, Label)
                                        Dim tempLabelName = "Label" & cellArray(k)
                                        If cellLabel.Name = tempLabelName Then
                                            If cellLabel.Text = "X" Then
                                                tempArray(k) = 1
                                            Else
                                                tempArray(k) = 0
                                            End If
                                        End If
                                    Next
                                Next

                                If tempArray(0) = 1 Or tempArray(1) = 1 Or tempArray(2) = 1 Or tempArray(3) = 1 Then
                                    firstCellNumber = random.Next(1, 101)
                                Else
                                    err = False
                                End If
                            End While
                    End Select

                    For j = 0 To 3
                        For Each control In field.Controls
                            Dim cellLabel = TryCast(control, Label)
                            Dim cellLabelName = "Label" & cellArray(j)
                            If cellLabel.Name = cellLabelName Then
#If DEBUG Then
                                cellLabel.BackColor = Color.Red
#End If
                                cellLabel.Text = "X"
                                Exit For
                            End If
                        Next
                    Next

                    ' Submarine'i paigutamine väljakule
                Case Ships.Submarine
                    Select Case shipDirection
                        Case direction.vertical
                            Dim err As Boolean = True
                            Dim tempArray(2) As Integer
                            While err = True
                                If firstCellNumber > 80 Then
                                    firstCellNumber = random.Next(0, 81)
                                End If

                                If whoseField = "player" Then
                                    cellArray(0) = firstCellNumber + 100
                                Else
                                    cellArray(0) = firstCellNumber
                                End If

                                For k = 0 To 2
                                    If k <> 0 Then
                                        cellArray(k) = cellArray(k - 1) + 10
                                    End If

                                    For Each control In field.Controls
                                        Dim cellLabel = TryCast(control, Label)
                                        Dim tempLabelName = "Label" & cellArray(k)
                                        If cellLabel.Name = tempLabelName Then
                                            If cellLabel.Text = "X" Then
                                                tempArray(k) = 1
                                            Else
                                                tempArray(k) = 0
                                            End If
                                        End If
                                    Next
                                Next

                                If tempArray(0) = 1 Or tempArray(1) = 1 Or tempArray(2) = 1 Then
                                    firstCellNumber = random.Next(1, 81)
                                Else
                                    err = False
                                End If
                            End While

                        Case direction.horizontal
                            Dim err = True
                            Dim tempArray(2) As Integer
                            While err = True
                                While (firstCellNumber > 8 And firstCellNumber < 11) Or (firstCellNumber > 18 And firstCellNumber < 21) _
                                    Or (firstCellNumber > 28 And firstCellNumber < 31) Or (firstCellNumber > 38 And firstCellNumber < 41) _
                                    Or (firstCellNumber > 48 And firstCellNumber < 51) Or (firstCellNumber > 58 And firstCellNumber < 61) _
                                    Or (firstCellNumber > 68 And firstCellNumber < 71) Or (firstCellNumber > 78 And firstCellNumber < 81) _
                                    Or (firstCellNumber > 88 And firstCellNumber < 91) Or (firstCellNumber > 98 And firstCellNumber < 101)

                                    firstCellNumber = random.Next(1, 101)
                                End While

                                If whoseField = "player" Then
                                    cellArray(0) = firstCellNumber + 100
                                Else
                                    cellArray(0) = firstCellNumber
                                End If

                                For k = 0 To 2
                                    If k <> 0 Then
                                        cellArray(k) = cellArray(k - 1) + 1
                                    End If

                                    For Each control In field.Controls
                                        Dim cellLabel = TryCast(control, Label)
                                        Dim tempLabelName = "Label" & cellArray(k)
                                        If cellLabel.Name = tempLabelName Then
                                            If cellLabel.Text = "X" Then
                                                tempArray(k) = 1
                                            Else
                                                tempArray(k) = 0
                                            End If
                                        End If
                                    Next
                                Next

                                If tempArray(0) = 1 Or tempArray(1) = 1 Or tempArray(2) = 1 Then
                                    firstCellNumber = random.Next(1, 101)
                                Else
                                    err = False
                                End If
                            End While
                    End Select

                    For j = 0 To 2
                        For Each control In field.Controls
                            Dim cellLabel = TryCast(control, Label)
                            Dim cellLabelName = "Label" & cellArray(j)
                            If cellLabel.Name = cellLabelName Then
#If DEBUG Then
                                cellLabel.BackColor = Color.Red
#End If
                                cellLabel.Text = "X"
                                Exit For
                            End If
                        Next
                    Next

                    ' Destroyer'i paigutamine väljakule
                Case Ships.Destroyer
                    Select Case shipDirection
                        Case direction.vertical
                            Dim err As Boolean = True
                            Dim tempArray(2) As Integer
                            While err = True
                                If firstCellNumber > 80 Then
                                    firstCellNumber = random.Next(0, 81)
                                End If

                                If whoseField = "player" Then
                                    cellArray(0) = firstCellNumber + 100
                                Else
                                    cellArray(0) = firstCellNumber
                                End If

                                For k = 0 To 2
                                    If k <> 0 Then
                                        cellArray(k) = cellArray(k - 1) + 10
                                    End If

                                    For Each control In field.Controls
                                        Dim cellLabel = TryCast(control, Label)
                                        Dim tempLabelName = "Label" & cellArray(k)
                                        If cellLabel.Name = tempLabelName Then
                                            If cellLabel.Text = "X" Then
                                                tempArray(k) = 1
                                            Else
                                                tempArray(k) = 0
                                            End If
                                        End If
                                    Next
                                Next

                                If tempArray(0) = 1 Or tempArray(1) = 1 Or tempArray(2) = 1 Then
                                    firstCellNumber = random.Next(1, 81)
                                Else
                                    err = False
                                End If
                            End While

                        Case direction.horizontal
                            Dim err = True
                            Dim tempArray(2) As Integer
                            While err = True
                                While (firstCellNumber > 8 And firstCellNumber < 11) Or (firstCellNumber > 18 And firstCellNumber < 21) _
                                    Or (firstCellNumber > 28 And firstCellNumber < 31) Or (firstCellNumber > 38 And firstCellNumber < 41) _
                                    Or (firstCellNumber > 48 And firstCellNumber < 51) Or (firstCellNumber > 58 And firstCellNumber < 61) _
                                    Or (firstCellNumber > 68 And firstCellNumber < 71) Or (firstCellNumber > 78 And firstCellNumber < 81) _
                                    Or (firstCellNumber > 88 And firstCellNumber < 91) Or (firstCellNumber > 98 And firstCellNumber < 101)

                                    firstCellNumber = random.Next(1, 101)
                                End While

                                If whoseField = "player" Then
                                    cellArray(0) = firstCellNumber + 100
                                Else
                                    cellArray(0) = firstCellNumber
                                End If

                                For k = 0 To 2
                                    If k <> 0 Then
                                        cellArray(k) = cellArray(k - 1) + 1
                                    End If

                                    For Each control In field.Controls
                                        Dim cellLabel = TryCast(control, Label)
                                        Dim tempLabelName = "Label" & cellArray(k)
                                        If cellLabel.Name = tempLabelName Then
                                            If cellLabel.Text = "X" Then
                                                tempArray(k) = 1
                                            Else
                                                tempArray(k) = 0
                                            End If
                                        End If
                                    Next
                                Next

                                If tempArray(0) = 1 Or tempArray(1) = 1 Or tempArray(2) = 1 Then
                                    firstCellNumber = random.Next(1, 101)
                                Else
                                    err = False
                                End If
                            End While
                    End Select

                    For j = 0 To 2
                        For Each control In field.Controls
                            Dim cellLabel = TryCast(control, Label)
                            Dim cellLabelName = "Label" & cellArray(j)
                            If cellLabel.Name = cellLabelName Then
#If DEBUG Then
                                cellLabel.BackColor = Color.Red
#End If
                                cellLabel.Text = "X"
                                Exit For
                            End If
                        Next
                    Next

                    ' Patrol boat'i paigutamine väljakule
                Case Ships.patrolBoat
                    Select Case shipDirection
                        Case direction.vertical
                            Dim err As Boolean = True
                            Dim tempArray(1) As Integer
                            While err = True
                                If firstCellNumber > 90 Then
                                    firstCellNumber = random.Next(0, 91)
                                End If

                                If whoseField = "player" Then
                                    cellArray(0) = firstCellNumber + 100
                                Else
                                    cellArray(0) = firstCellNumber
                                End If

                                For k = 0 To 1
                                    If k <> 0 Then
                                        cellArray(k) = cellArray(k - 1) + 10
                                    End If

                                    For Each control In field.Controls
                                        Dim cellLabel = TryCast(control, Label)
                                        Dim tempLabelName = "Label" & cellArray(k)
                                        If cellLabel.Name = tempLabelName Then
                                            If cellLabel.Text = "X" Then
                                                tempArray(k) = 1
                                            Else
                                                tempArray(k) = 0
                                            End If
                                        End If
                                    Next
                                Next

                                If tempArray(0) = 1 Or tempArray(1) = 1 Then
                                    firstCellNumber = random.Next(1, 91)
                                Else
                                    err = False
                                End If
                            End While

                        Case direction.horizontal
                            Dim err = True
                            Dim tempArray(1) As Integer
                            While err = True
                                While (firstCellNumber > 9 And firstCellNumber < 11) Or (firstCellNumber > 19 And firstCellNumber < 21) _
                                    Or (firstCellNumber > 29 And firstCellNumber < 31) Or (firstCellNumber > 39 And firstCellNumber < 41) _
                                    Or (firstCellNumber > 49 And firstCellNumber < 51) Or (firstCellNumber > 59 And firstCellNumber < 61) _
                                    Or (firstCellNumber > 69 And firstCellNumber < 71) Or (firstCellNumber > 79 And firstCellNumber < 81) _
                                    Or (firstCellNumber > 89 And firstCellNumber < 91) Or (firstCellNumber > 99 And firstCellNumber < 101)

                                    firstCellNumber = random.Next(1, 101)
                                End While

                                If whoseField = "player" Then
                                    cellArray(0) = firstCellNumber + 100
                                Else
                                    cellArray(0) = firstCellNumber
                                End If

                                For k = 0 To 1
                                    If k <> 0 Then
                                        cellArray(k) = cellArray(k - 1) + 1
                                    End If

                                    For Each control In field.Controls
                                        Dim cellLabel = TryCast(control, Label)
                                        Dim tempLabelName = "Label" & cellArray(k)
                                        If cellLabel.Name = tempLabelName Then
                                            If cellLabel.Text = "X" Then
                                                tempArray(k) = 1
                                            Else
                                                tempArray(k) = 0
                                            End If
                                        End If
                                    Next
                                Next

                                If tempArray(0) = 1 Or tempArray(1) = 1 Then
                                    firstCellNumber = random.Next(1, 101)
                                Else
                                    err = False
                                End If
                            End While
                    End Select

                    For j = 0 To 1
                        For Each control In field.Controls
                            Dim cellLabel = TryCast(control, Label)
                            Dim cellLabelName = "Label" & cellArray(j)
                            If cellLabel.Name = cellLabelName Then
#If DEBUG Then
                                cellLabel.BackColor = Color.Red
#End If
                                cellLabel.Text = "X"
                                Exit For
                            End If
                        Next
                    Next
            End Select
        Next
    End Sub
End Class