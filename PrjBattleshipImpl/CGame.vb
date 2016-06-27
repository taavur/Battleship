Public Class CGame
    Implements IGame

    Public Enum Ships
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
    Public playerOneHits As Integer = 0
    Public playerTwoHits As Integer = 0
    Public playerOneTurn As Boolean = True
    Public gameStarted As Boolean = False
    Public fieldGenerated As Boolean = False

    Private Sub Clear_Fields(field As Object) Implements IGame.Clear_Fields
        Dim i As Integer
        i = 0
    End Sub

    Private Sub Field_Generator(field As Object, whoseField As String) Implements IGame.Field_Generator
        Dim i As Integer
        i = 0

    End Sub
End Class
