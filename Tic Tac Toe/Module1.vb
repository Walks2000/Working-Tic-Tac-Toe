Module Module1
    Public Grid(8)
    Public xInput As String
    Public yInput As String
    Public xWinner As Boolean = False
    Public yWinner As Boolean = False

    Sub Main()
        Console.Clear()
        Grid(0) = "0"
        Grid(1) = "1"
        Grid(2) = "2"
        Grid(3) = "3"
        Grid(4) = "4"
        Grid(5) = "5"
        Grid(6) = "6"
        Grid(7) = "7"
        Grid(8) = "8"
        yWinner = False
        xWinner = False
        Console.Title = "Noughts and Crosses" : Console.ForegroundColor = ConsoleColor.Green
        Console.WriteLine("Welcome to Noughts and Crosses!")
        Console.WriteLine("Before we begin, would you like a quick tutorial to understand ")
        Console.WriteLine("how the game works?")
        Console.WriteLine("1. Yes" & vbCrLf & "2. No")
        Dim UserTutorial As String = Console.ReadLine
        Select Case UserTutorial
            Case 1
                Tutorial()
            Case 2
                Game()
            Case Else
                Main()
        End Select

        Console.ReadLine()
    End Sub
    Sub Tutorial()
        Console.Clear()
        Console.WriteLine("| {0} | {1} | {2} |", Grid(0), Grid(1), Grid(2))
        Console.WriteLine("| {0} | {1} | {2} |", Grid(3), Grid(4), Grid(5))
        Console.WriteLine("| {0} | {1} | {2} |", Grid(6), Grid(7), Grid(8))
        Console.WriteLine("Above is how the game board will look." & vbCrLf & "(Press enter to continue.)")
        Console.ReadLine()
        Console.WriteLine("" & vbCrLf & "Above is the values for each section of the grid.")
        Console.ReadLine()
        Console.WriteLine("Remember them carefully, you will need them throughout.")
        Console.ReadLine()
        Console.WriteLine("You can reset at any time by typing 'Reset' on your turn.")
        Console.ReadLine()
        Console.WriteLine("This concludes the tutorial, good luck, have fun!")

        Console.ReadLine()
        Game()

    End Sub
    Sub Game()
        Dim Turns As Integer = 0
        Dim Players(1)
        Console.WriteLine("The first turn is based on a random generator.")
        Console.WriteLine("If the value is above 50, X gets first go.")
        Dim Random As New Random
        Dim FirstTurn As Integer = Random.Next(0, 101)
        Console.WriteLine("The number generated was {0}", FirstTurn)
        If FirstTurn > 50 Then
            Players(0) = "X"
            Players(1) = "Y"
        Else
            Players(0) = "Y"
            Players(1) = "X"
        End If
        If Players(0) = "X" Then
            Console.WriteLine("X has first turn.")
        Else
            Console.WriteLine("Y has first turn.")
        End If
        Console.ReadLine()
        For i = 1 To 9
            Console.Clear()
            CheckWinX()
            CheckWinY()
            If xWinner = True Then
                Console.WriteLine("Congratulations X! You won the game!")
            End If
            If yWinner = True Then
                Console.WriteLine("Congratulations Y! You won the game!")
            End If

            If xWinner Or yWinner = True Then
                EndGame()
            End If
            DrawBoard()
            Console.WriteLine("It is now turn: {0}", Turns)
            Turns = Turns + 1
            If i Mod 2 = 1 Then
                Console.WriteLine("It is your turn, {0}.", Players(0))
                Console.WriteLine("This is the board as it stands, please enter a grid number now. (0-8)")
                If Players(0) = "X" Then
                    xInputValid()
                Else
                    yInputValid()
                End If
            End If

            If i Mod 2 = 0 Then
                Console.WriteLine("It is your turn, {0}.", Players(1))
                Console.WriteLine("This is the board as it stands, please enter a grid number now. (0-8)")
                If Players(1) = "X" Then
                    xInputValid()
                Else
                    yInputValid()
                End If
            End If


        Next
        Console.Clear()
        DrawBoard()
        Console.WriteLine("It's a draw!")
        EndGame()

    End Sub
    Sub DrawBoard()
        Console.WriteLine("| {0} | {1} | {2} |", Grid(0), Grid(1), Grid(2))
        Console.WriteLine("| {0} | {1} | {2} |", Grid(3), Grid(4), Grid(5))
        Console.WriteLine("| {0} | {1} | {2} |", Grid(6), Grid(7), Grid(8))
    End Sub
    Sub xInputValid()
        xInput = Console.ReadLine()
        If xInput = "Reset" Then
            Reset()
        End If
        Dim xInputGrid As String = xInput

        If xInput > 8 Or xInput < 0 Then
            Dim xInputValid As Boolean = False
            While xInputValid = False
                Console.WriteLine("Sorry, that is not a valid number, try again.")
                xInput = Console.ReadLine()
                If xInput > 8 Or xInput < 0 Then
                    xInputValid = False
                Else
                    xInputValid = True
                End If
            End While
        End If
        If Grid(xInputGrid) = "O" Or Grid(xInputGrid) = "X" Then
            Dim xInputValid As Boolean = False
            While xInputValid = False
                Console.WriteLine("Sorry, that spot is taken, try again.")
                xInput = Console.ReadLine()
                If Grid(xInputGrid) = "O" Or Grid(xInputGrid) = "X" Then
                    xInputValid = False
                Else
                    xInputValid = True
                End If
            End While
        End If
        Grid(xInputGrid) = "X"
    End Sub
    Sub yInputValid()
        yInput = Console.ReadLine()
        If yInput = "Reset" Then
            Reset()
        End If
        Dim yInputGrid As String = yInput
        If yInput > 8 Or yInput < 0 Then
            Dim yInputValid As Boolean = False
            While yInputValid = False
                Console.WriteLine("Sorry, that is not a valid number, try again.")
                yInput = Console.ReadLine()
                If yInput > 8 Or yInput < 0 Then
                    yInputValid = False
                Else
                    yInputValid = True
                End If
            End While
        End If
        If Grid(yInputGrid) = "O" Or Grid(yInputGrid) = "X" Then
            Dim yInputValid As Boolean = False
            While yInputValid = False
                Console.WriteLine("Sorry, that spot is taken, try again.")
                yInput = Console.ReadLine()
                If Grid(yInputGrid) = "O" Or Grid(yInputGrid) = "X" Then
                    yInputValid = False
                Else
                    yInputValid = True
                End If
            End While
        End If
        Grid(yInputGrid) = "Y"
    End Sub
    Function CheckWinX()
        If Grid(0) = "X" And Grid(1) = "X" And Grid(2) = "X" Or
         Grid(0) = "X" And Grid(4) = "X" And Grid(8) = "X" Or
         Grid(1) = "X" And Grid(4) = "X" And Grid(7) = "X" Or
         Grid(2) = "X" And Grid(4) = "X" And Grid(6) = "X" Or
         Grid(3) = "X" And Grid(4) = "X" And Grid(5) = "X" Or
         Grid(2) = "X" And Grid(5) = "X" And Grid(8) = "X" Or
         Grid(0) = "X" And Grid(3) = "X" And Grid(6) = "X" Or
         Grid(6) = "X" And Grid(7) = "X" And Grid(8) = "X" Then
            xWinner = True
        End If
    End Function
    Function CheckWinY()
        If Grid(0) = "Y" And Grid(1) = "Y" And Grid(2) = "Y" Or
         Grid(0) = "Y" And Grid(4) = "Y" And Grid(8) = "Y" Or
         Grid(1) = "Y" And Grid(4) = "Y" And Grid(7) = "Y" Or
         Grid(2) = "Y" And Grid(4) = "Y" And Grid(6) = "Y" Or
         Grid(3) = "Y" And Grid(4) = "Y" And Grid(5) = "Y" Or
         Grid(2) = "Y" And Grid(5) = "Y" And Grid(8) = "Y" Or
         Grid(0) = "Y" And Grid(3) = "Y" And Grid(6) = "Y" Or
         Grid(6) = "Y" And Grid(7) = "Y" And Grid(8) = "Y" Then
            yWinner = True
        End If
    End Function
    Sub EndGame()
        Dim PlayAgain As String = 0
        Console.WriteLine("Would you like to play again?" & vbCrLf & "1. Yes" & vbCrLf & "2. No")
        PlayAgain = Console.ReadLine
        Select Case PlayAgain
            Case 1
                Main()
            Case 2
                End
            Case Else
                Console.WriteLine("Sorry, that is not valid, try again.")
                Console.ReadLine()
                Console.Clear()
                EndGame()
        End Select
    End Sub
    Sub Reset()
        Grid(0) = "0"
        Grid(1) = "1"
        Grid(2) = "2"
        Grid(3) = "3"
        Grid(4) = "4"
        Grid(5) = "5"
        Grid(6) = "6"
        Grid(7) = "7"
        Grid(8) = "8"
        yWinner = False
        xWinner = False
        Game()
    End Sub
End Module
