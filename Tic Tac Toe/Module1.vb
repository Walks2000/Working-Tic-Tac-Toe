Module Module1
    Public Grid(8)
    Public xInput As String
    Public yInput As String
    Public xWinner As Boolean = False
    Public yWinner As Boolean = False
    Public GameType As String

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
        Console.WriteLine("1. Tutorial" & vbCrLf & "2. Competitive AI" & vbCrLf & "3. Fun AI" & vbCrLf & "4. Player vs Player")
        Dim UserTutorial As String = Console.ReadLine
        Select Case UserTutorial
            Case 1
                Tutorial()
            Case 2
                GameType = "AI"
                GameAI()
            Case 3
                GameType = "FunAI"
                FunAI()
            Case 4
                GameType = "PVP"
                GamePVP()
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
        Main()
    End Sub
    Sub GamePVP()
        Console.Clear()
        Dim Turns As Integer = 0
        Dim Players(1)
        Console.WriteLine("The first turn is based on a random generator.")
        Console.WriteLine("If the value is above 50, X gets first go.")
        Dim Random As New Random
        Dim FirstTurn As Integer = Random.Next(0, 101)
        Console.WriteLine("The number generated was {0}", FirstTurn)
        If FirstTurn > 50 Then
            Players(0) = "X"
            Players(1) = "O"
        Else
            Players(0) = "O"
            Players(1) = "X"
        End If
        If Players(0) = "X" Then
            Console.WriteLine("X has first turn.")
        Else
            Console.WriteLine("O has first turn.")
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
                Console.WriteLine("Congratulations O! You won the game!")
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
        CheckWinX()
        CheckWinY()
        If xWinner = True Then
            Console.WriteLine("Congratulations X! You won the game!")
        End If
        If yWinner = True Then
            Console.WriteLine("Congratulations O! You won the game!")
        End If

        If xWinner Or yWinner = True Then
            EndGame()
        End If
        Console.Clear()
        DrawBoard()
        Console.WriteLine("It's a draw!")
        EndGame()

    End Sub
    Sub GameAI()
        Console.Clear()
        Dim Turns As Integer = 0
        Dim Players(1)
        Console.WriteLine("In AI mode, you get to choose who goes first.")
        Console.WriteLine("X = you, O = AI")
        Dim FirstTurn As String
        FirstTurn = Console.ReadLine()
        Dim FirstTurnError As Boolean = True
        While FirstTurnError = True
            If FirstTurn = "X" Then
                FirstTurnError = False
                Players(0) = "X"
                Players(1) = "O"
            End If
            If FirstTurn = "O" Then
                FirstTurnError = False
                Players(0) = "O"
                Players(1) = "X"
            End If
        End While
        If Players(0) = "X" Then
            Console.WriteLine("X has first turn.")
        Else
            Console.WriteLine("O has first turn.")
        End If
        Console.ReadLine()
        For i = 1 To 9
            Console.Clear()
            DrawBoard()
            CheckWinX()
            CheckWinY()
            If xWinner = True Then
                Console.WriteLine("Congratulations X! You won the game!")
            End If
            If yWinner = True Then
                Console.WriteLine("Congratulations O! You won the game!")
            End If

            If xWinner Or yWinner = True Then
                EndGame()
            End If
            Console.WriteLine("It is now turn: {0}", Turns)
            Turns = Turns + 1
            If i Mod 2 = 1 Then
                If Players(0) = "X" Then
                    Console.WriteLine("It is your turn, {0}.", Players(0))
                    Console.WriteLine("This is the board as it stands, please enter a grid number now. (0-8)")
                    xInputValid()
                Else
                    ExperimentalAI()
                End If
            End If

            If i Mod 2 = 0 Then
                If Players(1) = "X" Then
                    Console.WriteLine("It is your turn, {0}.", Players(1))
                    Console.WriteLine("This is the board as it stands, please enter a grid number now. (0-8)")
                    xInputValid()
                Else
                    ExperimentalAI()
                End If
            End If


        Next
        CheckWinX()
        CheckWinY()
        If xWinner = True Then
            Console.WriteLine("Congratulations X! You won the game!")
        End If
        If yWinner = True Then
            Console.WriteLine("Congratulations O! You won the game!")
        End If

        If xWinner Or yWinner = True Then
            EndGame()
        End If
        Console.Clear()
        DrawBoard()
        Console.WriteLine("It's a draw!")
        EndGame()
    End Sub
    Sub FunAI()
        Console.Clear()
        Dim Turns As Integer = 0
        Dim Players(1)
        Console.WriteLine("In AI mode, you get to choose who goes first.")
        Console.WriteLine("X = you, O = AI")
        Dim FirstTurn As String
        FirstTurn = Console.ReadLine()
        Dim FirstTurnError As Boolean = True
        While FirstTurnError = True
            If FirstTurn = "X" Then
                FirstTurnError = False
                Players(0) = "X"
                Players(1) = "O"
            End If
            If FirstTurn = "O" Then
                FirstTurnError = False
                Players(0) = "O"
                Players(1) = "X"
            End If
        End While
        If Players(0) = "X" Then
            Console.WriteLine("X has first turn.")
        Else
            Console.WriteLine("O has first turn.")
        End If
        Console.ReadLine()
        For i = 1 To 9
            Console.Clear()
            DrawBoard()
            CheckWinX()
            CheckWinY()
            If xWinner = True Then
                Console.WriteLine("Congratulations X! You won the game!")
            End If
            If yWinner = True Then
                Console.WriteLine("Congratulations O! You won the game!")
            End If

            If xWinner Or yWinner = True Then
                EndGame()
            End If
            Console.WriteLine("It is now turn: {0}", Turns)
            Turns = Turns + 1
            If i Mod 2 = 1 Then
                If Players(0) = "X" Then
                    Console.WriteLine("It is your turn, {0}.", Players(0))
                    Console.WriteLine("This is the board as it stands, please enter a grid number now. (0-8)")
                    xInputValid()
                Else
                    ValidFunAI()
                End If
            End If

            If i Mod 2 = 0 Then
                If Players(1) = "X" Then
                    Console.WriteLine("It is your turn, {0}.", Players(1))
                    Console.WriteLine("This is the board as it stands, please enter a grid number now. (0-8)")
                    xInputValid()
                Else
                    ValidFunAI()
                End If
            End If


        Next
        CheckWinX()
        CheckWinY()
        If xWinner = True Then
            Console.WriteLine("Congratulations X! You won the game!")
        End If
        If yWinner = True Then
            Console.WriteLine("Congratulations O! You won the game!")
            DrawBoard()
        End If

        If xWinner Or yWinner = True Then
            EndGame()
        End If
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

        If xInput = "Menu" Then
            Main()
        End If

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

        If xInput = "Menu" Then
            Main()
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
        Grid(yInputGrid) = "O"
    End Sub
    Sub ValidFunAI()
        Dim MoveUsed As Boolean
        MoveUsed = False
        If Grid(0) = "X" And Grid(1) = "X" And MoveUsed = False And Grid(2) <> "X" And Grid(2) <> "O" Then
            Grid(2) = "O"
            MoveUsed = True
        End If
        If Grid(0) = "X" And Grid(4) = "X" And MoveUsed = False And Grid(8) <> "X" And Grid(8) <> "O" Then
            Grid(8) = "O"
            MoveUsed = True
        End If
        If Grid(1) = "X" And Grid(4) = "X" And MoveUsed = False And Grid(7) <> "X" And Grid(7) <> "O" Then
            Grid(7) = "O"
            MoveUsed = True
        End If
        If Grid(3) = "X" And Grid(4) = "X" And MoveUsed = False And Grid(5) <> "X" And Grid(5) <> "O" Then
            Grid(5) = "O"
            MoveUsed = True
        End If
        If Grid(4) = "X" And Grid(5) = "X" And MoveUsed = False And Grid(3) <> "X" And Grid(3) <> "O" Then
            Grid(3) = "O"
            MoveUsed = True
        End If
        If Grid(2) = "X" And Grid(4) = "X" And MoveUsed = False And Grid(6) <> "X" And Grid(6) <> "O" Then
            Grid(6) = "O"
            MoveUsed = True
        End If
        If Grid(4) = "X" And Grid(8) = "X" And MoveUsed = False And Grid(0) <> "X" And Grid(0) <> "O" Then
            Grid(0) = "O"
            MoveUsed = True
        End If
        If Grid(6) = "X" And Grid(4) = "X" And MoveUsed = False And Grid(2) <> "X" And Grid(2) <> "O" Then
            Grid(2) = "O"
            MoveUsed = True
        End If
        If Grid(6) = "X" And Grid(7) = "X" And MoveUsed = False And Grid(8) <> "X" And Grid(8) <> "O" Then
            Grid(8) = "O"
            MoveUsed = True
        End If
        If Grid(7) = "X" And Grid(8) = "X" And MoveUsed = False And Grid(6) <> "X" And Grid(6) <> "O" Then
            Grid(6) = "O"
            MoveUsed = True
        End If
        If Grid(0) = "X" And Grid(2) = "X" And MoveUsed = False And Grid(1) <> "X" And Grid(1) <> "O" Then
            Grid(1) = "O"
            MoveUsed = True
        End If
        If Grid(1) = "X" And Grid(7) = "X" And MoveUsed = False And Grid(4) <> "X" And Grid(4) <> "O" Then
            Grid(4) = "O"
            MoveUsed = True
        End If
        If Grid(3) = "X" And Grid(5) = "X" And MoveUsed = False And Grid(4) <> "X" And Grid(4) <> "O" Then
            Grid(4) = "O"
            MoveUsed = True
        End If
        If Grid(6) = "X" And Grid(8) = "X" And MoveUsed = False And Grid(7) <> "X" And Grid(7) <> "O" Then
            Grid(7) = "O"
            MoveUsed = True
        End If
        If Grid(0) = "X" And Grid(8) = "X" And MoveUsed = False And Grid(4) <> "X" And Grid(4) <> "O" Then
            Grid(4) = "O"
            MoveUsed = True
        End If
        If Grid(2) = "X" And Grid(6) = "X" And MoveUsed = False And Grid(4) <> "X" And Grid(4) <> "O" Then
            Grid(4) = "O"
            MoveUsed = True
        End If
        If MoveUsed = False Then
            Dim RandomNumberBool As Boolean = True
            Dim Random As New Random
            Dim RandomNumber As Integer
            While RandomNumberBool = True
                RandomNumber = Random.Next(0, 8)
                If Grid(RandomNumber) = "O" Or Grid(RandomNumber) = "X" Then
                    RandomNumberBool = True
                Else
                    Grid(RandomNumber) = "O"
                    RandomNumberBool = False
                End If
            End While
        End If
    End Sub
    Sub CheckWinX()
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
    End Sub
    Sub CheckWinY()
        If Grid(0) = "O" And Grid(1) = "O" And Grid(2) = "O" Or
         Grid(0) = "O" And Grid(4) = "O" And Grid(8) = "O" Or
         Grid(1) = "O" And Grid(4) = "O" And Grid(7) = "O" Or
         Grid(2) = "O" And Grid(4) = "O" And Grid(6) = "O" Or
         Grid(3) = "O" And Grid(4) = "O" And Grid(5) = "O" Or
         Grid(2) = "O" And Grid(5) = "O" And Grid(8) = "O" Or
         Grid(0) = "O" And Grid(3) = "O" And Grid(6) = "O" Or
         Grid(6) = "O" And Grid(7) = "O" And Grid(8) = "O" Then
            yWinner = True
        End If
    End Sub
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
        If GameType = "AI" Then
            GameAI()
        End If
        If GameType = "PVP" Then
            GamePVP()
        End If
        If GameType = "FunAI" Then
            FunAI()
        End If
    End Sub

    Sub ChooseGame()
        Dim GameChoice As String
        Console.Clear()
        Console.WriteLine("You now have the choice of doing:")
        Console.WriteLine("1. Player vs Player")
        Console.WriteLine("2. Player vs AI")
        GameChoice = Console.ReadLine
        Select Case GameChoice
            Case 1
                GamePVP()
            Case 2
                GameAI()
            Case Else
                Dim GameChoiceInvalid As Boolean = True
                While GameChoiceInvalid = False
                    Console.WriteLine("Sorry, that is not a valid input, try again.")
                    GameChoice = Console.ReadLine
                    Select Case GameChoice
                        Case 1
                            GameChoiceInvalid = False
                            GamePVP()
                        Case 2
                            GameChoiceInvalid = False
                            GameAI()
                    End Select
                End While
        End Select
    End Sub
    Sub ExperimentalAI()
        'Trying to develop a minmax feature, but might go the for the 'fun' AI first due to the easier nature.
        Dim MoveUsed As Boolean
        MoveUsed = False
        If Grid(0) = "O" And Grid(1) = "O" And MoveUsed = False And Grid(2) <> "X" And Grid(1) <> "X" And Grid(2) <> "Y" Then
            Grid(2) = "O"
            MoveUsed = True
        End If
        If Grid(0) = "O" And Grid(4) = "O" And MoveUsed = False And Grid(8) <> "X" And Grid(4) <> "X" And Grid(8) <> "Y" Then
            Grid(8) = "O"
            MoveUsed = True
        End If
        If Grid(1) = "O" And Grid(4) = "O" And MoveUsed = False And Grid(7) <> "X" And Grid(4) <> "X" And Grid(7) <> "Y" Then
            Grid(7) = "O"
            MoveUsed = True
        End If
        If Grid(3) = "O" And Grid(4) = "O" And MoveUsed = False And Grid(5) <> "X" And Grid(4) <> "X" And Grid(5) <> "Y" Then
            Grid(5) = "O"
            MoveUsed = True
        End If
        If Grid(4) = "O" And Grid(5) = "O" And MoveUsed = False And Grid(3) <> "X" And Grid(5) <> "X" And Grid(3) <> "Y" Then
            Grid(3) = "O"
            MoveUsed = True
        End If
        If Grid(2) = "O" And Grid(4) = "O" And MoveUsed = False And Grid(6) <> "X" And Grid(4) <> "X" And Grid(6) <> "Y" Then
            Grid(6) = "O"
            MoveUsed = True
        End If
        If Grid(4) = "O" And Grid(8) = "O" And MoveUsed = False And Grid(0) <> "X" And Grid(8) <> "X" And Grid(0) <> "Y" Then
            Grid(0) = "O"
            MoveUsed = True
        End If
        If Grid(6) = "O" And Grid(4) = "O" And MoveUsed = False And Grid(2) <> "X" And Grid(2) <> "X" And Grid(4) <> "Y" Then
            Grid(2) = "O"
            MoveUsed = True
        End If
        If Grid(6) = "O" And Grid(7) = "O" And MoveUsed = False And Grid(8) <> "X" And Grid(7) <> "X" And Grid(8) <> "Y" Then
            Grid(8) = "O"
            MoveUsed = True
        End If
        If Grid(7) = "O" And Grid(8) = "O" And MoveUsed = False And Grid(6) <> "X" And Grid(8) <> "X" And Grid(6) <> "Y" Then
            Grid(6) = "O"
            MoveUsed = True
        End If
        If Grid(0) = "O" And Grid(2) = "O" And MoveUsed = False And Grid(1) <> "X" And Grid(2) <> "X" And Grid(1) <> "Y" Then
            Grid(1) = "O"
            MoveUsed = True
        End If
        If Grid(1) = "O" And Grid(7) = "O" And MoveUsed = False And Grid(4) <> "X" And Grid(7) <> "X" And Grid(4) <> "Y" Then
            Grid(4) = "O"
            MoveUsed = True
        End If
        If Grid(3) = "O" And Grid(5) = "O" And MoveUsed = False And Grid(4) <> "X" And Grid(5) <> "X" And Grid(4) <> "Y" Then
            Grid(4) = "O"
            MoveUsed = True
        End If
        If Grid(6) = "O" And Grid(8) = "O" And MoveUsed = False And Grid(7) <> "X" And Grid(8) <> "X" And Grid(7) <> "Y" Then
            Grid(7) = "O"
            MoveUsed = True
        End If
        If Grid(0) = "O" And Grid(8) = "O" And MoveUsed = False And Grid(4) <> "X" And Grid(8) <> "X" And Grid(4) <> "Y" Then
            Grid(4) = "O"
            MoveUsed = True
        End If
        If Grid(2) = "O" And Grid(6) = "O" And MoveUsed = False And Grid(4) <> "X" And Grid(6) <> "X" And Grid(4) <> "Y" Then
            Grid(4) = "O"
            MoveUsed = True
        End If
        'End of winning moves, begins blocking player win.

        If Grid(0) = "X" And Grid(1) = "X" And MoveUsed = False And Grid(2) <> "X" And Grid(2) <> "O" Then
            Grid(2) = "O"
            MoveUsed = True
        End If
        If Grid(0) = "X" And Grid(4) = "X" And MoveUsed = False And Grid(8) <> "X" And Grid(8) <> "O" Then
            Grid(8) = "O"
            MoveUsed = True
        End If
        If Grid(1) = "X" And Grid(4) = "X" And MoveUsed = False And Grid(7) <> "X" And Grid(7) <> "O" Then
            Grid(7) = "O"
            MoveUsed = True
        End If
        If Grid(3) = "X" And Grid(4) = "X" And MoveUsed = False And Grid(5) <> "X" And Grid(5) <> "O" Then
            Grid(5) = "O"
            MoveUsed = True
        End If
        If Grid(4) = "X" And Grid(5) = "X" And MoveUsed = False And Grid(3) <> "X" And Grid(3) <> "O" Then
            Grid(3) = "O"
            MoveUsed = True
        End If
        If Grid(2) = "X" And Grid(4) = "X" And MoveUsed = False And Grid(6) <> "X" And Grid(6) <> "O" Then
            Grid(6) = "O"
            MoveUsed = True
        End If
        If Grid(4) = "X" And Grid(8) = "X" And MoveUsed = False And Grid(0) <> "X" And Grid(0) <> "O" Then
            Grid(0) = "O"
            MoveUsed = True
        End If
        If Grid(6) = "X" And Grid(4) = "X" And MoveUsed = False And Grid(2) <> "X" And Grid(2) <> "O" Then
            Grid(2) = "O"
            MoveUsed = True
        End If
        If Grid(6) = "X" And Grid(7) = "X" And MoveUsed = False And Grid(8) <> "X" And Grid(8) <> "O" Then
            Grid(8) = "O"
            MoveUsed = True
        End If
        If Grid(7) = "X" And Grid(8) = "X" And MoveUsed = False And Grid(6) <> "X" And Grid(6) <> "O" Then
            Grid(6) = "O"
            MoveUsed = True
        End If
        If Grid(0) = "X" And Grid(2) = "X" And MoveUsed = False And Grid(1) <> "X" And Grid(1) <> "O" Then
            Grid(1) = "O"
            MoveUsed = True
        End If
        If Grid(1) = "X" And Grid(7) = "X" And MoveUsed = False And Grid(4) <> "X" And Grid(4) <> "O" Then
            Grid(4) = "O"
            MoveUsed = True
        End If
        If Grid(3) = "X" And Grid(5) = "X" And MoveUsed = False And Grid(4) <> "X" And Grid(4) <> "O" Then
            Grid(4) = "O"
            MoveUsed = True
        End If
        If Grid(6) = "X" And Grid(8) = "X" And MoveUsed = False And Grid(7) <> "X" And Grid(7) <> "O" Then
            Grid(7) = "O"
            MoveUsed = True
        End If
        If Grid(0) = "X" And Grid(8) = "X" And MoveUsed = False And Grid(4) <> "X" And Grid(4) <> "O" Then
            Grid(4) = "O"
            MoveUsed = True
        End If
        If Grid(2) = "X" And Grid(6) = "X" And MoveUsed = False And Grid(4) <> "X" And Grid(4) <> "O" Then
            Grid(4) = "O"
            MoveUsed = True
        End If
        'End of blocking moves, beginning of win setups

        If Grid(0) = "O" And Grid(2) <> "X" And Grid(1) <> "X" And MoveUsed = False And Grid(2) <> "O" And Grid(1) <> "O" Then
            Grid(2) = "O"
            MoveUsed = True
        End If
        If Grid(0) = "O" And Grid(3) <> "X" And Grid(6) <> "X" And MoveUsed = False And Grid(3) <> "O" And Grid(6) <> "O" Then
            Grid(6) = "O"
            MoveUsed = True
        End If
        If Grid(0) = "O" And Grid(4) <> "X" And Grid(8) <> "X" And MoveUsed = False And Grid(4) <> "O" And Grid(8) <> "O" Then
            Grid(8) = "O"
            MoveUsed = True
        End If
        If Grid(2) = "O" And Grid(5) <> "X" And Grid(8) <> "X" And MoveUsed = False And Grid(5) <> "O" And Grid(8) <> "O" Then
            Grid(8) = "O"
            MoveUsed = True
        End If
        If Grid(2) = "O" And Grid(4) <> "X" And Grid(6) <> "X" And MoveUsed = False And Grid(6) <> "O" And Grid(4) <> "O" Then
            Grid(6) = "O"
            MoveUsed = True
        End If
        If Grid(2) = "O" And Grid(1) <> "X" And Grid(0) <> "X" And MoveUsed = False And Grid(0) <> "O" And Grid(1) <> "O" Then
            Grid(0) = "O"
            MoveUsed = True
        End If
        If Grid(8) = "O" And Grid(6) <> "X" And Grid(7) <> "X" And MoveUsed = False And Grid(6) <> "O" And Grid(7) <> "O" Then
            Grid(6) = "O"
            MoveUsed = True
        End If
        If Grid(8) = "O" And Grid(4) <> "X" And Grid(0) <> "X" And MoveUsed = False And Grid(4) <> "O" And Grid(0) <> "O" Then
            Grid(0) = "O"
            MoveUsed = True
        End If
        If Grid(8) = "O" And Grid(2) <> "X" And Grid(5) <> "X" And MoveUsed = False And Grid(2) <> "O" And Grid(5) <> "O" Then
            Grid(2) = "O"
            MoveUsed = True
        End If
        If Grid(6) = "O" And Grid(0) <> "X" And Grid(3) <> "X" And MoveUsed = False And Grid(0) <> "O" And Grid(3) <> "O" Then
            Grid(0) = "O"
            MoveUsed = True
        End If
        If Grid(6) = "O" And Grid(2) <> "X" And Grid(4) <> "X" And MoveUsed = False And Grid(2) <> "O" And Grid(4) <> "O" Then
            Grid(2) = "O"
            MoveUsed = True
        End If
        If Grid(6) = "O" And Grid(7) <> "X" And Grid(8) <> "X" And MoveUsed = False And Grid(7) <> "O" And Grid(8) <> "O" Then
            Grid(8) = "O"
            MoveUsed = True
        End If
        If Grid(1) = "O" And Grid(7) <> "X" And Grid(4) <> "X" And MoveUsed = False And Grid(7) <> "O" And Grid(4) <> "O" Then
            Grid(7) = "O"
            MoveUsed = True
        End If
        If Grid(5) = "O" And Grid(4) <> "X" And Grid(3) <> "X" And MoveUsed = False And Grid(3) <> "O" And Grid(4) <> "O" Then
            Grid(3) = "O"
            MoveUsed = True
        End If
        If Grid(7) = "O" And Grid(1) <> "X" And Grid(4) <> "X" And MoveUsed = False And Grid(4) <> "O" And Grid(1) <> "O" Then
            Grid(1) = "O"
            MoveUsed = True
        End If
        If Grid(3) = "O" And Grid(5) <> "X" And Grid(4) <> "X" And MoveUsed = False And Grid(5) <> "O" And Grid(4) <> "O" Then
            Grid(5) = "O"
            MoveUsed = True
        End If
        If MoveUsed = False Then
            Dim RandomNumberBool As Boolean = True
            Dim Random As New Random
            Dim RandomNumber As Integer
            While RandomNumberBool = True
                RandomNumber = Random.Next(0, 8)
                If Grid(RandomNumber) = "O" Or Grid(RandomNumber) = "X" Then
                    RandomNumberBool = True
                Else
                    Grid(RandomNumber) = "O"
                    RandomNumberBool = False
                    MoveUsed = True
                End If
            End While
        End If
    End Sub
End Module
