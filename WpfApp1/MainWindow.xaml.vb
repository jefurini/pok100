Imports System.Runtime.InteropServices
Imports System.Windows.Threading
Imports Discord
Imports Discord.WebSocket

Class MainWindow
    <DllImport("user32.dll")>
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SetCursorPos(ByVal x As Integer, ByVal y As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SendInput(ByVal nInputs As Integer, ByRef pInputs As INPUT, ByVal cbSize As Integer) As Integer
    End Function

    Private WithEvents Client As New DiscordSocketClient
    Private WinhWnd As IntPtr

    Private BotRunning As Boolean = True
    Private TestModeRunning As Boolean = False
    Private MacroThreadLock As New Object

    'Mouse Points
    Private RemoveBuilder As MacroPoint
    Private RemoveResearch As MacroPoint
    Private RemoveTraining As MacroPoint
    Private CloseSearch As MacroPoint
    Private SearchText As MacroPoint
    Private SearchButton As MacroPoint
    Private ConferButton As MacroPoint
    Private ScrollUp As MacroPoint
    Private ScrollDown As MacroPoint

    Private UserNickDict As New Dictionary(Of String, String)
    Private BuilderList As New List(Of String)
    Private ResearchList As New List(Of String)
    Private TrainingList As New List(Of String)
    Private WithEvents BuilderTimer As New DispatcherTimer With {.Interval = New TimeSpan(0, 0, 120)}
    Private WithEvents ResearchTimer As New DispatcherTimer With {.Interval = New TimeSpan(0, 0, 120)}
    Private WithEvents TrainingTimer As New DispatcherTimer With {.Interval = New TimeSpan(0, 0, 120)}

    Private BuffChannel As ISocketMessageChannel = Nothing
    Private POChannel As ISocketMessageChannel = Nothing
    Private LastBuff As String = "Builder"

    Const INPUT_MOUSE As UInt32 = 0
    Const INPUT_KEYBOARD As Integer = 1
    Const INPUT_HARDWARE As Integer = 2
    Const XBUTTON1 As UInt32 = &H1
    Const XBUTTON2 As UInt32 = &H2
    Const MOUSEEVENTF_MOVE As UInt32 = &H1
    Const MOUSEEVENTF_LEFTDOWN As UInt32 = &H2
    Const MOUSEEVENTF_LEFTUP As UInt32 = &H4
    Const MOUSEEVENTF_WHEEL As UInt32 = &H800
    Const MOUSEEVENTF_VIRTUALDESK As UInt32 = &H4000
    Const MOUSEEVENTF_ABSOLUTE As UInt32 = &H8000
    Const KEYEVENTF_KEYDOWN As UInt32 = &H0 'Key down flag
    Const KEYEVENTF_KEYUP As UInt32 = &H2 'Key up flag
    Const VK_LCONTROL As UInt32 = &HA2 'Left Control key code

    Public Async Sub MainAsync()
        Await Client.LoginAsync(TokenType.Bot, "NzAwNDY0MDIxNDQ2NTkwNTU1.XpjUxw.prEa7uP_6u8nvUks-7jxK_zziDk")
        Await Client.StartAsync()

    End Sub

    Private Sub MainWindow_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles Me.Closing
        Client.StopAsync()
    End Sub

    Private Function OnClientMsg(arg As SocketMessage) As Task Handles Client.MessageReceived
        If arg.Author.IsBot = False AndAlso arg.Channel.Name = "po-box" Then
            If POChannel Is Nothing Then
                POChannel = arg.Channel
            End If
            Dim NickName As String = DirectCast(arg.Author, Discord.WebSocket.SocketGuildUser).Nickname
            If NickName = "" Then
                NickName = arg.Author.Username
            End If



            If arg.Content.StartsWith("~") Then
                'ignore messages with ~ to give a way to explain the commands
            Else
                Select Case arg.Content
                    Case "!help"
                        arg.Channel.SendMessageAsync("PO Bot Settings:" & vbCrLf &
                             "!start = Bot starts listening to commands in buff-requests." & vbCrLf &
                             "!stop = Bot stops listening to commands in buff-requests and closes the program." & vbCrLf &
                             "!test = Bot only listens to commands in po-box to test it.")
                    Case "!start"
                        BotRunning = True
                        TestModeRunning = False
                        arg.Channel.SendMessageAsync(NickName & " started PO Bot.")
                    Case "!stop"
                        BotRunning = False
                        TestModeRunning = False
                        arg.Channel.SendMessageAsync(NickName & " stopped PO Bot.")
                        CloseCMD()
                    Case "!test"
                        TestModeRunning = True
                        BotRunning = False
                        arg.Channel.SendMessageAsync(NickName & " started PO Bot in test mode.")
                End Select
            End If
        End If




        If (BotRunning = True AndAlso arg.Author.IsBot = False AndAlso arg.Channel.Name = "buff-requests") OrElse (TestModeRunning = True AndAlso arg.Author.IsBot = False AndAlso arg.Channel.Name = "po-box") Then
            SyncLock MacroThreadLock
                If BuffChannel Is Nothing Then
                    BuffChannel = arg.Channel
                End If
                Dim NickName As String = DirectCast(arg.Author, Discord.WebSocket.SocketGuildUser).Nickname
                Dim UserName As String = arg.Author.Username
                If NickName = "" Then
                    NickName = UserName
                End If

                If arg.Content.StartsWith("~") Then
                    'ignore messages with ~ to give a way to explain the commands
                Else
                    If arg.Content = "!help" Then
                        arg.Channel.SendMessageAsync("!builder , !research or !training to request buffs." & vbCrLf & "!done to end your 2 mins early." & vbCrLf &
                                                     "Commands will work with your Nickname/Username unless you add IGN: {name}. Example: !builder IGN: L0rd Tyrion")
                    Else

                        If arg.Content.StartsWith("!builder") Then
                            If arg.Content.Contains("IGN: ") Then
                                NickName = arg.Content.Substring(arg.Content.IndexOf("IGN: ") + 5)
                            End If

                            If CheckQueue(UserName, NickName) = False Then
                                GrantBuff(BuilderList, BuilderTimer, UserName, NickName, "Builder")
                            End If
                        End If

                        If arg.Content.StartsWith("!research") Then
                            If arg.Content.Contains("IGN: ") Then
                                NickName = arg.Content.Substring(arg.Content.IndexOf("IGN: ") + 5)
                            End If

                            If CheckQueue(UserName, NickName) = False Then
                                GrantBuff(ResearchList, ResearchTimer, UserName, NickName, "Research")
                            End If
                        End If

                        If arg.Content.StartsWith("!training") Then
                            If arg.Content.Contains("IGN: ") Then
                                NickName = arg.Content.Substring(arg.Content.IndexOf("IGN: ") + 5)
                            End If

                            If CheckQueue(UserName, NickName) = False Then
                                GrantBuff(TrainingList, TrainingTimer, UserName, NickName, "Training")
                            End If
                        End If

                        If arg.Content.StartsWith("!done") Then
                            If BuilderList.Contains(UserName) Then
                                If BuilderList(0) = UserName Then
                                    RevokeBuff(BuilderList, BuilderTimer, UserName, "Builder")
                                Else
                                    BuilderList.Remove(UserName)
                                    BuffChannel.SendMessageAsync(NickName & " has been removed from the Builder buff queue.")
                                End If
                            End If

                            If ResearchList.Contains(UserName) Then
                                If ResearchList(0) = UserName Then
                                    RevokeBuff(ResearchList, ResearchTimer, UserName, "Research")
                                Else
                                    ResearchList.Remove(UserName)
                                    BuffChannel.SendMessageAsync(NickName & " has been removed from the Research buff queue.")
                                End If
                            End If

                            If TrainingList.Contains(UserName) Then
                                If TrainingList(0) = UserName Then
                                    RevokeBuff(TrainingList, TrainingTimer, UserName, "Training")
                                Else
                                    TrainingList.Remove(UserName)
                                    BuffChannel.SendMessageAsync(NickName & " has been removed from the Training buff queue.")
                                End If
                            End If
                        End If
                    End If
                End If
            End SyncLock
        End If

        Return Task.CompletedTask
    End Function

    Private Function CheckQueue(ByVal UserName As String, ByVal NickName As String)
        Dim InQueue As Boolean = False
        If BuilderList.Contains(UserName) Then
            If UserNickDict(UserName) = NickName Then
                BuffChannel.SendMessageAsync(NickName & " is already queued for the Builder buff.")
            Else
                BuffChannel.SendMessageAsync(UserName & " already queued " & UserNickDict(UserName) & " for the Builder buff.")
            End If
            InQueue = True
        End If

        If ResearchList.Contains(UserName) Then
            If UserNickDict(UserName) = NickName Then
                BuffChannel.SendMessageAsync(NickName & " is already queued for the Research buff.")
            Else
                BuffChannel.SendMessageAsync(UserName & " already queued " & UserNickDict(UserName) & " for the Research buff.")
            End If
            InQueue = True
        End If

        If TrainingList.Contains(UserName) Then
            If UserNickDict(UserName) = NickName Then
                BuffChannel.SendMessageAsync(NickName & " is already queued for the Training buff.")
            Else
                BuffChannel.SendMessageAsync(UserName & " already queued " & UserNickDict(UserName) & " for the Training buff.")
            End If
            InQueue = True
        End If
        Return InQueue
    End Function

    Private Sub BuilderTimer_Tick(sender As Object, e As EventArgs) Handles BuilderTimer.Tick
        SyncLock MacroThreadLock
            RevokeBuff(BuilderList, BuilderTimer, BuilderList(0), "Builder")
        End SyncLock
    End Sub

    Private Sub ResearchTimer_Tick(sender As Object, e As EventArgs) Handles ResearchTimer.Tick
        SyncLock MacroThreadLock
            RevokeBuff(ResearchList, ResearchTimer, ResearchList(0), "Research")
        End SyncLock
    End Sub

    Private Sub TrainingTimer_Tick(sender As Object, e As EventArgs) Handles TrainingTimer.Tick
        SyncLock MacroThreadLock
            RevokeBuff(TrainingList, TrainingTimer, TrainingList(0), "Training")
        End SyncLock
    End Sub

    Private Sub GrantBuff(ByRef NameList As List(Of String), ByRef BuffTimer As DispatcherTimer, ByVal UserName As String, ByVal NickName As String, ByVal BuffName As String)
        If NameList.Count = 0 Then
            NameList.Add(UserName)
            If UserNickDict.ContainsKey(UserName) Then
                UserNickDict(UserName) = NickName
            Else
                UserNickDict.Add(UserName, NickName)
            End If

            Dim MacroCMD As RequestCMD = RequestCMD.None
            Select Case BuffName
                Case "Builder"
                    MacroCMD = RequestCMD.SetBuilder
                Case "Research"
                    MacroCMD = RequestCMD.SetResearch
                Case "Training"
                    MacroCMD = RequestCMD.SetTraining
            End Select

            StartMouseMacro(MacroCMD, NickName)
            BuffChannel.SendMessageAsync(BuffName & " buff is active for " & NickName & ".")
            BuffTimer.Start()
        Else
            NameList.Add(UserName)
            If UserNickDict.ContainsKey(UserName) Then
                UserNickDict(UserName) = NickName
            Else
                UserNickDict.Add(UserName, NickName)
            End If
            BuffChannel.SendMessageAsync(NickName & " has been added to the " & BuffName & " buff queue." & vbCrLf & "Your current position in queue is: " & NameList.Count - 1)
        End If
    End Sub

    Private Sub RevokeBuff(ByRef NameList As List(Of String), ByRef BuffTimer As DispatcherTimer, ByVal UserName As String, ByVal BuffName As String)
        Dim MacroCMD As RequestCMD = RequestCMD.None
        BuffTimer.Stop()
        NameList.Remove(NameList(0))

        If NameList.Count > 0 Then
            Select Case BuffName
                Case "Builder"
                    MacroCMD = RequestCMD.RepBuilder
                Case "Research"
                    MacroCMD = RequestCMD.RepResearch
                Case "Training"
                    MacroCMD = RequestCMD.RepTraining
            End Select

            StartMouseMacro(MacroCMD, UserNickDict(NameList(0)))
            BuffChannel.SendMessageAsync(BuffName & " buff is active for " & UserNickDict(NameList(0)) & ".")
            BuffTimer.Start()
        Else
            Select Case BuffName
                Case "Builder"
                    MacroCMD = RequestCMD.RemBuilder
                Case "Research"
                    MacroCMD = RequestCMD.RemResearch
                Case "Training"
                    MacroCMD = RequestCMD.RemTraining
            End Select

            StartMouseMacro(MacroCMD, "")
            BuffChannel.SendMessageAsync(BuffName & " is open.")
        End If
    End Sub

    Private Sub MainWindow_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        Dim AppPath As String = System.AppDomain.CurrentDomain.BaseDirectory
        If FileIO.FileSystem.FileExists(AppPath & "setup.ini") = False Then
            Dim str As String = "RemoveBuilder=920;720" & vbCrLf & "RemoveResearch=1420;720" & vbCrLf &
                                "RemoveTraining=920;550" & vbCrLf & "TextInput=800;280" & vbCrLf & "Search=1150;275" & vbCrLf &
                                "Confer=1130;405" & vbCrLf & "Close=1200;200" & vbCrLf & "ScrollUp=1460;620" & vbCrLf & "ScrollDown=1460;840"
            FileIO.FileSystem.WriteAllText(AppPath & "setup.ini", str, False)
        End If
        Dim Settings As String = FileIO.FileSystem.ReadAllText(AppPath & "setup.ini")
        Dim strarray As String() = Settings.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
        For i As Integer = 0 To strarray.Count - 1
            If strarray(i).Contains("=") AndAlso strarray(i).Contains(";") Then
                Dim SetPoint As String = strarray(i).Substring(0, strarray(i).IndexOf("="))
                strarray(i) = strarray(i).Substring(strarray(i).IndexOf("=") + 1)

                Select Case SetPoint
                    Case "RemoveBuilder"
                        Integer.TryParse(strarray(i).Substring(0, strarray(i).IndexOf(";")), RemoveBuilder.X)
                        Integer.TryParse(strarray(i).Substring(strarray(i).IndexOf(";") + 1), RemoveBuilder.Y)
                    Case "RemoveResearch"
                        Integer.TryParse(strarray(i).Substring(0, strarray(i).IndexOf(";")), RemoveResearch.X)
                        Integer.TryParse(strarray(i).Substring(strarray(i).IndexOf(";") + 1), RemoveResearch.Y)
                    Case "RemoveTraining"
                        Integer.TryParse(strarray(i).Substring(0, strarray(i).IndexOf(";")), RemoveTraining.X)
                        Integer.TryParse(strarray(i).Substring(strarray(i).IndexOf(";") + 1), RemoveTraining.Y)
                    Case "TextInput"
                        Integer.TryParse(strarray(i).Substring(0, strarray(i).IndexOf(";")), SearchText.X)
                        Integer.TryParse(strarray(i).Substring(strarray(i).IndexOf(";") + 1), SearchText.Y)
                    Case "Search"
                        Integer.TryParse(strarray(i).Substring(0, strarray(i).IndexOf(";")), SearchButton.X)
                        Integer.TryParse(strarray(i).Substring(strarray(i).IndexOf(";") + 1), SearchButton.Y)
                    Case "Confer"
                        Integer.TryParse(strarray(i).Substring(0, strarray(i).IndexOf(";")), ConferButton.X)
                        Integer.TryParse(strarray(i).Substring(strarray(i).IndexOf(";") + 1), ConferButton.Y)
                    Case "Close"
                        Integer.TryParse(strarray(i).Substring(0, strarray(i).IndexOf(";")), CloseSearch.X)
                        Integer.TryParse(strarray(i).Substring(strarray(i).IndexOf(";") + 1), CloseSearch.Y)
                    Case "ScrollUp"
                        Integer.TryParse(strarray(i).Substring(0, strarray(i).IndexOf(";")), ScrollUp.X)
                        Integer.TryParse(strarray(i).Substring(strarray(i).IndexOf(";") + 1), ScrollUp.Y)
                    Case "ScrollDown"
                        Integer.TryParse(strarray(i).Substring(0, strarray(i).IndexOf(";")), ScrollDown.X)
                        Integer.TryParse(strarray(i).Substring(strarray(i).IndexOf(";") + 1), ScrollDown.Y)
                End Select
            End If
        Next i

        MsgBox("Please select your Game Window and wait for 10 seconds.")
        System.Threading.Thread.Sleep(10000)
        WinhWnd = GetForegroundWindow()
        If WinhWnd <> 0 Then
            MsgBox("Game window successfully selected.")
        Else
            MsgBox("Game window not found.")
            Me.Close()
            Exit Sub
        End If

        MainAsync()
        Me.WindowState = WindowState.Minimized
    End Sub

    Private Sub CloseCMD()
        Dispatcher.Invoke(Sub()
                              Me.Close()
                          End Sub)
    End Sub

    Private Sub SetName(ByVal NickName As String)
        Dispatcher.Invoke(Sub()
                              Windows.Clipboard.SetText(NickName)
                          End Sub)
    End Sub

    Private Sub StartMouseMacro(ByVal CMD As RequestCMD, ByVal NickName As String)
        SetName(NickName)
        SetForegroundWindow(WinhWnd)

        Dim RemBuff As MacroPoint
        Dim ReqScrollDown As Boolean = False
        Dim ReqScrollUp As Boolean = False

        Select Case CMD
            Case RequestCMD.RemBuilder, RequestCMD.RepBuilder, RequestCMD.SetBuilder
                If LastBuff = "Training" Then
                    ReqScrollDown = True
                End If
                LastBuff = "Builder"

                RemBuff.X = RemoveBuilder.X
                RemBuff.Y = RemoveBuilder.Y
            Case RequestCMD.RemResearch, RequestCMD.RepResearch, RequestCMD.SetResearch
                If LastBuff = "Training" Then
                    ReqScrollDown = True
                End If
                LastBuff = "Research"

                RemBuff.X = RemoveResearch.X
                RemBuff.Y = RemoveResearch.Y
            Case RequestCMD.RemTraining, RequestCMD.RepTraining, RequestCMD.SetTraining
                If LastBuff <> "Training" Then
                    ReqScrollUp = True
                End If
                LastBuff = "Training"

                RemBuff.X = RemoveTraining.X
                RemBuff.Y = RemoveTraining.Y
        End Select

        'Close search window (safety)
        System.Threading.Thread.Sleep(1000)
        SetCursorPos(CloseSearch.X, CloseSearch.Y)
        MouseLeftClick()

        'Scroll up/down depending on last requested buff
        System.Threading.Thread.Sleep(1000)
        If ReqScrollDown = True Then
            SetCursorPos(ScrollDown.X, ScrollDown.Y)
            MouseLeftClick()
            System.Threading.Thread.Sleep(500)
            MouseLeftClick()
        End If
        If ReqScrollUp = True Then
            SetCursorPos(ScrollUp.X, ScrollUp.Y)
            MouseLeftClick()
            System.Threading.Thread.Sleep(500)
            MouseLeftClick()
        End If

        Select Case CMD
            Case RequestCMD.RemBuilder, RequestCMD.RemResearch, RequestCMD.RemTraining, RequestCMD.RepBuilder, RequestCMD.RepResearch, RequestCMD.RepTraining
                'Remove current buff user
                System.Threading.Thread.Sleep(1000)
                SetCursorPos(RemBuff.X, RemBuff.Y)
                MouseLeftClick()
            Case Else
                'No current buff user
        End Select

        Select Case CMD
            Case RequestCMD.RepBuilder, RequestCMD.RepResearch, RequestCMD.RepTraining, RequestCMD.SetBuilder, RequestCMD.SetResearch, RequestCMD.SetTraining
                'Confer buff to new user
                System.Threading.Thread.Sleep(1000)
                SetCursorPos(RemBuff.X - 100, RemBuff.Y)
                MouseLeftClick()

                'User entry textbox
                System.Threading.Thread.Sleep(1000)
                SetCursorPos(SearchText.X, SearchText.Y)
                MouseLeftClick()

                'Paste user name
                System.Threading.Thread.Sleep(500)
                STRGV()

                'Search user name
                SetCursorPos(SearchButton.X, SearchButton.Y)
                MouseLeftClick()

                'Confer buff to new user
                System.Threading.Thread.Sleep(1000)
                SetCursorPos(ConferButton.X, ConferButton.Y)
                MouseLeftClick()

                'Close search window
                System.Threading.Thread.Sleep(1000)
                SetCursorPos(CloseSearch.X, CloseSearch.Y)
                MouseLeftClick()
            Case Else
                'Only remove buff
        End Select
    End Sub

    Public Shared Sub MouseLeftClick()
        Dim it As New INPUT
        it.dwType = INPUT_MOUSE
        it.mkhi.mi.dwFlags = MOUSEEVENTF_LEFTDOWN
        SendInput(1, it, Marshal.SizeOf(it))

        System.Threading.Thread.Sleep(50)

        Dim it2 As New INPUT
        it2.dwType = INPUT_MOUSE
        it2.mkhi.mi.dwFlags = MOUSEEVENTF_LEFTUP
        SendInput(1, it2, Marshal.SizeOf(it2))
    End Sub

    Public Shared Sub STRGV()
        Dim it As New INPUT
        it.dwType = INPUT_KEYBOARD
        it.mkhi.ki.wVk = VK_LCONTROL
        it.mkhi.ki.dwFlags = KEYEVENTF_KEYDOWN
        SendInput(1, it, Marshal.SizeOf(it))

        System.Threading.Thread.Sleep(10)

        Dim it2 As New INPUT
        it2.dwType = INPUT_KEYBOARD
        it2.mkhi.ki.wVk = &H56 'V
        it2.mkhi.ki.dwFlags = KEYEVENTF_KEYDOWN
        SendInput(1, it2, Marshal.SizeOf(it2))

        System.Threading.Thread.Sleep(100)

        Dim it3 As New INPUT
        it3.dwType = INPUT_KEYBOARD
        it3.mkhi.ki.wVk = VK_LCONTROL
        it3.mkhi.ki.dwFlags = KEYEVENTF_KEYUP
        SendInput(1, it3, Marshal.SizeOf(it3))

        System.Threading.Thread.Sleep(10)

        Dim it4 As New INPUT
        it4.dwType = INPUT_KEYBOARD
        it4.mkhi.ki.wVk = &H56 'V
        it4.mkhi.ki.dwFlags = KEYEVENTF_KEYUP
        SendInput(1, it4, Marshal.SizeOf(it4))
    End Sub

    Private Structure INPUT
        Dim dwType As Integer
        Dim mkhi As MOUSEKEYBDHARDWAREINPUT
    End Structure

    <StructLayout(LayoutKind.Explicit)>
    Private Structure MOUSEKEYBDHARDWAREINPUT
        <FieldOffset(0)> Public mi As MOUSEINPUT
        <FieldOffset(0)> Public ki As KEYBDINPUT
        <FieldOffset(0)> Public hi As HARDWAREINPUT
    End Structure

    Private Structure MOUSEINPUT
        Public dx As Integer
        Public dy As Integer
        Public mouseData As Integer
        Public dwFlags As Integer
        Public time As Integer
        Public dwExtraInfo As Integer
    End Structure

    Private Structure KEYBDINPUT
        Public wVk As Short
        Public wScan As Short
        Public dwFlags As Integer
        Public time As Integer
        Public dwExtraInfo As Integer
    End Structure

    Private Structure HARDWAREINPUT
        Public uMsg As Integer
        Public wParamL As Short
        Public wParamH As Short
    End Structure
End Class

Public Structure MacroPoint

    Public X As Integer
    Public Y As Integer

End Structure

Public Enum RequestCMD

    None = 0
    SetBuilder = 1
    SetResearch = 2
    SetTraining = 3
    RepBuilder = 5
    RepResearch = 6
    RepTraining = 7
    RemBuilder = 10
    RemResearch = 11
    RemTraining = 12
    KeepBuilder = 15
    KeepResearch = 16
    KeepTraining = 17

End Enum
