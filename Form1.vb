'Imports System.Globalization

Imports SpeechLib



Public Class Form1



    Dim voice As New SpVoice
    Dim datestr As String
    Dim datesdt As Date
    Dim strDAY As String
    Dim strSpeak As String
    Dim strApptSpeak As String
    Dim strApptday As Date
    Shared Form1 As Object


    Public Sub ReadiCalFile()
        Dim AllText As String = "", LineOfText As String = ""
        Dim DTSTARTvar As String = "DTSTART"
        Dim SUMMARYvar As String = "SUMMARY"
        Dim start As Integer
        Dim datestr As String
        Dim datesdt As Date
        Dim timestr As String
        Dim filename As String
        Dim longtxt As String
        Dim yrstr As String
        Dim mnthstr As String
        Dim daystr As String
        Dim srchdate As Date
        Dim i As Integer


        Dim appointments(30) As String
        Dim AllText2 As String = "", LineOfText2 As String = ""
        TextBox1.Text = AllText2

        filename = "F:\Vista Calendar\ARon's Vista Calendar.ics"

        Try 'open file and trap any errors using handler

            FileOpen(1, filename, OpenMode.Input)

            Do Until EOF(1) 'read all lines from file

                LineOfText = LineInput(1)
                'LineOfText variable is not displayed but used to hold lines of file input
                ' for manipulation & transfer into appointments array
                ' LineOfText & AllText variable can be displayed 
                ' by uncommenting line at end of Sub Procedure

                'add each line to the AllText variable
                If LineOfText.StartsWith(DTSTARTvar) Then
                    'index start of date string in line:
                    start = LineOfText.IndexOf(":2")
                    longtxt = LineOfText.Substring(start + 1)

                    'pull date string from Line of Text
                    datestr = longtxt.Substring(0, 8)

                    If longtxt.Length >= 10 Then
                        timestr = longtxt.Substring(9, 4)
                    ElseIf longtxt.Length <= 10 Then
                        timestr = "0"
                    End If

                    yrstr = datestr.Substring(0, 4)
                    mnthstr = datestr.Substring(4, 2)
                    daystr = datestr.Substring(6)
                    'concatanate datestr to equal valid date string using 3 substrings
                    datestr = mnthstr + "/" + daystr + "/" + yrstr

                    'datesdt = DateValue(datestr)
                    For i = 0 To 30

                        srchdate = Date.Today.AddDays(i)
                        datesdt = DateValue(datestr)
                        'search input line of text for next 30 days dates
                        If datesdt = srchdate Then
                            Dim j As Integer
                            Dim arraystr As String

                            strDAY = datesdt.ToString("ddd")
                            'datesdt.ToString("MMM dd ")
                            'Debug.WriteLine("b4" + strDAY)
                            'insrtSpaces(strDAY)
                            'Debug.WriteLine("af" + strDAY)
                            LineOfText = datestr + "  " + strDAY
                            LineOfText = datestr + "  " + datesdt.DayOfWeek.ToString + "                   a"
                            arraystr = LineOfText.Substring(0, 10) 'was 23
                            AllText = AllText + LineOfText.Substring(0, 10)
                            yrstr = LineInput(1)
                            LineOfText = LineInput(1)
                            LineOfText = LineOfText.Substring(8)

                            Dim lineCR As String
                            Dim lineCR2 As String

                            If CType(LineOfText.IndexOf("\n"), Object) Is Nothing Then
                                LineOfText = LineOfText
                            ElseIf LineOfText.IndexOf("\n") >= 0 Then
                                lineCR2 = LineInput(1)
                                If lineCR2.StartsWith(" ") Then LineOfText = LineOfText + lineCR2.Trim(" ")
                                Do While LineOfText.IndexOf("\n") >= 0
                                    lineCR = LineOfText.Substring((0), LineOfText.IndexOf("\n"))
                                    LineOfText = LineOfText
                                    LineOfText = LineOfText.Remove(0, LineOfText.IndexOf("\n") + 2)
                                    lineCR = lineCR & vbCrLf & LineOfText
                                    LineOfText = lineCR
                                Loop

                                'lineCR = ""
                            End If
                            AllText = AllText & "--" & LineOfText & vbCrLf
                            j = j + 1
                            arraystr = arraystr + "--" + LineOfText

                            appointments(j) = arraystr
                            'Debug.WriteLine(j.ToString)
                            'Debug.WriteLine(appointments(j))
                        End If
                    Next i
                End If
            Loop                   'update label

            'lblNote.Text = OpenFileDialog1.FileName
            'lblNote.Text = filename
            'txtNote.Text = AllText 'display file
            'txtNote.Enabled = True 'allow text cursor
            TextBox1.Enabled = True 'allow text cursor

        Catch
            MsgBox("Error opening file. It might be too big.")
        Finally
            FileClose(1) 'close file
        End Try
        'End If
        Dim s As Integer
        Dim appt As Integer

        Array.Sort(appointments)
        TextBox1.Text = AllText2 = ""

        'Get # of appointments, ignoring empty fields in sorted array
        For s = 0 To 30
            If appointments(s) <> "" Then
                'Debug.WriteLine(s.ToString)
                'LineOfText2 = (appointments(s))
                appt = appt + 1

                'AllText2 = AllText2 & LineOfText2 & vbCrLf
            End If
        Next
        LineOfText2 = "You have " & (appt.ToString) & " Appointments in the next 30 days." & vbCrLf
        strSpeak = LineOfText2

        LineOfText2 = LineOfText2.ToUpper
        AllText2 = AllText2 & LineOfText2

        Dim daysAway As String
        appt = 0
        For s = 0 To 30
            If appointments(s) <> "" Then
                LineOfText2 = (appointments(s))
                datesdt = appointments(s).Substring(0, 10)

                'daysAway = ((datesdt - Date.Today).TotalDays)
                daysAway = ((datesdt - Date.Today).TotalDays)
                LineOfText2 = LineOfText2.Insert(11, "--" & daysAway & " days away---")
                'Debug.WriteLine(LineOfText2.Length)
                LineOfText2 = LineOfText2.Remove(0, 10)
                'datesdt.ToString("MMM dd ") 'ddd= abbrev. dayofweek
                LineOfText2 = LineOfText2.Insert(0, datesdt.ToString("MMM dd" + "--" + " ddd"))
                appt = appt + 1
                If appt = 1 Then
                    strApptday = datesdt
                    strApptSpeak = LineOfText2
                    Debug.WriteLine(strApptday + strApptSpeak)
                End If

                AllText2 = AllText2 & LineOfText2 & vbCrLf
            End If
        Next
        TextBox1.HideSelection = True
        'Uncomment the following line to display the AllText variable
        'TextBox1.Text = AllText
        TextBox1.Text = AllText2

        TextBox1.Select(0, 0)
        TextBox1.ReadOnly = True
        Me.Refresh()
        'TextBox1.Refresh()
        'Label1.Refresh()
        'Timer1.Enabled = False


        ' DoVoice()
        'ISpEventSink: DoVoice()







    End Sub
    Private Sub DoVoice()
        Me.Update()
        Dim Greeting As String = ""
        
        TimeOfDay = System.DateTime.Now

        'call Greet procedure to determine hour of greeting
        Greet(TimeOfDay, Greeting)
       


        voice.Speak(Greeting, SpeechVoiceSpeakFlags.SVSFDefault) '& System.DateTime.Now.ToString("ddddd " + "d  MMMM" + "h:mm tt")) 'System.DateTime.Now.ToString("h m tt"))
        Wait(1)
        voice.Speak("Today is " + System.DateTime.Now.ToString("dddd "))
        Wait(1)
        voice.Speak(strSpeak)
        Wait(1)
        strApptSpeak = strApptSpeak.Remove(0, 12)
        Wait(0.75)
        Debug.WriteLine(strApptSpeak)
        If (strApptSpeak.IndexOf("Dr.") > 0) Then
            strApptSpeak = strApptSpeak.Insert(strApptSpeak.IndexOf("Dr."), "Doctor ")
            strApptSpeak = strApptSpeak.Remove(strApptSpeak.IndexOf("Dr."), 3)
            Debug.WriteLine(strApptSpeak)
        End If

        If strApptday <> Date.Today Then
            voice.Speak("your next appointment is on " + strApptday.ToString("dddd, MMMM, dd"))
            Wait(0.5)
            voice.Speak("  " + strApptSpeak)
        End If

        If strApptday = Date.Today Then
            voice.Speak("you have an appointment today " + strApptday.ToString("dddd, MMMM, dd"))
            Wait(0.5)
            strApptSpeak = strApptSpeak.Remove(0, 16)
            voice.Speak("  " + strApptSpeak)
        End If

        Debug.WriteLine((strApptday.ToString("dddd, MMMM, dd") + strApptSpeak))
        Wait(1)
        voice.Speak("What's next massa Ronnie")
        Wait(0.5)
        voice.Speak("You be duh mahnn!")

    End Sub

    Private Sub Greet(ByVal TimeOfDay As System.DateTime, ByRef Greeting As String)

        Select Case Greeting = ""
            Case TimeOfDay >= TimeValue("12:00") And TimeOfDay <= TimeValue("17:59")
                Greeting = "Good Afternoon Massa Ronnie"
                Return

            Case TimeOfDay >= TimeValue("00:00") And TimeOfDay <= TimeValue("11:59")
                Greeting = "Good Morning Massa Ronnie"
                Return

            Case TimeOfDay >= TimeValue("18:00") And TimeOfDay <= TimeValue("23:59")
                Greeting = "Good Evening Massa Ronnie"
                Return
        End Select

    End Sub


    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Label1.Text = System.DateTime.Now.ToString("ddddd " + "---  " + "d  MMMM" + vbCrLf + "h:mm tt", _
                  Globalization.CultureInfo.InstalledUICulture) '.CreateSpecificCulture("en-US"))


    End Sub

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        My.Settings.Reload()
        
        Me.DesktopLocation = New Point(My.Settings.x, My.Settings.y)
       
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing


        My.Settings.x = Me.Location.X
        My.Settings.y = Me.Location.Y
        
        My.Settings.Save()
        voice.Speak("Goodbye Commander, have a nice day")
        

    End Sub


    Private Sub Form1_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        
        ReadiCalFile()
        TextBox1.Refresh()
        Label1.Refresh()
       

    End Sub


    Public Sub Wait(ByVal seconds As Double)
        Dim newDate As Date
        newDate = Now.AddSeconds(seconds)
        While Now.Ticks <= newDate.Ticks
            Application.DoEvents()
        End While
    End Sub

    


End Class
