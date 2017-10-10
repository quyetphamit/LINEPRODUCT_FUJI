Imports System
Imports System.IO
Public Class Control

    'Public Sub FormatDateOfMonth(ByVal year As Integer, ByVal month As Integer)
    '    Select Case month
    '        Case 1, 3, 5, 7, 8, 10, 12
    '            ReDim xdateofMonth(30)
    '            ReDim ydateofMonth(30)
    '        Case 4, 6, 9, 11
    '            ReDim xdateofMonth(29)
    '            ReDim ydateofMonth(29)
    '        Case 2
    '            If (((year Mod 4 = 0) And (year Mod 100 <> 0) Or (year Mod 400 = 0))) Then
    '                ReDim xdateofMonth(28)
    '                ReDim ydateofMonth(28)
    '            Else
    '                ReDim xdateofMonth(27)
    '                ReDim ydateofMonth(27)
    '            End If
    '    End Select
    '    For index = 1 To xdateofMonth.Length
    '        xdateofMonth(index - 1) = index
    '        ydateofMonth(index - 1) = 0
    '    Next
    'End Sub

    Public Sub deletefile(ByVal targetDirectory As String)
        Dim fileEntries As String() = Directory.GetFiles(targetDirectory)
        ' Process the list of files found in the directory.
        Dim fileName As String
        For Each fileName In fileEntries
            If fileName.Contains(setfilenamereport) = True Then
                File.Delete(fileName)
            End If
        Next fileName
        Dim subdirectoryEntries As String() = Directory.GetDirectories(targetDirectory)
        ' Recurse into subdirectories of this directory.
        'Dim subdirectory As String
        'For Each subdirectory In subdirectoryEntries
        '    deletefile(subdirectory)
        'Next subdirectory
    End Sub 'ProcessDirectory
    Public Sub ShiftDirectory(ByVal targetDirectory As String)



        If Directory.Exists(targetDirectory & "\" & Now.Date.ToString("yyyyMMdd")) = False Then
            Directory.CreateDirectory(targetDirectory & "\" & Now.Date.ToString("yyyyMMdd"))
        End If
        If Directory.Exists(targetDirectory & "\" & Now.Date.ToString("yyyyMMdd") & "\" & ModelCurrent) = False Then
            Directory.CreateDirectory(targetDirectory & "\" & Now.Date.ToString("yyyyMMdd") & "\" & ModelCurrent)

        End If
        If Directory.Exists(targetDirectory & "\" & Now.Date.ToString("yyyyMMdd") & "\" & ModelCurrent & "\CaNgay") = False Then
            Directory.CreateDirectory(targetDirectory & "\" & Now.Date.ToString("yyyyMMdd") & "\" & ModelCurrent & "\CaNgay")
        End If
        If Directory.Exists(targetDirectory & "\" & Now.Date.ToString("yyyyMMdd") & "\" & ModelCurrent & "\CaDem") = False Then
            Directory.CreateDirectory(targetDirectory & "\" & Now.Date.ToString("yyyyMMdd") & "\" & ModelCurrent & "\CaDem")
        End If
        Dim fileEntries As String() = Directory.GetFiles(targetDirectory & "\" & Now.Date.ToString("yyyyMMdd"))
        ' Process the list of files found in the directory.
        Dim fileName As String
        For Each fileName In fileEntries
            If fileName.Contains(setfilenamereport) = True And fileName.Contains(ModelCurrent) = True Then
                Dim filecut As String = Mid(fileName, InStrRev(fileName, "\", -1, CompareMethod.Text) + 1, fileName.Length)
                If File.GetLastWriteTime(fileName).Hour >= 8 And File.GetLastWriteTime(fileName).Hour <= 19 Then
                    File.Move(fileName, targetDirectory & "\" & Now.Date.ToString("yyyyMMdd") & "\" & ModelCurrent & "\CaNgay\" & filecut)
                Else
                    File.Move(fileName, targetDirectory & "\" & Now.Date.ToString("yyyyMMdd") & "\" & ModelCurrent & "\CaDem" & filecut)
                End If
            End If
        Next fileName
        Dim fileEntries2 As String() = Directory.GetFiles(targetDirectory & "\" & Now.Date.ToString("yyyyMMdd") & "\" & ModelCurrent & "\CaNgay")
        For Each fileName In fileEntries2
            If fileName.Contains(setfilenamereport) = True Then CountFileBeginDay += 1
        Next fileName
        Dim fileEntries3 As String() = Directory.GetFiles(targetDirectory & "\" & Now.Date.ToString("yyyyMMdd") & "\" & ModelCurrent & "\CaDem")
        For Each fileName In fileEntries2
            If fileName.Contains(setfilenamereport) = True Then CountFileBeginNight += 1
        Next fileName

    End Sub 'ProcessDirectory
    Public Sub ProcessDirectory(ByVal targetDirectory As String)


        Dim fileEntries As String() = Directory.GetFiles(targetDirectory & "\" & Now.Date.ToString("yyyyMMdd"))

        If Directory.Exists(targetDirectory & "\" & Now.Date.ToString("yyyyMMdd")) = False Then
            Directory.CreateDirectory(targetDirectory & "\" & Now.Date.ToString("yyyyMMdd"))
        End If
        If Directory.Exists(targetDirectory & "\" & Now.Date.ToString("yyyyMMdd") & "\" & ModelCurrent) = False Then
            Directory.CreateDirectory(targetDirectory & "\" & Now.Date.ToString("yyyyMMdd") & "\" & ModelCurrent)

        End If
        If Directory.Exists(targetDirectory & "\" & Now.Date.ToString("yyyyMMdd") & "\" & ModelCurrent & "\CaNgay") = False Then
            Directory.CreateDirectory(targetDirectory & "\" & Now.Date.ToString("yyyyMMdd") & "\" & ModelCurrent & "\CaNgay")
        End If
        If Directory.Exists(targetDirectory & "\" & Now.Date.ToString("yyyyMMdd") & "\" & ModelCurrent & "\CaDem") = False Then
            Directory.CreateDirectory(targetDirectory & "\" & Now.Date.ToString("yyyyMMdd") & "\" & ModelCurrent & "\CaDem")
        End If
        ' Process the list of files found in the directory.
        Dim fileName As String
        For Each fileName In fileEntries
    
            If fileName.Contains(setfilenamereport) = True And fileName.Contains(ModelCurrent) = True Then
                Dim filecut As String = Mid(fileName, InStrRev(fileName, "\", -1, CompareMethod.Text) + 1, fileName.Length)
                If File.GetLastWriteTime(fileName).Hour >= 8 And File.GetLastWriteTime(fileName).Hour <= 19 Then
                    File.Move(fileName, targetDirectory & "\" & Now.Date.ToString("yyyyMMdd") & "\" & ModelCurrent & "\CaNgay\" & filecut)
                    CountFileCurrentDay = CountFileBeginDay + 1
                    IncreaseProduct()
                Else
                    File.Move(fileName, targetDirectory & "\" & Now.Date.ToString("yyyyMMdd") & "\" & ModelCurrent & "\CaDem\" & filecut)
                    CountFileCurrentNight = CountFileBeginNight + 1
                    IncreaseProduct()
                End If
            End If
        Next fileName
    End Sub 'ProcessDirectory
    Public Sub FormatNgayCasx()
        CheckCaSX()
        If Shiftcheck = True Then
            Datecheck = Now.Date.ToString("dd-MM-yyyy")
        Else
            If Now.Hour >= 0 And Now.Hour <= 7 Then
                If Now.Day > 1 Then
                    Datecheck = (Now.Day - 1).ToString.PadLeft(2, "0"c) & "-" & Now.Month.ToString().PadLeft(2, "0"c) & "-" & Now.Year

                Else
                    Select Case Now.Month - 1
                        Case 1
                            Datecheck = Now.Year - 1 & "1231"

                        Case 3, 5, 7, 8, 10, 12
                            Datecheck = "31-" & (Now.Month - 1).ToString().PadLeft(2, "0"c) & "-" & Now.Year

                        Case 4, 6, 9, 11

                            Datecheck = "30-" & (Now.Month - 1).ToString().PadLeft(2, "0"c) & "-" & Now.Year
                        Case 2

                            If (((Now.Year Mod 4 = 0) And (Now.Year Mod 100 <> 0) Or (Now.Year Mod 400 = 0))) Then
                                Datecheck = "29-02" & Now.Year
                            Else
                                Datecheck = "28-02" & Now.Year
                            End If
                    End Select
                End If
            Else
                Datecheck = Now.Date.ToString("dd-MM-yyyy")
            End If
        End If
        TextDateProduct.Text = Datecheck
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadsetting()
    End Sub
    Public Sub SetupDisplay()
        If ComControlPort.IsOpen = True Then
            If ComControlPort.IsOpen = True Then ComControlPort.WriteLine("S+0000000000000*R")
            If ComControlPort.IsOpen = True Then ComControlPort.Write("C")
        End If
        ComboModel.Enabled = True
        ComboModel.Text = ""
        For index = 1 To 10
            Table1.Controls("TextTime" & index).Text = ""
            Table1.Controls("TextPlan" & index).Text = ""
            Table1.Controls("TextActual" & index).Text = ""
            Table1.Controls("TextBalance" & index).Text = ""
            Table1.Controls("TextComment" & index).Text = ""
        Next
        Copyright.Text = My.Application.Info.Copyright.ToString
        Me.Text = My.Application.Info.Title.ToString & "_" & My.Application.Info.Version.ToString
        TextCycleTimeCurrent.Text = 0
        TextCycleTimeModel.Text = 0
        TextShiftProduct.Clear()
        TextDateProduct.Clear()
        PauseProduct = False
        StartProduct = False
        StatusLine = 0
        CountProduct = 0
        ProductPlan = 0
        For index = 1 To 9
            CountProductPerHour(index) = 0
        Next
        NoPeople = 0
        ModelCurrent = ""
        LabelStatus.Text = "Tình trạng Line: Chọn model"
        Shape2.Visible = False
        Shape1.Visible = False
        Shape3.Visible = False
        LabelShapeOnline.Visible = False
        LabelShapeOffLine.Visible = False
        LabelShapeError.Visible = False
        Timer1.Enabled = True
        TextSerial.Enabled = False
        TextCycleTimeCurrent.Text = 0
        TextPerson.Text = 0
        TextPlan.Text = 0
        TextActual.Text = 0
        TextBalance.Text = 0
        CountProduct = 0
        LabelCountProduct.Text = CountProduct
        TextShiftProduct.Text = ""
        BtStart.Enabled = False
        BtStop.Enabled = False
        BtStart.Text = "Bắt đầu"
        BtStart.Image = Image.FromFile(PathApplication & "\Icon\play.png")
        ' If CheckServer = True Then RecordDatabase()
    End Sub
    Private Sub loadsetting()

        If CheckComControlPort() = True Then
            If CheckModelList() = True Then
                If CheckPathSetup() = True Then
                    BalanceProduction = 0
                    SetupDisplay()
                    ComControlPort.WriteLine("S+0000000000000*")
                Else
                    MessageBox.Show("File Setup Path WIP Not Found", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ' End
                End If
            Else
                MessageBox.Show("File Setup List Model Not Found", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End
            End If
        Else
            End
        End If
    End Sub
    Public Sub LoadProduction()

        Dim fileload As String = PathPassrate & "\" & ModelCurrent & "\" & Datecheck & "_" & Shiftcheck & ".txt"
        Dim FolderLoad As String = PathPassrate & "\" & ModelCurrent
        If Directory.Exists(FolderLoad) = False Then Directory.CreateDirectory(FolderLoad)
        If System.IO.File.Exists(fileload) = True Then
            ProductPlan = Val(ReadTextFile(fileload, 2))
            CountProduct = Val(ReadTextFile(fileload, 4))
            For index = 1 To 10
                CountProductPerHour(index) = Val(ReadTextFile(fileload, 2 * (index + 2)))
            Next
        Else
            FilePassrate = fileload
            For index = 1 To 10
                CountProductPerHour(index) = 0
            Next
            Dim text_passrate As New System.IO.StreamWriter(FilePassrate)
            text_passrate.WriteLine("# 1 Plan")
            text_passrate.WriteLine(ProductPlan)
            text_passrate.WriteLine("#2 Actual")
            text_passrate.WriteLine(CountProduct)
            text_passrate.WriteLine("# 3 Production In Time 1 8>9")
            text_passrate.WriteLine(0)
            text_passrate.WriteLine("# 4 Production In Time 2 9>10")
            text_passrate.WriteLine(0)
            text_passrate.WriteLine("# 5 Production In Time 3	10>11")
            text_passrate.WriteLine(0)
            text_passrate.WriteLine("# 6 Production In Time 4	11>12")
            text_passrate.WriteLine(0)
            text_passrate.WriteLine("# 7 Production In Time 5	12>13")
            text_passrate.WriteLine(0)
            text_passrate.WriteLine("# 8 Production In Time 6	13>14")
            text_passrate.WriteLine(0)
            text_passrate.WriteLine("# 9 Production In Time 7	14>15")
            text_passrate.WriteLine(0)
            text_passrate.WriteLine("# 10 Production In Time 8	15>16")
            text_passrate.WriteLine(0)
            text_passrate.WriteLine("# 11 Production In Time 9	16>17")
            text_passrate.WriteLine(0)
            text_passrate.WriteLine("# 12 Production In Time 10	17>20")
            text_passrate.WriteLine(0)
            text_passrate.Close()
        End If
        BalanceProduction = CountProduct - ProductPlan
        If (BalanceProduction) < BalanceErrorSetup Then
            StatusLine = 3
            ShowStatus(StatusLine, True)
        ElseIf (BalanceProduction) < BalanceAlarmSetup Then
            StatusLine = 2
            ShowStatus(StatusLine, True)
        Else
            StatusLine = 1
            ShowStatus(StatusLine, True)
        End If

        TextPlan.Text = ProductPlan
        TextActual.Text = CountProduct
        TextBalance.Text = BalanceProduction
        'MsgBox(Format(BalanceProduction, "0000"))

        If BalanceProduction < 0 Then
                If Math.Abs(BalanceProduction) >= 1000 Then
                    ArraySend = "S-" & Format(999, "000") & Format(CountProduct, "0000") & Format(ProductPlan, "0000") & Format(NoPeople, "00") & "*"
                Else
                    ArraySend = "S" & Format(BalanceProduction, "000") & Format(CountProduct, "0000") & Format(ProductPlan, "0000") & Format(NoPeople, "00") & "*"
                End If

            Else
                ArraySend = "S+" & Format(BalanceProduction, "000") & Format(CountProduct, "0000") & Format(ProductPlan, "0000") & Format(NoPeople, "00") & "*"

            End If

            If ComControlPort.IsOpen = True Then ComControlPort.WriteLine(ArraySend)
        ' RecordDatabase()
    End Sub


    Public Sub RecordProduction()
        If ModelCurrent <> "" Then
            Dim fileload As String = PathPassrate & "\" & ModelCurrent & "\" & Datecheck & "_" & Shiftcheck & ".txt"
            If System.IO.File.Exists(fileload) = True Then
                ReplaceLine(fileload, 2, ProductPlan)
                ReplaceLine(fileload, 4, CountProduct)

                For index = 1 To 10
                    ReplaceLine(fileload, 2 * (index + 2), CountProductPerHour(index))
                Next
            Else
                FilePassrate = fileload
                Dim text_passrate As New System.IO.StreamWriter(FilePassrate)
                text_passrate.WriteLine("# 1 Plan")
                text_passrate.WriteLine(0)
                text_passrate.WriteLine("#2 Actual")
                text_passrate.WriteLine(0)
                text_passrate.WriteLine("# 3 Production In Time 1 8>9")
                text_passrate.WriteLine(0)
                text_passrate.WriteLine("# 4 Production In Time 2 9>10")
                text_passrate.WriteLine(0)
                text_passrate.WriteLine("# 5 Production In Time 3	10>11")
                text_passrate.WriteLine(0)
                text_passrate.WriteLine("# 6 Production In Time 4	11>12")
                text_passrate.WriteLine(0)
                text_passrate.WriteLine("# 7 Production In Time 5	12>13")
                text_passrate.WriteLine(0)
                text_passrate.WriteLine("# 8 Production In Time 6	13>14")
                text_passrate.WriteLine(0)
                text_passrate.WriteLine("# 9 Production In Time 7	14>15")
                text_passrate.WriteLine(0)
                text_passrate.WriteLine("# 10 Production In Time 8	15>16")
                text_passrate.WriteLine(0)
                text_passrate.WriteLine("# 10 Production In Time 8	16>17")
                text_passrate.WriteLine(0)
                text_passrate.Close()
            End If
        End If
    End Sub

    Private Sub ComboModel_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboModel.SelectedIndexChanged
        If ComboModel.Text <> "" Then
            ModelCurrent = ComboModel.Text

            If LoadModelCurrent(ModelCurrent) = True Then
                ComboModel.Enabled = False
                TextCycleTimeModel.Text = CycleTimeModel
                TextCycleTimeCurrent.Text = ""
                TextPerson.Text = NoPeople
                'TextPlan.Text = ""
                'TextBalance.Text = ""
                FormatNgayCasx()

                BtStart.Enabled = True
                BtStop.Enabled = True
                LoadProduction()
                ShiftDirectory(PathLogWipSetup)
                LabelCountProduct.Text = CountProduct
                If CheckCaSX() = True Then
                    TextShiftProduct.Text = "Ca Ngày"
                    For index = 1 To 10
                        Table1.Controls("TextTime" & index).Text = TimeLine(1 + 2 * (index - 1)).Hour & ":" & TimeLine(1 + 2 * (index - 1)).Minute & ":" & TimeLine(1 + 2 * (index - 1)).Second & ">" & TimeLine(2 + 2 * (index - 1)).Hour & ":" & TimeLine(2 + 2 * (index - 1)).Minute & ":" & TimeLine(2 + 2 * (index - 1)).Second
                        Table1.Controls("TextPlan" & index).Text = Format(((Val(TimeLine(2 + 2 * (index - 1)).Hour - TimeLine(1 + 2 * (index - 1)).Hour)) * 3600 + (Val(TimeLine(2 + 2 * (index - 1)).Minute - TimeLine(1 + 2 * (index - 1)).Minute) * 60) + (Val(TimeLine(2 + 2 * (index - 1)).Second - TimeLine(1 + 2 * (index - 1)).Second))) / CycleTimeModel, "0")
                        If CountProductPerHour(index) > 0 Then Table1.Controls("TextActual" & index).Text = CountProductPerHour(index)
                    Next
                Else
                    TextShiftProduct.Text = "Ca Đêm"
                    For index = 1 To 10
                        Table1.Controls("TextTime" & index).Text = TimeLine(1 + 2 * (index - 1)).Hour & ":" & TimeLine(1 + 2 * (index - 1)).Minute & ":" & TimeLine(1 + 2 * (index - 1)).Second & ">" & TimeLine(2 + 2 * (index - 1)).Hour & ":" & TimeLine(2 + 2 * (index - 1)).Minute & ":" & TimeLine(2 + 2 * (index - 1)).Second
                        Table1.Controls("TextPlan" & index).Text = Format(((Val(TimeLine(2 + 2 * (index - 1)).Hour - TimeLine(1 + 2 * (index - 1)).Hour)) * 3600 + (Val(TimeLine(2 + 2 * (index - 1)).Minute - TimeLine(1 + 2 * (index - 1)).Minute) * 60) + (Val(TimeLine(2 + 2 * (index - 1)).Second - TimeLine(1 + 2 * (index - 1)).Second))) / CycleTimeModel, "0")
                        If CountProductPerHour(index) > 0 Then Table1.Controls("TextActual" & index).Text = CountProductPerHour(index)
                    Next
                End If

            End If
        End If

    End Sub
    Public Sub ShowStatus(ByVal value As Integer, ByVal VisibleShow As Boolean)

        Select Case value
            Case 0

                LabelStatus.Text = "Tình trạng Line:       "
                LabelShapeOnline.Visible = False
                LabelShapeOffLine.Visible = False
                LabelShapeError.Visible = False
                Shape1.Visible = False
                Shape2.Visible = False
                Shape3.Visible = False

            Case 1
                LabelStatus.Text = "Tình trạng Line:Online"
                LabelShapeOnline.Visible = VisibleShow
                LabelShapeOffLine.Visible = False
                LabelShapeError.Visible = False
                Shape1.Visible = VisibleShow
                If VisibleShow = False Then
                    If ComControlPort.IsOpen = True Then ComControlPort.Write("R")
                Else
                    If ComControlPort.IsOpen = True Then ComControlPort.Write("X")
                End If

                Shape2.Visible = False
                Shape3.Visible = False
            Case 2
                LabelStatus.Text = "Tình trạng Line: Bao Dong"
                LabelShapeOnline.Visible = False
                LabelShapeOffLine.Visible = VisibleShow
                LabelShapeError.Visible = False
                Shape1.Visible = False
                Shape2.Visible = VisibleShow
                If VisibleShow = False Then
                    If ComControlPort.IsOpen = True Then ComControlPort.Write("R")
                Else
                    If ComControlPort.IsOpen = True Then ComControlPort.Write("V")
                End If
                Shape3.Visible = False
            Case 3
                LabelStatus.Text = "Tình trạng Line: Bat Thuong"
                LabelShapeOnline.Visible = False
                LabelShapeOffLine.Visible = False
                LabelShapeError.Visible = VisibleShow
                Shape1.Visible = False
                Shape2.Visible = False
                Shape3.Visible = VisibleShow
                If VisibleShow = False Then
                    If ComControlPort.IsOpen = True Then ComControlPort.Write("R")
                Else
                    If ComControlPort.IsOpen = True Then ComControlPort.Write("D")
                End If
            Case Else

        End Select
    End Sub
    Private Sub BtStart_Click(sender As Object, e As EventArgs) Handles BtStart.Click
        ComboModel.Enabled = False
        LabelShapeOnline.Visible = True
        LabelShapeOffLine.Visible = False
        LabelShapeError.Visible = False
        Shape1.Visible = False
        Shape2.Visible = False
        Shape3.Visible = False
        If StartProduct = False Then
            ShowStatus(StatusLine, True)
            LabelStatus.Text = "Tình trạng Line: Online"
            PauseProduct = False
            StartProduct = True
            ' GroupBox3.Controls("Shape" & StatusLine).Visible = True
            BtStart.Text = "Tạm dừng"
            BtStart.Image = Image.FromFile(PathApplication & "\Icon\pause.png")
            Dim sumtime As Integer = ((Now.Hour * 100) + Now.Minute)
            If StatusLine = 0 Then
                StatusLine = 1
                ShowStatus(StatusLine, True)
                For index = 1 To 20
                    If index Mod 2 <> 0 And sumtime >= ((TimeLine(index).Hour) * 100 + TimeLine(index).Minute) And sumtime <= ((TimeLine(index + 1).Hour) * 100 + TimeLine(index + 1).Minute) Then

                        TimeLine(index) = New DateTime(Now.Year, Now.Month, Now.Day, Now.Hour, Now.Minute, Now.Second)
                        Table1.Controls("TextTime" & (index \ 2) + 1).Text = TimeLine(index).Hour & ":" & TimeLine(index).Minute & ":" & TimeLine(index).Second & ">" & TimeLine(index + 1).Hour & ":" & TimeLine(index + 1).Minute & ":" & TimeLine(index + 1).Second
                        Table1.Controls("TextPlan" & (index \ 2) + 1).Text = Format((((Val(TimeLine(index + 1).Hour - TimeLine(index).Hour)) * 3600 + (Val(TimeLine(index + 1).Minute - TimeLine(index).Minute) * 60) + (Val(TimeLine(index + 1).Second - TimeLine(index).Second))) / CycleTimeModel) + CountProductPerHour((index \ 2) + 1), "0")
                        TextPlan.Text = Table1.Controls("TextPlan" & (index \ 2) + 1).Text
                        TextActual.Text = CountProduct
                        Exit For
                    End If
                    'If TimeLine(index).Hour = Now.Hour Then
                    '    If Now.Minute >= TimeLine(index).Minute And Now.Minute <= TimeLine(index + 1).Minute Then ' neu nam trong khoang thoi gian nao do
                    '        TimeLine(index) = New DateTime(Now.Year, Now.Month, Now.Day, Now.Hour, Now.Minute, Now.Second)
                    '        Table1.Controls("TextTime" & (index \ 2) + 1).Text = TimeLine(index).Hour & ":" & TimeLine(index).Minute & ":" & TimeLine(index).Second & ">" & TimeLine(index + 1).Hour & ":" & TimeLine(index + 1).Minute & ":" & TimeLine(index + 1).Second
                    '        Table1.Controls("TextPlan" & (index \ 2) + 1).Text = Format((((Val(TimeLine(index + 1).Hour - TimeLine(index).Hour)) * 3600 + (Val(TimeLine(index + 1).Minute - TimeLine(index).Minute) * 60) + (Val(TimeLine(index + 1).Second - TimeLine(index).Second))) / CycleTimeModel) + CountProductPerHour((index \ 2) + 1), "0")
                    '        'TextPlan.Text = Table1.Controls("TextPlan" & (index \ 2) + 1).Text
                    '        TextActual.Text = CountProduct
                    '        TimeUse((index \ 2) + 1) = True
                    '        Exit For
                    '    End If
                    'End If
                Next
                If BarcodeEnable = True Then TextSerial.Focus()
            Else
                LabelStatus.Text = "Tình trạng Line: Online"
                PauseProduct = False
                StartProduct = True
                ShowStatus(StatusLine, True)
                For index = 1 To 20
                    If index Mod 2 <> 0 And sumtime >= ((TimeLine(index).Hour) * 100 + TimeLine(index).Minute) And sumtime <= ((TimeLine(index + 1).Hour) * 100 + TimeLine(index + 1).Minute) Then
                        Table1.Controls("TextTime" & (index \ 2) + 1).Text = TimeLine(index).Hour & ":" & TimeLine(index).Minute & ":" & TimeLine(index).Second & ">" & TimeLine(index + 1).Hour & ":" & TimeLine(index + 1).Minute & ":" & TimeLine(index + 1).Second
                        Table1.Controls("TextPlan" & (index \ 2) + 1).Text = Format((((Val(TimeLine(index + 1).Hour - TimeLine(index).Hour)) * 3600 + (Val(TimeLine(index + 1).Minute - TimeLine(index).Minute) * 60) + (Val(TimeLine(index + 1).Second - TimeLine(index).Second))) / CycleTimeModel) + CountProductPerHour((index \ 2) + 1), "0")
                        TextPlan.Text = Table1.Controls("TextPlan" & (index \ 2) + 1).Text
                        Exit For
                    End If

                Next

                TimePauseLine = 0
            End If
        Else
            BtStart.Text = "Bắt đầu"
            BtStart.Image = Image.FromFile(PathApplication & "\Icon\play.png")
            LabelStatus.Text = "Tình trạng Line: Tạm dừng"
            PauseProduct = True
            StartProduct = False
        End If

    End Sub

    Private Sub BtStop_Click(sender As Object, e As EventArgs) Handles BtStop.Click

        SetupDisplay()
    End Sub

    'Private Sub TimerPress_Tick(sender As Object, e As EventArgs) Handles TimerPress.Tick
    '    If StartProduct = True Then
    '        ComPressPort.Write("B")
    '        If ComPressPort.ReadExisting() = "B" Then
    '            BitPress = True
    '        ElseIf ComPressPort.ReadExisting() <> "B" And BitPress = True Then
    '            BitPress = False
    '            IncreaseProduct()
    '        End If
    '    End If
    'End Sub
    Public Sub IncreaseProduct()
        If StartProduct = True Then
            Dim sumtime As Integer = ((Now.Hour * 100) + Now.Minute)
            For index = 1 To 20
                If index Mod 2 <> 0 And sumtime >= ((TimeLine(index).Hour) * 100 + TimeLine(index).Minute) And sumtime <= ((TimeLine(index + 1).Hour) * 100 + TimeLine(index + 1).Minute) Then
                    CountProduct = CountProduct + 1
                    If CountProduct = 1 Then
                        CycleTimeActual = TimeCycleActual
                    Else
                        CycleTimeActual = Format((TimeCycleActual + CycleTimeActual) / 2, "0")
                    End If
                    TextCycleTimeCurrent.Text = CycleTimeActual
                    TimeCycleActual = 0
                    LabelCountProduct.Text = CountProduct
                    CountProductPerHour((index \ 2) + 1) = CountProductPerHour((index \ 2) + 1) + 1
                    Table1.Controls("TextActual" & (index \ 2) + 1).Text = CountProductPerHour((index \ 2) + 1)
                    Table1.Controls("TextBalance" & (index \ 2) + 1).Text = CountProductPerHour((index \ 2) + 1) - Val(Table1.Controls("TextPlan" & (index \ 2) + 1).Text)
                    Exit For
                End If
            Next
            RecordProduction()
        End If
    End Sub
    Public Sub ReduceProduct()
        If StartProduct = True And CountProduct > 0 Then
            Dim sumtime As Integer = ((Now.Hour * 100) + Now.Minute)
            For index = 1 To 20
                If index Mod 2 <> 0 And sumtime >= ((TimeLine(index).Hour) * 100 + TimeLine(index).Minute) And sumtime <= ((TimeLine(index + 1).Hour) * 100 + TimeLine(index + 1).Minute) Then
                    CountProduct = CountProduct - 1
                    LabelCountProduct.Text = CountProduct
                    CountProductPerHour((index \ 2) + 1) = CountProductPerHour((index \ 2) + 1) - 1
                    Table1.Controls("TextActual" & (index \ 2) + 1).Text = CountProductPerHour((index \ 2) + 1)
                    Table1.Controls("TextBalance" & (index \ 2) + 1).Text = CountProductPerHour((index \ 2) + 1) - Val(Table1.Controls("TextPlan" & (index \ 2) + 1).Text)
                    Exit For
                End If
            Next
            RecordProduction()
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        LabelTimeDate.Text = Now.ToString("HH:mm:ss _ dd/MM/yyyy")
        If (Now.Hour = 8 And Now.Minute = 0 And Now.Second = 10) Or (Now.Hour = 20 And Now.Minute = 0 And Now.Second = 10) Then
            BtStop_Click(Nothing, Nothing)
        End If
        If StartProduct = True Then
            ProcessDirectory(PathLogWipSetup)
            TimerPress.Enabled = True
            Dim sumtime As Integer = ((Now.Hour * 100) + Now.Minute)
            For index = 1 To 20
                If index Mod 2 <> 0 And sumtime >= ((TimeLine(index).Hour) * 100 + TimeLine(index).Minute) And sumtime <= ((TimeLine(index + 1).Hour) * 100 + TimeLine(index + 1).Minute) Then
                    TimeCountPlan = TimeCountPlan + 1
                    TimeCycleActual = TimeCycleActual + 1
                    If TimeCountPlan >= CycleTimeModel * 10 Then
                        TimeCountPlan = 0
                        ProductPlan = ProductPlan + 1
                    End If
                    Exit For
                End If
            Next
        Else
            TimeCountPlan = 0
            If PauseProduct = True Then
                TimePauseLine = TimePauseLine + 1
                If TimePauseLine Mod 2 Then
                    ShowStatus(StatusLine, True)
                Else
                    ShowStatus(StatusLine, False)
                End If
            End If
        End If
        If BarcodeEnable = True Then TextSerial.Focus()
        BalanceProduction = CountProduct - ProductPlan
        If StartProduct = True Then
            If (BalanceProduction) < BalanceErrorSetup Then
                StatusLine = 3
                ShowStatus(StatusLine, True)
            ElseIf (BalanceProduction) < BalanceAlarmSetup Then
                StatusLine = 2
                ShowStatus(StatusLine, True)
            Else
                StatusLine = 1
                ShowStatus(StatusLine, True)
            End If
        End If
        TextPlan.Text = ProductPlan
        TextActual.Text = CountProduct
        TextBalance.Text = BalanceProduction
        'MsgBox(Format(BalanceProduction, "0000"))
        If StartProduct = True Then
            If BalanceProduction < 0 Then
                If Math.Abs(BalanceProduction) >= 1000 Then
                    ArraySend = "S-" & Format(999, "000") & Format(CountProduct, "0000") & Format(ProductPlan, "0000") & Format(NoPeople, "00") & "*"
                Else
                    ArraySend = "S" & Format(BalanceProduction, "000") & Format(CountProduct, "0000") & Format(ProductPlan, "0000") & Format(NoPeople, "00") & "*"
                End If

            Else
                ArraySend = "S+" & Format(BalanceProduction, "000") & Format(CountProduct, "0000") & Format(ProductPlan, "0000") & Format(NoPeople, "00") & "*"

            End If

            If ComControlPort.IsOpen = True Then ComControlPort.WriteLine(ArraySend)
            ' RecordDatabase()
        End If
    End Sub

    Private Sub BtIncrease_Click(sender As Object, e As EventArgs) Handles BtIncrease.Click
        IncreaseProduct()
    End Sub

    Private Sub BtReduce_Click(sender As Object, e As EventArgs) Handles BtReduce.Click
        ReduceProduct()
    End Sub

    Private Sub Control_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        RecordProduction()
        SetupDisplay()


    End Sub

    Private Sub TextSerial_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextSerial.KeyPress
        If Asc(e.KeyChar) = 13 Then
            TextSerial.Text.TrimStart.TrimEnd()
            If TextSerial.TextLength = IdCodeLenght Then
                If Mid(TextSerial.Text, ModelRevPosition, ModelRev.Length) = ModelRev Then
                    If StartProduct = True Then
                        TextSerial.Clear()
                        IncreaseProduct()
                    End If
                Else
                    NG_FORM.Show()
                    NG_FORM.Lb_inform_NG.Text = "Sai ma quy dinh model:" & ModelRev
                End If
            Else
                NG_FORM.Show()
                NG_FORM.Lb_inform_NG.Text = "Khong dung do dai label:" & IdCodeLenght
            End If
        End If
    End Sub
End Class
