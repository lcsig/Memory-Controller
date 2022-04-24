Imports System.Runtime.InteropServices

Public Class Form1

    <DllImport("kernel32.dll", SetLastError:=True)> _
    Public Shared Function ReadProcessMemory(ByVal hProcess As IntPtr, _
                                             ByVal lpBaseAddress As IntPtr, _
                                             <Out()> ByVal lpBuffer() As Byte, _
                                             ByVal dwSize As Integer, _
                                             ByRef lpNumberOfBytesRead As Integer) As Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)> _
    Public Shared Function WriteProcessMemory(ByVal hProcess As IntPtr,
                                              ByVal lpBaseAddress As IntPtr,
                                              ByVal lpBuffer As Byte(),
                                              ByVal nSize As System.UInt32,
                                              <Out()> ByRef lpNumberOfBytesWritten As IntPtr) As Boolean
    End Function

    <DllImport("kernel32.dll", EntryPoint:="OpenProcess", SetLastError:=True)> _
    Private Shared Function OpenProcess(ByVal dwDesiredAccess As UInteger, <MarshalAsAttribute(UnmanagedType.Bool)>
                                        ByVal bInheritHandle As Boolean, ByVal dwProcessId As UInteger) As IntPtr
    End Function

    Declare Function CloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Integer) As Integer

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If ComboBox2.SelectedIndex = -1 Then
                MessageBox.Show("Choose your Type of Value", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim x As Int32 = Process.GetProcessesByName(TextBox1.Text)(0).Handle
            Dim Y As Byte()
            Dim VOfcomb As Integer = ComboBox2.SelectedIndex

            If VOfcomb = 0 Then ' string
                Y = System.Text.Encoding.Default.GetBytes(TextBox3.Text)
            ElseIf VOfcomb = 1 Then ' string with unicode
                Y = System.Text.Encoding.Unicode.GetBytes(TextBox3.Text)
            ElseIf VOfcomb = 2 Then ' Float
                Dim Sing As Single = CType(TextBox3.Text, Single)
                Y = BitConverter.GetBytes(Sing)
            ElseIf VOfcomb = 3 Then ' Double 
                Dim Doub As Double = CType(TextBox3.Text, Double)
                Y = BitConverter.GetBytes(Doub)
            ElseIf VOfcomb = 4 Then ' Integer  
                Dim Int As Integer = CType(TextBox3.Text, Integer)
                Y = BitConverter.GetBytes(Int)
            End If

            If x = 0 Then
                MessageBox.Show("Not exist or the process is protect", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If WriteProcessMemory(x, TextBox2.Text, Y, Y.Length, 0) Then
                    MessageBox.Show("Wrote successfully", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("Wrote Faild or the process is protect", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try

            If ComboBox1.SelectedIndex = -1 Then
                MessageBox.Show("Choose your Type of Value", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            Dim x As Int32 = Process.GetProcessesByName(TextBox6.Text)(0).Handle
            Dim VOfcomb As Integer = ComboBox1.SelectedIndex
            Dim Batee(1024) As Byte
            Dim INT As Integer
            Dim Sing As Single
            Dim Doub As Double

            If x = 0 Then
                MessageBox.Show("Not exist or the process is protect", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                If ReadProcessMemory(x, TextBox5.Text, Batee, 1024, 0) Then

                    If VOfcomb = 0 Then ' string
                        TextBox4.Text = System.Text.Encoding.Default.GetString(Batee)
                    ElseIf VOfcomb = 1 Then ' string with unicode
                        TextBox4.Text = System.Text.Encoding.Unicode.GetString(Batee)
                    ElseIf VOfcomb = 2 Then ' Float
                        Sing = BitConverter.ToSingle(Batee, 0)
                        TextBox4.Text = Sing
                    ElseIf VOfcomb = 3 Then ' Double 
                        Doub = BitConverter.ToDouble(Batee, 0)
                        TextBox4.Text = Doub
                    ElseIf VOfcomb = 4 Then ' Integer  
                        INT = BitConverter.ToInt32(Batee, 0)
                        TextBox4.Text = INT
                    End If

                Else
                    MessageBox.Show("Read Faild or the process is protect", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub

    Public n
    Const ALL_ACCESS = &H38
    Public HWND As Integer

    Public ThreadD As Threading.Thread

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try
            If Button5.Text = "Get Process" Then
                If Process.GetProcessesByName(TextBox9.Text).Length <= 0 Then
                    MessageBox.Show("Not exist or the process is protect", "unfortunately", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                HWND = OpenProcess(ALL_ACCESS, False, Process.GetProcessesByName(TextBox9.Text)(0).Id)
                If HWND = 0 Then
                    MessageBox.Show("Not exist or the process is protect", "unfortunately", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Button5.Text = "Close Handle"
            ElseIf Button5.Text = "Close Handle" Then
                CloseHandle(HWND)
                HWND = 0
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub



    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Try

            Control.CheckForIllegalCrossThreadCalls = False

            If ComboBox3.SelectedIndex = -1 Then
                MessageBox.Show("Choose your Type of Value", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            ElseIf TextBox11.Text = "&H" Or TextBox10.Text = "" Then
                MessageBox.Show("type your information", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            ElseIf HWND = 0 Then
                MessageBox.Show("Not exist or the process is protect", "unfortunately", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If



            If Button4.Text = "Search" Then
                baseAddress = CType(TextBox11.Text, Int32)
                If ComboBox3.SelectedIndex = 0 Then string32Object = TextBox10.Text
                If ComboBox3.SelectedIndex = 1 Then string32Object = TextBox10.Text
                If ComboBox3.SelectedIndex = 2 Then single32Object = CType(TextBox10.Text, Single)
                If ComboBox3.SelectedIndex = 3 Then Double32Object = CType(TextBox10.Text, Double)
                If ComboBox3.SelectedIndex = 4 Then int32Object = CType(TextBox10.Text, Integer)
                If CheckBox2.Checked = True Then PartW = True Else PartW = False
                CheckBox2.Enabled = False
                ProgressBar1.Value = 0
                VOfcomb = ComboBox3.SelectedIndex
                ThreadD = New Threading.Thread(AddressOf BackgroundWorker1)
                ThreadD.Start()
                ListView1.Items.Clear()
                Button4.Text = "Stop"
                Button5.Enabled = False
                TextBox9.Enabled = False
                TextBox10.Enabled = False
                Button3.Enabled = False
                ComboBox3.Enabled = False
            ElseIf Button4.Text = "Stop" Then
                ThreadD.Abort()
                Exit Sub
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub



    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try

            If ComboBox3.SelectedIndex = -1 Then
                MessageBox.Show("Choose your Type of Value", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            ElseIf TextBox11.Text = "&H" Or TextBox10.Text = "" Then
                MessageBox.Show("type your information", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            ElseIf HWND = 0 Then
                MessageBox.Show("Not exist or the process is protect", "unfortunately", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Try
                Button3.Enabled = False
                Button5.Enabled = False
                TextBox9.Enabled = False
                TextBox10.Enabled = False
                ComboBox3.Enabled = False
                Button4.Enabled = False
                CheckBox2.Enabled = False
                If ListView1.Items.Count > 0 Then

                    Dim Y(1024) As Byte
                    Dim VOfcomb As Integer = ComboBox3.SelectedIndex
                    Dim NumberOfOrginal As Integer = ListView1.Items.Count

                    For i = 0 To NumberOfOrginal - 1

                        Application.DoEvents()
                        Dim Add As Integer = CType(ListView1.Items(i).Text, Integer)
                        If ReadProcessMemory(HWND, Add, Y, 1024, 0) Then
                            Dim INT As Integer
                            Dim Sing As Single
                            Dim Doub As Double
                            Dim str As String

                            If VOfcomb = 0 Then ' string
                                str = System.Text.Encoding.Default.GetString(Y)

                                If str.StartsWith(TextBox10.Text) Then
                                    n = ListView1.Items.Add(ListView1.Items(i).Text)
                                    n.SubItems.Add(str)
                                Else
                                End If
                            ElseIf VOfcomb = 1 Then ' string with unicode
                                str = System.Text.Encoding.Unicode.GetString(Y)

                                If str.StartsWith(TextBox10.Text) Then
                                    n = ListView1.Items.Add(ListView1.Items(i).Text)
                                    n.SubItems.Add(str)
                                Else
                                End If
                            ElseIf VOfcomb = 2 Then ' Float
                                Sing = BitConverter.ToSingle(Y, 0)

                                If Sing = CType(TextBox10.Text, Single) Then
                                    n = ListView1.Items.Add(ListView1.Items(i).Text)
                                    n.SubItems.Add(Sing)
                                Else
                                End If
                            ElseIf VOfcomb = 3 Then ' Double 
                                Doub = BitConverter.ToDouble(Y, 0)

                                If Doub = CType(TextBox10.Text, Double) Then
                                    n = ListView1.Items.Add(ListView1.Items(i).Text)
                                    n.SubItems.Add(Doub)
                                Else
                                End If
                            ElseIf VOfcomb = 4 Then ' Integer  
                                INT = BitConverter.ToInt32(Y, 0)

                                If INT = CType(TextBox10.Text, Integer) Then
                                    n = ListView1.Items.Add(ListView1.Items(i).Text)
                                    n.SubItems.Add(INT)
                                Else
                                End If
                            End If

                        Else
                            n = ListView1.Items.Add(ListView1.Items(i).Text)
                            n.SubItems.Add("Error Or nothing")
                        End If

                    Next
                    If MessageBox.Show("Remove Orginal", "", MessageBoxButtons.YesNo, MessageBoxIcon.Information) = Windows.Forms.DialogResult.Yes Then
                        For jkn As Integer = 0 To NumberOfOrginal - 1
                            ListView1.Items.RemoveAt(0)
                            Application.DoEvents()
                        Next
                    Else
                    End If

                    MessageBox.Show("Finshed", "Ok", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    GoTo m
                Else
                    MessageBox.Show("There is No items To DO Next Scan", "!!!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    GoTo m
                End If

            Catch ex As Exception
                CloseHandle(HWND)
                HWND = 0
                MsgBox(ex.Message)
                GoTo m
            End Try

            Exit Sub
M:
            Button3.Enabled = True
            Button5.Enabled = True
            TextBox9.Enabled = True
            TextBox10.Enabled = True
            ComboBox3.Enabled = True
            Button4.Enabled = True
            CheckBox2.Enabled = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub


    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        If ListView1.SelectedItems.Count > 0 Then
            'Clipboard.SetText(ListView1.SelectedItems(0).SubItems(0).Text)
            TextBox2.Text = ListView1.SelectedItems(0).SubItems(0).Text
            TextBox1.Text = TextBox9.Text
            TextBox6.Text = TextBox9.Text
            TextBox5.Text = ListView1.SelectedItems(0).SubItems(0).Text
            ComboBox1.SelectedIndex = ComboBox3.SelectedIndex
            ComboBox2.SelectedIndex = ComboBox3.SelectedIndex
        End If
    End Sub


    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If HWND = 0 Then
            Label9.Text = "Status : UnHandled"
            Label9.ForeColor = Color.Red
            Button5.Text = "Get Process"
        Else
            Label9.Text = "Status : Handled   |   " + CType(HWND, String)
            Label9.ForeColor = Color.Green
            Button5.Text = "Close Handle"
        End If
    End Sub

    Public int32Object As Integer
    Public Double32Object As Double
    Public single32Object As Single
    Public string32Object As String
    Public PartW As Boolean
    Public lastAddress As Integer = &H7FFFFFFF
    Public baseAddress As Integer
    Public ReadStackSize As Integer = 20480
    Public VOfcomb As Integer

    Private Sub BackgroundWorker1()
        Try
            Control.CheckForIllegalCrossThreadCalls = False
            Dim arraysDifference As Integer = 4 - 1
            If VOfcomb = 0 Then arraysDifference = string32Object.Length '        Assci
            If VOfcomb = 1 Then arraysDifference = string32Object.Length * 2 '    Unicode
            If VOfcomb = 2 Then arraysDifference = 4 - 1 '        single
            If VOfcomb = 3 Then arraysDifference = 8 - 1 '        Double
            Dim array As Byte()
            Dim memorySize As Int64 = Val(lastAddress) - Val(baseAddress)
            Dim accsess As UInteger = &H38
            If memorySize >= ReadStackSize Then
                Dim loopsCount As Integer = memorySize / ReadStackSize
                Dim outOfBounds As Integer = memorySize Mod ReadStackSize
                Dim currentAddress As Int64 = Val(baseAddress)
                Dim bytesReadSize As Integer
                Dim bytesToRead As Integer = ReadStackSize
                Dim progress As Integer
                For i As Integer = 0 To loopsCount - 2
                    array = New [Byte](bytesToRead - 1) {}
                    ReadProcessMemory(HWND, currentAddress, array, CUInt(bytesToRead), bytesReadSize)
                    progress = CInt(Math.Truncate((CDbl(currentAddress - CInt(baseAddress)) / CDbl(memorySize)) * 100.0))
                    ProgressBar1.Value = progress
                    If bytesReadSize > 0 Then
                        For j As Integer = 0 To array.Length - arraysDifference - 1

                            If PartW = True Then
                                GetValuePart(currentAddress, array, j)
                            Else
                                GetValue(currentAddress, array, j)
                            End If

                        Next
                    End If
                    currentAddress += Val(array.Length) - Val(arraysDifference)
                    bytesToRead = ReadStackSize + arraysDifference

                Next
                If outOfBounds > 0 Then
                    Dim outOfBoundsBytes As Byte() = New Byte((CInt(lastAddress) - currentAddress) - 1) {}
                    ReadProcessMemory(HWND, currentAddress, outOfBoundsBytes, CUInt(CInt(lastAddress) - currentAddress), bytesReadSize)
                    If bytesReadSize > 0 Then
                        For j As Integer = 0 To outOfBoundsBytes.Length - arraysDifference - 1
                            If PartW = True Then
                                GetValuePart(currentAddress, outOfBoundsBytes, j)
                            Else
                                GetValue(currentAddress, outOfBoundsBytes, j)
                            End If
                        Next
                    End If
                End If
            Else
                Dim blockSize As Integer = memorySize Mod ReadStackSize
                Dim currentAddress As Integer = Val(baseAddress)
                Dim bytesReadSize As Integer
                ReadProcessMemory(HWND, currentAddress, array, blockSize, bytesReadSize)
                If bytesReadSize > 0 Then
                    For j As Integer = 0 To array.Length - arraysDifference - 1

                        If PartW = True Then
                            GetValuePart(currentAddress, array, j)
                        Else
                            GetValue(currentAddress, array, j)
                        End If
                    Next
                End If
            End If

            ProgressBar1.Value = ProgressBar1.Maximum
            If Button4.Text = "Search" Then
                Button4.Text = "Stop"
                Button5.Enabled = False
                TextBox9.Enabled = False
                TextBox10.Enabled = False
                ComboBox3.Enabled = False
                Button3.Enabled = False
            ElseIf Button4.Text = "Stop" Then
                Button4.Text = "Search"
                Button5.Enabled = True
                TextBox9.Enabled = True
                TextBox10.Enabled = True
                ComboBox3.Enabled = True
                Button3.Enabled = True

            End If

            MessageBox.Show("Finshed", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
            GoTo m


        Catch ex As Exception

            If Button4.Text = "Search" Then
                Button4.Text = "Stop"
                Button5.Enabled = False
                TextBox9.Enabled = False
                TextBox10.Enabled = False
                ComboBox3.Enabled = False
                Button3.Enabled = False
            ElseIf Button4.Text = "Stop" Then
                Button4.Text = "Search"
                Button5.Enabled = True
                TextBox9.Enabled = True
                TextBox10.Enabled = True
                ComboBox3.Enabled = True
                Button3.Enabled = True
                CheckBox2.Enabled = True
            End If
            CloseHandle(HWND)
            HWND = 0

            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
            'TheardD.Abort()
        End Try
m:



    End Sub

    Public Function GetValue(ByVal currentAddress As Long, ByVal array As Byte(), ByVal j As Integer)

        If VOfcomb = 0 Then ' string
            Dim k As String = System.Text.Encoding.Default.GetString(array, j, Val(string32Object.Length))
            If k = string32Object Then
                n = ListView1.Items.Add(j + CInt(currentAddress))
                n.SubItems.Add(k)
            End If
        ElseIf VOfcomb = 1 Then ' string with unicode

            Dim k As String = System.Text.Encoding.Unicode.GetString(array, j, Val(string32Object.Length * 2))
            If k = string32Object Then
                n = ListView1.Items.Add(j + CInt(currentAddress))
                n.SubItems.Add(k)
            End If

        ElseIf VOfcomb = 2 Then ' Float
            Dim k As Single = BitConverter.ToSingle(array, j)
            If k = single32Object Then
                n = ListView1.Items.Add(j + CInt(currentAddress))
                n.SubItems.Add(k)
            End If
        ElseIf VOfcomb = 3 Then ' Double 

            Dim k As Double = BitConverter.ToDouble(array, j)
            If k = Double32Object Then
                n = ListView1.Items.Add(j + CInt(currentAddress))
                n.SubItems.Add(k)
            End If
        ElseIf VOfcomb = 4 Then ' Integer  

            Dim k As Integer = BitConverter.ToInt32(array, j)
            If k = int32Object Then
                n = ListView1.Items.Add(j + CInt(currentAddress))
                n.SubItems.Add(k)
            End If
        End If
        Return Nothing
    End Function

    Public Function GetValuePart(ByVal currentAddress As Long, ByVal array As Byte(), ByVal j As Integer)

        If VOfcomb = 0 Then ' string
            Dim k As String = System.Text.Encoding.Default.GetString(array, j, Val(string32Object.Length))
            If k.ToString.StartsWith(string32Object) Then
                n = ListView1.Items.Add(j + CInt(currentAddress))
                n.SubItems.Add(k)
            End If
        ElseIf VOfcomb = 1 Then ' string with unicode

            Dim k As String = System.Text.Encoding.Unicode.GetString(array, j, Val(string32Object.Length * 2))
            If k.ToString.StartsWith(string32Object) Then
                n = ListView1.Items.Add(j + CInt(currentAddress))
                n.SubItems.Add(k)
            End If

        ElseIf VOfcomb = 2 Then ' Float
            Dim k As Single = BitConverter.ToSingle(array, j)

            If k.ToString.StartsWith(single32Object) Then
                n = ListView1.Items.Add(j + CInt(currentAddress))
                n.SubItems.Add(k)
            End If
        ElseIf VOfcomb = 3 Then ' Double 

            Dim k As Double = BitConverter.ToDouble(array, j)
            If k.ToString.StartsWith(Double32Object) Then
                n = ListView1.Items.Add(j + CInt(currentAddress))
                n.SubItems.Add(k)
            End If
        ElseIf VOfcomb = 4 Then ' Integer  

            Dim k As Integer = BitConverter.ToInt32(array, j)
            If k.ToString.StartsWith(int32Object) Then
                n = ListView1.Items.Add(j + CInt(currentAddress))
                n.SubItems.Add(k)
            End If
        End If
        Return Nothing
    End Function

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        TextBox6.Text = TextBox1.Text
        TextBox5.Text = TextBox2.Text
        ComboBox1.SelectedIndex = ComboBox2.SelectedIndex
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        TextBox1.Text = TextBox6.Text
        TextBox2.Text = TextBox5.Text
        ComboBox2.SelectedIndex = ComboBox1.SelectedIndex
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        ThreadD.Abort()
        End
    End Sub


    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        Label7.Text = Val(TextBox11.Text)
    End Sub

End Class
