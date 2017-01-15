


Public Class Form1
    Dim hotkey As Integer = 123
    Dim msdown As Boolean = False

    Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vkey As Integer) As Short
    Dim proc_num As Integer
    Dim proc_ary(10000) As String

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim judge As Boolean = False
        Dim pList() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcesses

        Debug.Print(GetAsyncKeyState(hotkey))
        judge = GetAsyncKeyState(hotkey)

        If judge = True Then
            For Each proc As System.Diagnostics.Process In pList
                For i = 0 To proc_num - 1
                    If proc.ProcessName = proc_ary(i) Then
                        proc.Kill()
                    End If
                Next
            Next
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim proc_name As String
        Dim judge As Boolean = False

        proc_name = InputBox("請輸入程序名稱(不用加上.exe)", "程序輸入")

        If proc_name = "" Then
            Exit Sub
        End If

        For i = 0 To proc_num - 1
            If proc_name = proc_ary(i) Then
                MsgBox("請勿輸入重複程序", MsgBoxStyle.Exclamation, "錯誤")
                judge = True
                Exit Sub
            End If
        Next

        If judge = False Then
            proc_ary(proc_num) = proc_name
            proc_num += 1
            If proc_num = 1 Then
                TextBox1.AppendText(proc_name)
            Else
                TextBox1.AppendText(vbNewLine & proc_name)
            End If
            Timer1.Enabled = True
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        proc_ary(0) = "GTA5"
        proc_num = 1
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        If proc_num > 1 Then
            TextBox1.Text = Strings.Left(TextBox1.Text, Strings.Len(TextBox1.Text) - Strings.Len(vbNewLine) - Strings.Len(proc_ary(proc_num - 1)))
            proc_num -= 1
        ElseIf proc_num = 1 Then
            TextBox1.Text = ""
            proc_num = 0
            Timer1.Enabled = False
        End If
        Debug.Print(proc_num)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        proc_num = 0
        TextBox1.Text = ""
        Timer1.Enabled = False
    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown
        e.SuppressKeyPress = True
        If msdown = True Then
            Debug.Print(e.KeyCode.ToString())
            TextBox2.Text = "按一下變更(目前：" & e.KeyCode.ToString & ")"
            TextBox2.ForeColor = Color.Black
            msdown = False
            hotkey = e.KeyCode


        End If
    End Sub

    Private Sub TextBox2_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyUp
        Timer1.Enabled = True
    End Sub


    Private Sub TextBox2_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TextBox2.MouseClick
        TextBox2.ForeColor = Color.Red
        TextBox2.Text = "請按下關閉程式熱鍵"
        Timer1.Enabled = False
        msdown = True
    End Sub
End Class


