
';--------+---------+---------+---------+---------+---------+---------+---------+
'  MemoryMapEditor_VB_BLA21-12R3-C01 Version 1.00.1 Programmed by Futaba Corporation 2024
'	Current Platform - windows10
'	Programmed in Visual Basic 2017 Express by Microsoft
';--------+---------+---------+---------+---------+---------+---------+---------+

'履歴
'2024/11 MemoryMapEditor_VB_BLA21-12R3-C01 ver.1.00.1公開

'著作権／免責事項／サポートについて
'● 著作権
'本ソフトウェアの著作権は双葉電子工業株式会社に帰属します。
'Microsoft、Net Framework、Visual Basic 2017 Express は、
'米国 Microsoft Corporation の米国およびその他の国における登録商標または商標です。

'● 配布・免責
'営利・非営利、添付・単独を問わず配布は自由ですが、ダウンロードサイトなどの転載などの際には、
'ファイル内容に十分注意をして下さい。ただし、改造や改変したサンプルのソースを公開や配布をする
'場合は、著作権は弊社にあることと改変したことを明記して下さい。
'本ソフトウェアの使用により生じる如何なる損害に対してもその法的根拠に関わらず弊社は責任を負
'いません。これに同意した上でソフトウェアをご利用下さい。

'● サポート
'本サンプルの障害報告やご質問などは以下のお問い合わせ先でお受けしていますが、サポートできない
'場合もありますのでご了承下さい。プログラム言語についてのサポートは致しかねますのでご遠慮下さい。
'お問い合わせ先:
'〒299-4395 千葉県長生郡長生村薮塚 1080
'ロボティクスソリューション事業センター 営業部
'TEL 0475(32)6111(代) FAX 0475(32)2915
'双葉電子工業㈱のホームページ:
'https://www.futaba.co.jp/support/contact




Public Class Form1
    Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Long

    'COM Port変更
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        'COMポートが開いている場合はCOMポート変更不可
        If SerialPort1.IsOpen = False Then
            SerialPort1.PortName = ComboBox1.Text
        End If
    End Sub

    'BaudRate変更
    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged, ComboBox3.SelectedIndexChanged, ComboBox4.SelectedIndexChanged, ComboBox5.SelectedIndexChanged, ComboBox7.SelectedIndexChanged
        If SerialPort1.IsOpen = False Then
            SerialPort1.BaudRate = CInt(ComboBox2.Text)
        End If

    End Sub

    'ComboBox2.Enabled = True
    'GroupBox5.Enabled = False

    '全ID一括指令（ID=255）
    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        'ID=255の場合はデータ取得不可（Ackは可）
        If CheckBox2.Checked = True Then
            NumericUpDown1.Maximum = 255
            NumericUpDown1.Value = 255
            NumericUpDown1.Enabled = False
            Button1.Enabled = False
            Button24.Enabled = True
            Button25.Enabled = True
        Else
            NumericUpDown1.Maximum = 127
            NumericUpDown1.Value = 1
            NumericUpDown1.Enabled = True
            Button1.Enabled = True
            Button24.Enabled = False
            Button25.Enabled = False
        End If
    End Sub

    'フラッシュROM書き込みボタン
    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        If ComboBox4.Text = "OFF" Then
            'フラッシュROM書き込み中は他の操作禁止
            Me.Enabled = False
            Call WriteFlashROM()
            Me.Enabled = True
        Else
            MessageBox.Show("トルクOFFになっていません。ROM書き込みの際にはトルクOFFしてください")
            TextBox3.Text = "TRQ Error"
            GroupBox4.Enabled = True
        End If
    End Sub

    'サーボ初期化ボタン
    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        GroupBox2.Enabled = False

        '初期化パケット送信
        Call InitializeSx()

        'IDを1に直す（ID=255の場合除く)
        If CheckBox2.Checked = False Then
            NumericUpDown1.Value = 1
            NumericUpDown2.Value = 1
        End If

    End Sub

    'searchID/BaudRateボタン
    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        Dim i, pID, Ack, BaudRate(13) As Integer

        '全ID一括指令解除
        CheckBox2.Checked = False

        'search実行中、ボタン無効/パラメータ編集無効
        TextBox3.Text = "searching..."
        GroupBox2.Enabled = False
        GroupBox4.Enabled = False

        'BaudRate表
        BaudRate(0) = 9600
        BaudRate(1) = 14400
        BaudRate(2) = 19200
        BaudRate(3) = 28800
        BaudRate(4) = 38400
        BaudRate(5) = 57600
        BaudRate(6) = 76800
        BaudRate(7) = 115200
        BaudRate(8) = 153600
        BaudRate(9) = 230400
        'BaudRate(10) = 460800
        'BaudRate(11) = 691200
        'BaudRate(12) = 1382400

        'BaudRate = 115,200bpsでCOMポート確認
        Try
            SerialPort1.BaudRate = BaudRate(7)
            SerialPort1.Open()
        Catch ex As Exception
            MessageBox.Show("【シリアルポートエラー】" & vbCrLf & "COMポートが開けませんでした")
            TextBox3.Text = "Port Error"
            GroupBox4.Enabled = True
            Exit Sub
        End Try

        'SerialPort1のBaudRate設定可能な範囲か確認
        'サーボに直接通信する場合
        For i = 0 To 9
            ComboBox2.Text = BaudRate(i)
            Try
                SerialPort1.BaudRate = BaudRate(i)
            Catch ex As Exception
                MessageBox.Show("【サーボが見つかりませんでした】" & vbCrLf & "サーボが正しく接続されているか確認してください" & vbCrLf & "サーボが2個以上接続されていないか確認してください")
                TextBox3.Text = "Not Found"
                GroupBox4.Enabled = True
                NumericUpDown1.Value = 1
                SerialPort1.Close()
                Exit Sub
            End Try

            '各BaudRateに対しID=255でAckを実行し、サーボが接続されているか確認する。
            If CheckAck(255) = 1 Then
                Exit For
            ElseIf i = 9 Then
                MessageBox.Show("【サーボが見つかりませんでした】" & vbCrLf & "サーボが正しく接続されているか確認してください" & vbCrLf & "サーボが2個以上接続されていないか確認してください")
                ComboBox2.Text = "115,200"
                TextBox3.Text = "Not Found"
                SerialPort1.BaudRate = 115200
                GroupBox4.Enabled = True
                NumericUpDown1.Value = 1
                SerialPort1.Close()
                Exit Sub
            End If
            Wait(10)
        Next

        'サーボの接続が確認されたBaudRateに対し、ID1～127の順で検索
        For pID = 1 To 127
            NumericUpDown1.Value = pID
            Ack = CheckAck(pID)
            If Ack = 1 Then
                TextBox3.Text = "OK"
                Exit For
            ElseIf pID = 127 And Ack = 0 Then
                TextBox3.Text = "NG"
            End If
        Next

        '接続が確認されたサーボのデータ取得
        Call GetParameters()

        'ボタン有効化
        GroupBox2.Enabled = True
        GroupBox4.Enabled = True

        If SerialPort1.IsOpen = True Then
            SerialPort1.Close()
        End If

    End Sub

    'パラメータセットボタン
    Private Sub SetButton_Clicked(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click, Button3.Click, Button4.Click, Button5.Click, Button6.Click, Button7.Click, Button8.Click, Button9.Click, Button10.Click, Button11.Click, Button12.Click, Button12.Click, Button13.Click, Button14.Click, Button15.Click, Button16.Click, Button17.Click, Button18.Click, Button19.Click, Button20.Click, Button21.Click, Button27.Click, Button28.Click, Button30.Click, Button32.Click, Button33.Click, Button34.Click, Button35.Click, Button36.Click, Button38.Click, Button39.Click, Button41.Click, Button40.Click, Button42.Click, Button43.Click
        Dim Packet(255) As Byte
        Dim temp As Integer

        'Flag:0（返信要求無し）
        Packet(3) = &H0

        'Count:1（1個だけに送信）
        Packet(6) = &H1

        If sender Is Button2 Then
            'IDボタン
            Packet(4) = &H4
            Packet(5) = &H1
            Packet(7) = NumericUpDown2.Value

        ElseIf sender Is Button21 Then
            'Reverseボタン
            Packet(4) = &H5
            Packet(5) = &H1

            Select Case ComboBox6.Text
                Case "Normal"
                    Packet(7) = &H0
                Case "Reverse"
                    Packet(7) = &H1
            End Select

        ElseIf sender Is Button3 Then
            'Baud Rateボタン
            Packet(4) = &H6
            Packet(5) = &H1

            Select Case ComboBox3.Text
                Case "9,600"
                    Packet(7) = &H0
                Case "14,400"
                    Packet(7) = &H1
                Case "19,200"
                    Packet(7) = &H2
                Case "28,800"
                    Packet(7) = &H3
                Case "38,400"
                    Packet(7) = &H4
                Case "57,600"
                    Packet(7) = &H5
                Case "76,800"
                    Packet(7) = &H6
                Case "115,200"
                    Packet(7) = &H7
                Case "153,600"
                    Packet(7) = &H8
                Case "230,400"
                    Packet(7) = &H9
                    'Case "460,800"
                    '    Packet(7) = &HA
                    'Case "691,200"
                    '    Packet(7) = &HB
                    'Case "1,382,400"
                    '    Packet(7) = &HC
                Case Else
                    Packet(7) = &H7
            End Select

        ElseIf sender Is Button28 Then
            'Return Delayボタン
            Packet(4) = 7
            Packet(5) = 1

            Packet(7) = NumericUpDown19.Value

        ElseIf sender Is Button4 Then
            'Angle Limit(CW)ボタン
            Packet(4) = 8
            Packet(5) = 2

            temp = NumericUpDown3.Value

            TrackBar1.Maximum = temp
            'NumericUpDown13.Maximum = temp
            'NumericUpDown18.Maximum = temp
            Label106.Text = temp.ToString()

            Packet(7) = temp And &HFF
            Packet(8) = Int(temp / 256) And &HFF

        ElseIf sender Is Button5 Then
            'Angle Limit(CCW)ボタン
            Packet(4) = 10
            Packet(5) = 2

            temp = NumericUpDown4.Value

            TrackBar1.Minimum = temp
            'NumericUpDown13.Minimum = temp
            'NumericUpDown18.Minimum = temp
            Label105.Text = temp.ToString()

            Packet(7) = temp And &HFF
            Packet(8) = Int(temp / 256) And &HFF

        ElseIf sender Is Button38 Then
            'Origin Positionボタン
            Packet(4) = 12
            Packet(5) = 2

            temp = NumericUpDown26.Value

            Packet(7) = temp And &HFF
            Packet(8) = Int(temp / 256) And &HFF

        ElseIf sender Is Button33 Then
            'Temperature Limitボタン
            Packet(4) = 14
            Packet(5) = 2

            temp = NumericUpDown21.Value

            Packet(7) = temp And &HFF
            Packet(8) = Int(temp / 256) And &HFF

        ElseIf sender Is Button34 Then
            'Speed Limit(CW)ボタン
            Packet(4) = 16
            Packet(5) = 1

            Packet(7) = NumericUpDown22.Value

        ElseIf sender Is Button35 Then
            'Speed Limit(CCW)ボタン
            Packet(4) = 17
            Packet(5) = 1

            temp = NumericUpDown23.Value

            If temp < 0 Then
                Packet(7) = temp + 256
            Else
                Packet(7) = temp
            End If

        ElseIf sender Is Button36 Then
            'Torque Limitボタン
            Packet(4) = 18
            Packet(5) = 2

            temp = NumericUpDown24.Value

            Packet(7) = temp And &HFF
            Packet(8) = Int(temp / 256) And &HFF

            'ElseIf sender Is Button37 Then
            '    'Torque Limit(CCW)ボタン
            '    Packet(4) = 19
            '    Packet(5) = 1

            '    temp = NumericUpDown25.Value

            '    If temp < 0 Then
            '        Packet(7) = temp + 256
            '    Else
            '        Packet(7) = temp
            '    End If

        ElseIf sender Is Button6 Then
            'Damperボタン
            Packet(4) = 20
            Packet(5) = 1

            Packet(7) = NumericUpDown5.Value

        ElseIf sender Is Button7 Then
            'Torque in Silenceボタン
            Packet(4) = 22
            Packet(5) = 1

            Select Case ComboBox7.Text
                Case "OFF"
                    Packet(7) = &H0
                Case "ON"
                    Packet(7) = &H1
                Case "Brake"
                    Packet(7) = &H2
            End Select

        ElseIf sender Is Button8 Then
            'Warm-Up Timeボタン
            Packet(4) = 23
            Packet(5) = 1

            Packet(7) = NumericUpDown7.Value

        ElseIf sender Is Button9 Then
            'Compliance Margin(CW)ボタン
            Packet(4) = 24
            Packet(5) = 1

            Packet(7) = NumericUpDown8.Value

        ElseIf sender Is Button10 Then
            'Compliance Margin(CCW)ボタン
            Packet(4) = 25
            Packet(5) = 1

            Packet(7) = NumericUpDown9.Value

        ElseIf sender Is Button11 Then
            'Compliance Slope(CW)ボタン
            Packet(4) = 26
            Packet(5) = 1

            Packet(7) = NumericUpDown10.Value

        ElseIf sender Is Button12 Then
            'Compliance Slope(CCW)ボタン
            Packet(4) = 27
            Packet(5) = 1

            Packet(7) = NumericUpDown11.Value

        ElseIf sender Is Button13 Then
            'Punchボタン
            Packet(4) = 28
            Packet(5) = 2

            temp = NumericUpDown12.Value
            Packet(7) = temp And &HFF
            Packet(8) = Int(temp / 256) And &HFF

        ElseIf sender Is Button14 Then
            'Goal Positionボタン
            Packet(4) = 30
            Packet(5) = 2

            temp = NumericUpDown13.Value

            Packet(7) = temp And &HFF
            Packet(8) = Int(temp / 256) And &HFF

            NumericUpDown18.Value = temp

            If temp > TrackBar1.Maximum Then
                TrackBar1.Value = TrackBar1.Maximum
            ElseIf temp < TrackBar1.Minimum Then
                TrackBar1.Value = TrackBar1.Minimum
            Else
                TrackBar1.Value = temp
            End If

        ElseIf sender Is Button27 Then
            'Set"0"ボタン
            Packet(4) = 30
            Packet(5) = 2

            Packet(7) = 0
            Packet(8) = 0
            NumericUpDown18.Value = 0
            NumericUpDown13.Value = 0
            TrackBar1.Value = 0

        ElseIf sender Is Button30 Then
            '動作ボタン
            Packet(4) = 30
            Packet(5) = 2

            temp = NumericUpDown18.Value

            Packet(7) = temp And &HFF
            Packet(8) = Int(temp / 256) And &HFF

            NumericUpDown13.Value = temp

            If temp > TrackBar1.Maximum Then
                TrackBar1.Value = TrackBar1.Maximum
            ElseIf temp < TrackBar1.Minimum Then
                TrackBar1.Value = TrackBar1.Minimum
            Else
                TrackBar1.Value = temp
            End If


        ElseIf sender Is Button15 Then
            'Goal Timeボタン
            Packet(4) = 32
            Packet(5) = 2

            temp = NumericUpDown14.Value
            Packet(7) = temp And &HFF
            Packet(8) = Int(temp / 256) And &HFF

        ElseIf sender Is Button16 Then
            'Acceleration timeボタン
            Packet(4) = 34
            Packet(5) = 1

            Packet(7) = NumericUpDown15.Value

        ElseIf sender Is Button17 Then
            'Max Torqueボタン
            Packet(4) = 35
            Packet(5) = 1

            Packet(7) = NumericUpDown16.Value

        ElseIf sender Is Button18 Then
            'トルクEnableボタン
            Packet(4) = 36
            Packet(5) = 1

            Select Case ComboBox4.Text
                Case "OFF"
                    Packet(7) = &H0
                Case "ON"
                    Packet(7) = &H1
                Case "Brake"
                    Packet(7) = &H2
            End Select

        ElseIf sender Is Button32 Then
            'Goal Speedボタン
            Packet(4) = 37
            Packet(5) = 2

            temp = NumericUpDown6.Value

            Packet(7) = temp And &HFF
            Packet(8) = Int(temp / 256) And &HFF

        ElseIf sender Is Button20 Then
            'Goal Torqueボタン
            Packet(4) = 39
            Packet(5) = 2

            temp = NumericUpDown17.Value

            Packet(7) = temp And &HFF
            Packet(8) = Int(temp / 256) And &HFF

        ElseIf sender Is Button19 Then
            'Cascade Enableボタン
            Packet(4) = 21
            Packet(5) = 1

            Select Case ComboBox5.Text
                Case "OFF"
                    Packet(7) = 0
                Case "ON"
                    Packet(7) = 1
            End Select

        ElseIf sender Is Button40 Then
            'Currenr PGain ボタン
            Packet(4) = 54
            Packet(5) = 1

            Packet(7) = NumericUpDown28.Value

        ElseIf sender Is Button41 Then
            'Currenr Deadband ボタン
            Packet(4) = 55
            Packet(5) = 1

            Packet(7) = NumericUpDown29.Value

        ElseIf sender Is Button39 Then
            'Currenr Deadband Output Rateボタン
            Packet(4) = 56
            Packet(5) = 1

            Packet(7) = NumericUpDown27.Value

        ElseIf sender Is Button42 Then
            'Speed Deadband ボタン
            Packet(4) = 57
            Packet(5) = 1

            Packet(7) = NumericUpDown30.Value

        ElseIf sender Is Button43 Then
            'BootloaderKey ボタン
            Packet(4) = 58
            Packet(5) = 1

            Packet(7) = NumericUpDown31.Value

        End If

        'パケット送信
        Call SendPacket(Packet)

        'IDを変更した場合は変更後のIDでデータ再取得(ID=255の場合除く）
        If sender Is Button2 And CheckBox2.Checked = False Then
            NumericUpDown1.Value = NumericUpDown2.Value
            GroupBox2.Enabled = False
            'シリアルポート開く
            Try
                SerialPort1.Open()
            Catch ex As Exception
                MessageBox.Show("【COMポートエラー】" & vbCrLf & "COMポートの指定を確認してください。")
            End Try

            Call GetParameters()

            If SerialPort1.IsOpen = True Then
                SerialPort1.Close()
            End If

            GroupBox2.Enabled = True
        End If

    End Sub

    'GetParametersボタン
    Private Sub GetButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Button1.Enabled = False

        Try
            SerialPort1.Open()
        Catch ex As Exception
            MessageBox.Show("【COMポートエラー】" & vbCrLf & "COMポートの指定を確認してください。")
            Button1.Enabled = True
            Exit Sub
        End Try

        Call GetParameters()

        SerialPort1.Close()
        Button1.Enabled = True

    End Sub

    'No.4-No.29セットボタン
    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        Dim packet(40) As Byte
        Dim temp As Integer
        Button22.Enabled = False

        'Flag:0（返信要求無し）
        packet(3) = &H0

        '開始アドレス
        packet(4) = 4

        'データのバイト数
        packet(5) = 26

        'Count:1（1個だけに送信）
        packet(6) = &H1

        packet(7) = NumericUpDown2.Value

        Select Case ComboBox6.Text
            Case "Normal"
                packet(8) = 0
            Case "Reverse"
                packet(8) = 1
        End Select

        Select Case ComboBox3.Text
            Case "9, 600"
                packet(9) = &H0
            Case "14, 400"
                packet(9) = &H1
            Case "19, 200"
                packet(9) = &H2
            Case "28, 800"
                packet(9) = &H3
            Case "38, 400"
                packet(9) = &H4
            Case "57, 600"
                packet(9) = &H5
            Case "76, 800"
                packet(9) = &H6
            Case "115, 200"
                packet(9) = &H7
            Case "153, 600"
                packet(9) = &H8
            Case "230, 400"
                packet(9) = &H9
                '  Case "460, 800"
                '      packet(9) = &HA
                '  Case "691, 200"
                '     packet(9) = &HB
                ' Case "1, 382, 400"
                '     packet(9) = &HC
            Case Else
                packet(9) = &H7
        End Select

        packet(10) = NumericUpDown19.Value

        temp = NumericUpDown3.Value
        packet(11) = temp And &HFF
        packet(12) = Int(temp / 256) And &HFF

        temp = NumericUpDown4.Value
        packet(13) = temp And &HFF
        packet(14) = Int(temp / 256) And &HFF

        temp = NumericUpDown26.Value 'CInt(TextBox6.Text)
        packet(15) = temp And &HFF
        packet(16) = Int(temp / 256) And &HFF

        temp = NumericUpDown21.Value
        packet(17) = temp And &HFF
        packet(18) = Int(temp / 256) And &HFF

        packet(19) = NumericUpDown22.Value

        temp = NumericUpDown23.Value
        If temp < 0 Then
            packet(7) = NumericUpDown23.Value + 256
        Else
            packet(7) = temp
        End If

        'packet(21) = NumericUpDown24.Value

        temp = NumericUpDown24.Value
        packet(21) = temp And &HFF
        packet(22) = Int(temp / 256) And &HFF

        'temp = NumericUpDown25.Value
        'If temp < 0 Then
        '    packet(7) = NumericUpDown25.Value + 256
        'Else
        '    packet(7) = temp
        'End If

        packet(23) = NumericUpDown5.Value

        packet(24) = CInt(TextBox12.Text)

        Select Case ComboBox7.Text
            Case "OFF"
                packet(25) = 0
            Case "ON"
                packet(25) = 1
            Case "Brake"
                packet(25) = 2
        End Select
        packet(26) = NumericUpDown7.Value
        packet(27) = NumericUpDown8.Value
        packet(28) = NumericUpDown9.Value
        packet(29) = NumericUpDown10.Value
        packet(30) = NumericUpDown11.Value

        temp = NumericUpDown12.Value
        packet(31) = temp And &HFF
        packet(32) = Int(temp / 256) And &HFF

        Call SendPacket(packet)

        Button22.Enabled = True

    End Sub

    'No.30-No.41セットボタン
    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        Dim packet(40) As Byte
        Dim temp As Integer
        Button23.Enabled = False

        'Flag:0（返信要求無し）
        packet(3) = &H0

        '開始アドレス
        packet(4) = 30

        'データのバイト数
        packet(5) = 12

        'Count:1（1個だけに送信）
        packet(6) = &H1

        temp = NumericUpDown13.Value
        packet(7) = temp And &HFF
        packet(8) = Int(temp / 256) And &HFF

        temp = NumericUpDown14.Value
        packet(9) = temp And &HFF
        packet(10) = Int(temp / 256) And &HFF

        packet(11) = NumericUpDown15.Value
        packet(12) = NumericUpDown16.Value

        Select Case ComboBox4.Text
            Case "OFF"
                packet(13) = 0
            Case "ON"
                packet(13) = 1
            Case "Brake "
                packet(13) = 2
        End Select

        temp = NumericUpDown6.Value
        packet(14) = temp And &HFF
        packet(15) = Int(temp / 256) And &HFF

        temp = NumericUpDown17.Value
        packet(16) = temp And &HFF
        packet(17) = Int(temp / 256) And &HFF

        Select Case ComboBox5.Text
            Case "OFF"
                packet(18) = 0
            Case "ON"
                packet(18) = 1
        End Select

        Call SendPacket(packet)

        Button23.Enabled = True
    End Sub

    'ACKボタン
    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button29.Click
        Dim pID, Ack As Integer
        Button1.Enabled = False
        Button24.Enabled = False
        Button25.Enabled = False
        Button26.Enabled = False
        Button29.Enabled = False
        CheckBox2.Enabled = False
        NumericUpDown1.Enabled = False
        TextBox3.Text = "Checking..."

        pID = NumericUpDown1.Value

        Try
            SerialPort1.Open()
        Catch ex As Exception
            MessageBox.Show("シリアルポートエラー")
        End Try

        Ack = CheckAck(pID)

        If Ack = 1 Then
            TextBox3.Text = "OK"
        ElseIf Ack = 0 Then
            TextBox3.Text = "NG"
        End If

        If SerialPort1.IsOpen = True Then
            SerialPort1.Close()
        End If

        Button24.Enabled = True
        Button25.Enabled = True
        Button26.Enabled = True
        Button29.Enabled = True
        CheckBox2.Enabled = True

        If CheckBox2.Checked = True And Ack = 0 Then
            Button1.Enabled = False
            NumericUpDown1.Enabled = False
        Else
            Button1.Enabled = True
            NumericUpDown1.Enabled = True
        End If

    End Sub

    'トラックバー操作
    Private Sub TrackBar1_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.Scroll
        NumericUpDown18.Value = TrackBar1.Value
        NumericUpDown13.Value = TrackBar1.Value
        Call MoveSx()
    End Sub

    'データ取得＆表示
    Sub GetParameters()
        Dim RequestReturn(11), ReturnPacket(100) As Byte
        Dim pGetAddress, pGetLength, i, j, k, sum, temp As Integer


        j = 0
        k = 0

        For i = 0 To 99
            ReturnPacket(i) = &H1
        Next
        sum = CalSum(ReturnPacket)

        'リターンデータ取得範囲設定
        pGetAddress = 0
        pGetLength = 60

        'リターンデータ要求パケット作成
        RequestReturn(0) = &HFA
        RequestReturn(1) = &HAF
        RequestReturn(2) = NumericUpDown1.Value
        RequestReturn(3) = &HF             'Flg
        RequestReturn(4) = pGetAddress     'Address
        RequestReturn(5) = pGetLength      'Length
        RequestReturn(6) = &H0             'Count
        RequestReturn(7) = CalSum(RequestReturn)             'Checksum

        'リターン要求送信
        If SerialPort1.IsOpen Then

            'サーボに直接送信
            SerialPort1.Write(RequestReturn, 0, 8)

            '受信バイト数が予定数になるまで待機
            Do Until SerialPort1.BytesToRead = 68 Or k > 100
                k = k + 1
                Wait(1)
            Loop

            If k > 100 Then
                SerialPort1.Close()
                MessageBox.Show("【リターンデータの取得に失敗しました】" & vbCrLf & "COMポートとIDを確認してください")
                GroupBox2.Enabled = False
                Button1.Enabled = True
                Button24.Enabled = False
                Button25.Enabled = False
                Exit Sub
            End If

            '受信データ取得
            SerialPort1.Read(ReturnPacket, 0, 68)
            SerialPort1.Close()
        End If

        'チェックサム確認
        If ReturnPacket(67) <> CalSum(ReturnPacket) Then
            MessageBox.Show("【リターンデータのチェックサムが不正でした】" & vbCrLf & "・もう一度データ取得を実行してください")
            GroupBox2.Enabled = False
            Exit Sub
        End If

        'データの表示
        '機種名
        TextBox1.Text = Hex(ReturnPacket(7) + 256 * ReturnPacket(8))

        '機種に応じて表示内容編集
        Call SelectServo(ReturnPacket(7) + 256 * ReturnPacket(8))

        'Firmware Version
        TextBox2.Text = ReturnPacket(9)

        'Reserved
        TextBox4.Text = ReturnPacket(10)

        'ID
        NumericUpDown2.Value = ReturnPacket(11)

        '回転方向:0なら正転、1なら反転
        Select Case ReturnPacket(12)
            Case 0
                ComboBox6.Text = "Normal"
            Case 1
                ComboBox6.Text = "Reverse"
        End Select

        'BaudRate:対応表
        Select Case ReturnPacket(13)
            Case 0
                ComboBox3.Text = "9,600"
            Case 1
                ComboBox3.Text = "14,400"
            Case 2
                ComboBox3.Text = "19,200"
            Case 3
                ComboBox3.Text = "28,800"
            Case &H4
                ComboBox3.Text = "38,400"
            Case 5
                ComboBox3.Text = "57,600"
            Case 6
                ComboBox3.Text = "76,800"
            Case 7
                ComboBox3.Text = "115,200"
            Case 8
                ComboBox3.Text = "153,600"
            Case 9
                ComboBox3.Text = "230,400"
           ' Case 10
                '    ComboBox3.Text = "460,800"
                'Case 11
                '    ComboBox3.Text = "691,200"
                'Case 12
                '    ComboBox3.Text = "1,382,400"
            Case Else
                ComboBox3.Text = "---"
        End Select

        'Return Delay
        NumericUpDown19.Value = ReturnPacket(14)

        'Angle Limit(CW)
        temp = ReturnPacket(15) + 256 * ReturnPacket(16)

        If temp > &H7FFF Then
            NumericUpDown3.Value = temp - 65536
        Else
            NumericUpDown3.Value = temp
        End If

        TrackBar1.Maximum = NumericUpDown3.Value
        Label106.Text = NumericUpDown3.Value.ToString()

        'Angle Limit(CCW)
        temp = ReturnPacket(17) + 256 * ReturnPacket(18)
        If temp > &H7FFF Then
            NumericUpDown4.Value = temp - 65536
        Else
            NumericUpDown4.Value = temp
        End If

        TrackBar1.Minimum = NumericUpDown4.Value
        Label105.Text = NumericUpDown4.Value.ToString()

        'Origin Position
        temp = ReturnPacket(19) + 256 * ReturnPacket(20)

        If temp > 32767 Then
            NumericUpDown26.Value = temp - 65536
        Else
            NumericUpDown26.Value = temp
        End If

        'Temperature Limit
        temp = ReturnPacket(21) + 256 * ReturnPacket(22)
        If temp > &H7FFF Then
            NumericUpDown21.Value = temp - 65536
        Else
            NumericUpDown21.Value = temp
        End If

        'Speed Limit(CW)
        NumericUpDown22.Value = ReturnPacket(23)

        'Speed Limit(CCW)
        temp = ReturnPacket(24)
        If temp > &H7F Then
            NumericUpDown23.Value = temp - 256
        Else
            NumericUpDown23.Value = temp
        End If

        'Torque Limit
        temp = ReturnPacket(25) + 256 * ReturnPacket(26)
        If temp > &H7FFF Then
            NumericUpDown24.Value = temp - 65536
        Else
            NumericUpDown24.Value = temp
        End If

        ''Torque Limit(CW)
        'NumericUpDown24.Value = ReturnPacket(25)

        ''Torque Limit(CCW)
        'temp = ReturnPacket(26)
        'If temp > &H7F Then
        '    NumericUpDown25.Value = temp - 256
        'Else
        '    NumericUpDown25.Value = temp
        'End If

        'Damper
        NumericUpDown5.Value = ReturnPacket(27)

        'CascadeEnable
        Select Case ReturnPacket(28)
            Case 0
                ComboBox5.Text = "OFF"
            Case 1
                ComboBox5.Text = "ON"
        End Select

        'Torque in Silence
        Select Case ReturnPacket(29)
            Case 0
                ComboBox7.Text = "OFF"
            Case 1
                ComboBox7.Text = "ON"
            Case 2
                ComboBox7.Text = "Brake"
        End Select

        'Warm-Up Time
        NumericUpDown7.Value = ReturnPacket(30)

        'Margin(CW)
        NumericUpDown8.Value = ReturnPacket(31)

        'Margin(CCW)
        NumericUpDown9.Value = ReturnPacket(32)

        'Slope(CW)
        NumericUpDown10.Value = ReturnPacket(33)

        'Slope(CCW)
        NumericUpDown11.Value = ReturnPacket(34)

        'Punch
        NumericUpDown12.Value = ReturnPacket(35) + 256 * ReturnPacket(36)

        'Goal Position
        temp = ReturnPacket(37) + 256 * ReturnPacket(38)

        If temp > 32767 Then
            NumericUpDown13.Value = temp - 65536
        Else
            NumericUpDown13.Value = temp
        End If

        'Goal Time
        NumericUpDown14.Value = ReturnPacket(39) + 256 * ReturnPacket(40)

        'Acceleration Time
        NumericUpDown15.Value = ReturnPacket(41)

        'Max Torque
        NumericUpDown16.Value = ReturnPacket(42)

        'Torque ON
        Select Case ReturnPacket(43)
            Case 0
                ComboBox4.Text = "OFF"
            Case 1
                ComboBox4.Text = "ON"
            Case 2
                ComboBox4.Text = "Brake"
        End Select

        'Goal Speed
        temp = ReturnPacket(44) + 256 * ReturnPacket(45)

        If temp > 32767 Then
            NumericUpDown6.Value = temp - 65536
        Else
            NumericUpDown6.Value = temp
        End If

        'Goal Torque
        temp = ReturnPacket(46) + 256 * ReturnPacket(47)

        If temp > 32767 Then
            NumericUpDown17.Value = temp - 65536
        Else
            NumericUpDown17.Value = temp
        End If

        'Reserved
        TextBox12.Text = ReturnPacket(48)

        'Present Position
        temp = ReturnPacket(49) + 256 * ReturnPacket(50)
        If temp > 32767 Then
            TextBox16.Text = temp - 65536
        Else
            TextBox16.Text = temp
        End If

        'Present Time
        TextBox17.Text = ReturnPacket(51) + 256 * ReturnPacket(52)

        'Present Speed
        temp = ReturnPacket(53) + 256 * ReturnPacket(54)
        If temp > 32767 Then
            TextBox18.Text = temp - 65536
        Else
            TextBox18.Text = temp
        End If

        'Present Current
        TextBox19.Text = ReturnPacket(55) + 256 * ReturnPacket(56)

        'Present Temperature
        TextBox20.Text = ReturnPacket(57) + 256 * ReturnPacket(58)

        'Present Voltage
        TextBox21.Text = ReturnPacket(59) + 256 * ReturnPacket(60)

        'Reserved
        'NumericUpDown28.Value = ReturnPacket(61)
        'NumericUpDown29.Value = ReturnPacket(62)
        'NumericUpDown27.Value = ReturnPacket(63)
        'NumericUpDown30.Value = ReturnPacket(64)
        NumericUpDown31.Value = ReturnPacket(65)
        TextBox27.Text = ReturnPacket(66)

        '表示欄有効化
        GroupBox2.Enabled = True
        Button24.Enabled = True
        Button25.Enabled = True

    End Sub

    'パケット送信（サーボに直接／RPU経由）
    Sub SendPacket(ByVal Packet() As Byte)
        Dim PacketLength, sum As Integer

        'パケットにヘッダー追加
        Packet(0) = &HFA
        Packet(1) = &HAF

        'パケットにID追加
        Packet(2) = NumericUpDown1.Value

        'パケットのチェックサム計算
        PacketLength = 7 + Packet(5) * Packet(6) + 1
        sum = Packet(2)

        For i = 3 To (PacketLength - 2)
            sum = sum Xor Packet(i)
        Next

        'パケットにチェックサム追加
        Packet(PacketLength - 1) = sum

        'スルーコマンド用パケット定義
        Dim RPUPacket(PacketLength + 2) As Byte

        'シリアルポート開く
        Try
            SerialPort1.Open()
        Catch ex As Exception
            MessageBox.Show("【COMポートエラー】" & vbCrLf & "COMポートの指定を確認してください。")
        End Try

        If SerialPort1.IsOpen = True Then

            'サーボに直接通信する場合はそのまま送信
            SerialPort1.Write(Packet, 0, PacketLength)

            SerialPort1.Close()
        End If

    End Sub

    'フラッシュROM書き込み送信
    Sub WriteFlashROM()
        Dim Packet(8) As Byte

        '書き込みフラグ用パケット生成
        Packet(3) = &H40
        Packet(4) = &HFF
        Packet(5) = &H0
        Packet(6) = &H0

        '書き込みパケット送信
        Call SendPacket(Packet)

        If SerialPort1.IsOpen = True Then
            SerialPort1.Close()
        End If

        '書き込み完了まで待機
        '待ち時間は機種により変更
        Select Case TextBox1.Text
            Case "6010", "4010", "4020"
                Wait(10000)
            Case Else
                Wait(3000)
        End Select

        '書き込み完了後再起動
        Call ResetSx()
        MessageBox.Show("FlashROM書き込み完了")
    End Sub

    'サーボ動作送信
    Sub MoveSx()
        Dim Packet(10) As Byte
        Dim temp As Integer

        '目標角度指定
        Packet(3) = &H0
        Packet(4) = 30
        Packet(5) = 2
        Packet(6) = 1

        '目標角度欄の表示内容修正
        temp = NumericUpDown18.Value
        NumericUpDown13.Value = temp

        Packet(7) = temp And &HFF
        Packet(8) = Int(temp / 256) And &HFF
        Call SendPacket(Packet)

    End Sub

    'サーボ再起動パケット送信
    Sub ResetSx()
        Dim Packet(8) As Byte

        '再起動パケット生成
        Packet(3) = &H20
        Packet(4) = &HFF
        Packet(5) = &H0
        Packet(6) = &H0

        '再起動パケット送信
        Call SendPacket(Packet)

    End Sub

    'メモリマップNo.4～29初期化パケット送信
    Sub InitializeSx()
        Dim Packet(8) As Byte

        '初期化フラグ用パケット生成
        Packet(3) = &H10
        Packet(4) = &HFF
        Packet(5) = &HFF
        Packet(6) = &H0

        '初期化フラグ用パケット送信
        Call SendPacket(Packet)

    End Sub

    'Ack確認
    Function CheckAck(ByVal ID As Integer) As Integer
        Dim RequestReturn(11) As Byte
        Dim ReturnPacket(1) As Byte
        Dim k As Integer

        'ACK要求パケット作成
        RequestReturn(0) = &HFA
        RequestReturn(1) = &HAF
        RequestReturn(2) = ID
        RequestReturn(3) = &H1
        RequestReturn(4) = &H0
        RequestReturn(5) = &H0
        RequestReturn(6) = &H0
        RequestReturn(7) = CalSum(RequestReturn)

        If SerialPort1.IsOpen Then

            'サーボに直接送信
            SerialPort1.Write(RequestReturn, 0, 8)

            '受信バイト数が予定数になるまで待機
            Do Until SerialPort1.BytesToRead = 1 Or k > 100
                k = k + 1
                Wait(1)
            Loop

            '受信データ取得
            If SerialPort1.BytesToRead = 1 Then
                SerialPort1.Read(ReturnPacket, 0, 1)
            Else
                ReturnPacket(0) = 0
            End If

        End If


        If ReturnPacket(0) = &H7 Then
            CheckAck = 1
        Else
            CheckAck = 0
        End If

    End Function

    'パケットのチェックサム計算
    Function CalSum(ByVal Packet() As Byte) As Integer
        Dim packetLength As Integer

        'パケット長確認
        packetLength = 7 + Packet(5) * Packet(6)

        'パケット長が256を超える場合はエラー
        If packetLength > 256 Then
            CalSum = 257
            Exit Function
        End If

        'チェックサム計算
        CalSum = Packet(2)

        For i = 3 To (packetLength - 1)
            CalSum = CalSum Xor Packet(i)
        Next

    End Function

    '機種ごとの設定内容変更
    Sub SelectServo(ByVal ModelNumber As Integer)
        '機種コードに応じて設定項目名、設定可能範囲変更
        Select Case ModelNumber

            Case &H125C
                ComboBox6.Enabled = True
                Button21.Enabled = True
                Label16.Text = "Reverse"
                NumericUpDown19.Enabled = True
                Button28.Enabled = True
                Label20.Text = "Return Delay"
                'NumericUpDown3.Maximum = 1500
                'NumericUpDown4.Minimum = -1500
                NumericUpDown5.Enabled = True
                Button6.Enabled = True
                Label38.Text = "Damper"
                ComboBox7.Enabled = True
                Button7.Enabled = True
                Label45.Text = "Torque in Silence"
                NumericUpDown7.Enabled = True
                Button8.Enabled = True
                Label47.Text = "Warm-Up Time"
                NumericUpDown10.Maximum = 255
                NumericUpDown11.Maximum = 255
                NumericUpDown12.Maximum = &H2710
                'NumericUpDown13.Maximum = 1500
                'NumericUpDown13.Minimum = -1500
                NumericUpDown15.Enabled = False
                Button16.Enabled = False
                Label63.Text = "Reserved"
                ComboBox5.Enabled = True
                Button19.Enabled = True
                'Label69.Text = "Reserved"
                NumericUpDown17.Enabled = True
                Button20.Enabled = True
                'Label71.Text = "Reserved"
                Label86.Text = "Present Speed"
                Label92.Text = "Present Voltage"
                'Label94.Text = "Reserved"
                'Label96.Text = "Reserved"
                'Label98.Text = "Reserved"
                'Label100.Text = "Reserved"
                NumericUpDown6.Enabled = True
                Button32.Enabled = True

                NumericUpDown28.Enabled = False
                NumericUpDown29.Enabled = False
                NumericUpDown27.Enabled = False
                NumericUpDown30.Enabled = False
                Button40.Enabled = False
                Button41.Enabled = False
                Button39.Enabled = False
                Button42.Enabled = False

        End Select


    End Sub

    'wait関数
    Sub Wait(ByVal waitTime As Long)
        Dim startTime As Single

        startTime = timeGetTime

        Do Until (timeGetTime - startTime) >= waitTime
            System.Windows.Forms.Application.DoEvents()
        Loop
    End Sub

End Class
