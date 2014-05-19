Public Class Form1

	'*-----------------------------------------------------------------------------*
	'* iniファイル読込み用
	'*-----------------------------------------------------------------------------*
	<System.Runtime.InteropServices.DllImport("KERNEL32.DLL")> _
	Public Shared Function GetPrivateProfileString( _
	ByVal lpAppName As String, _
	ByVal lpKeyName As String, ByVal lpDefault As String, _
	ByVal lpReturnedString As System.Text.StringBuilder, ByVal nSize As Integer, _
	ByVal lpFileName As String _
	) As Integer
	End Function

	'*-----------------------------------------------------------------------------*
	'* iniファイルから指定した文字列を取得
	'*-----------------------------------------------------------------------------*
	Private Function getIniFileString(ByVal fileName As String, ByVal sesName As String, ByVal KeyName As String) As String

		Dim strSb As New System.Text.StringBuilder
		' iniファイルの読込み
		GetPrivateProfileString(sesName, KeyName, "default", strSb, 255, fileName)

		Return strSb.ToString

	End Function

	Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
		Dim FolderName As String = Application.StartupPath
		''自分自身の実行ファイルのパスを取得する
		'Dim appPath As String = _
		'		System.Reflection.Assembly.GetExecutingAssembly().Location


		Dim strKey1Value = getIniFileString(FolderName & "\テスト.ini", "セッション１", "キー３")
		'Dim strKey1Value = getIniFileString(FolderName & "\テスト.ini", "aaa", "ccc")
		'Dim strKey1Value = getIniFileString("c:\テスト.ini", "セッション１", "キー３")
		If strKey1Value = "default" Then
			'エラー
			MsgBox("指定されたiniファイルから値が取得出来きません。")
			Exit Sub
		Else
			MsgBox("[取得]：【" & strKey1Value & "】")
		End If
	End Sub

	Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
		' 処理対象となる文字列変数を宣言する
		Dim stTarget As String = "1234"
		Dim i As Integer
		Dim st As String

		' 6 文字になるまで先頭を半角スペースで埋める
		MessageBox.Show("[" & stTarget.PadLeft(6) & "]")

		' 8 文字になるまで先頭を "0" で埋める
		MessageBox.Show("[" & stTarget.PadLeft(8, "0"c) & "]")

		For i = 1 To 10
			st = CStr(i).PadLeft(2, "0"c)
			MessageBox.Show(st)
		Next
	End Sub

	Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
		Dim FolderName As String = Application.StartupPath
		Dim FolderName2 As String = Application.StartupPath & "\"
		Dim st As String

		If System.IO.Directory.Exists(FolderName & "\") Then
			MessageBox.Show("有：[" & FolderName & "]")
		Else
			MessageBox.Show("無：[" & FolderName & "]")
		End If

		'[取得]最後の文字
		'st = Strings.Right(FolderName2, 1)
		st = FolderName2.Substring(FolderName2.Length - 1, 1)
		If st = "\" Then
			MessageBox.Show("えんです")
		Else
			MessageBox.Show("えんではありません")
		End If
	End Sub

	Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
		' 文字列を格納するための変数を宣言する
		Dim stTarget As String = "ABC1234あいうえお5678"
		Dim st As String
		Dim i As Integer = stTarget.Length
		st = stTarget.Substring(i - 1, 1)
		MessageBox.Show("[" & CStr(i) & "][" & st & "]")

		' 3 文字目の後から 9 文字の文字列を取得する
		stTarget = stTarget.Substring(3, 9)

		' 4 文字目の後から末尾までの文字列を取得する
		stTarget = stTarget.Substring(4)

		' 取得した文字列を表示する
		MessageBox.Show(stTarget)
	End Sub
End Class
