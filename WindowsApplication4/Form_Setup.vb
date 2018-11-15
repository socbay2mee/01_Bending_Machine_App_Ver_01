Public Class Form_Setup
    Dim objIniFile As New clsIni("E:\WorldInfo.ini")
    Public Class clsIni
        ' API functions
        Private Declare Ansi Function GetPrivateProfileString _
          Lib "kernel32.dll" Alias "GetPrivateProfileStringA" _
          (ByVal lpApplicationName As String, _
          ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As System.Text.StringBuilder, _
          ByVal nSize As Integer, ByVal lpFileName As String) As Integer

        Private Declare Ansi Function WritePrivateProfileString _
          Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
          (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer

        Private Declare Ansi Function GetPrivateProfileInt _
          Lib "kernel32.dll" Alias "GetPrivateProfileIntA" _
          (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer

        Private Declare Ansi Function FlushPrivateProfileString _
          Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
          (ByVal lpApplicationName As Integer, ByVal lpKeyName As Integer, ByVal lpString As Integer, ByVal lpFileName As String) As Integer
        Dim strFilename As String

        ' Constructor, accepting a filename
        Public Sub New(ByVal Filename As String)
            strFilename = Filename
        End Sub

        ' Read-only filename property
        ReadOnly Property FileName() As String
            Get
                Return strFilename
            End Get
        End Property

        Public Function GetString(ByVal Section As String, ByVal Key As String, ByVal [Default] As String) As String
            ' Returns a string from your INI file
            Dim intCharCount As Integer
            Dim objResult As New System.Text.StringBuilder(256)
            intCharCount = GetPrivateProfileString(Section, Key, [Default], objResult, objResult.Capacity, strFilename)
            Return objResult.ToString()
        End Function

        Public Function GetInteger(ByVal Section As String, ByVal Key As String, ByVal [Default] As Integer) As Integer
            ' Returns an integer from your INI file
            Return GetPrivateProfileInt(Section, Key, [Default], strFilename)
        End Function

        Public Function GetBoolean(ByVal Section As String, ByVal Key As String, ByVal [Default] As Boolean) As Boolean
            ' Returns a boolean from your INI file
            Return (GetPrivateProfileInt(Section, Key, CInt([Default]), strFilename) = 1)
        End Function

        Public Sub WriteString(ByVal Section As String, ByVal Key As String, ByVal Value As String)
            ' Writes a string to your INI file
            WritePrivateProfileString(Section, Key, Value, strFilename)
            Flush()
        End Sub

        Public Sub WriteInteger(ByVal Section As String, ByVal Key As String, ByVal Value As Integer)
            ' Writes an integer to your INI file
            WriteString(Section, Key, CStr(Value))
            Flush()
        End Sub

        Public Sub WriteBoolean(ByVal Section As String, ByVal Key As String, ByVal Value As Boolean)
            ' Writes a boolean to your INI file
            WriteString(Section, Key, CStr(CInt(Value)))
            Flush()
        End Sub

        Private Sub Flush()
            ' Stores all the cached changes to your INI file
            FlushPrivateProfileString(0, 0, 0, strFilename)
        End Sub
    End Class

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Txt_Cut_Angle.Text = objIniFile.GetString("Machine_Config", "V-CUT_Angle", "")
        Txt_Thickness.Text = objIniFile.GetString("Machine_Config", "Thickness", "")

        If (objIniFile.GetString("Machine_Config", "Inside", "") = "Clockwise") Then
            Clock_In.Checked = True
            Anti_In.Checked = False
        Else
            Anti_In.Checked = True
            Clock_In.Checked = False
        End If

        If (objIniFile.GetString("Machine_Config", "Outside", "") = "Clockwise") Then
            Clock_Out.Checked = True
            Anti_Out.Checked = False
        Else
            Anti_Out.Checked = True
            Clock_Out.Checked = False
        End If
    End Sub

    Private Sub Button_OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_OK.Click
        objIniFile.WriteString("Machine_Config", "V-CUT_Angle", Txt_Cut_Angle.Text)
        objIniFile.WriteString("Machine_Config", "Thickness", Txt_Thickness.Text)
        Me.Close()
    End Sub

    Private Sub Button_Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Cancel.Click
        Me.Close()
    End Sub

    Private Sub Clock_Out_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Clock_Out.CheckedChanged
        objIniFile.WriteString("Machine_Config", "Outside", "Clockwise")
    End Sub

    Private Sub Anti_Out_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Anti_Out.CheckedChanged
        objIniFile.WriteString("Machine_Config", "Outside", "Anti-clockwise")
    End Sub

    Private Sub Clock_In_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Clock_In.CheckedChanged
        objIniFile.WriteString("Machine_Config", "Inside", "Clockwise")
    End Sub

    Private Sub Anti_In_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Anti_In.CheckedChanged
        objIniFile.WriteString("Machine_Config", "Inside", "Anti-clockwise")
    End Sub
End Class