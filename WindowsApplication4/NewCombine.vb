﻿Option Strict On
Imports System
Imports System.IO.Ports
Imports System.Drawing.Drawing2D
Public Class Form1
    Private ppoint As Point
    Dim poly_len(100) As Double
    Dim coords(100, 5000) As PointF
    Dim ang_out(100, 5000) As Single
    Dim numofpoints(100) As Integer
    Dim numofpoly As Integer = 0
    Dim i As Integer
    Dim j As Integer
    Dim str As String
    Dim myPen As New Pen(Color.Red, 1)
    Dim gpanel As Graphics
    Dim minx As Single = Single.MaxValue
    Dim miny As Single = Single.MaxValue
    Dim maxx As Single = Single.MinValue
    Dim maxy As Single = Single.MinValue
    Dim sr As System.IO.StreamReader
    Dim Start_point() As Integer
    Dim start_select As Boolean = False
    Dim Lboundp(0 To 100) As PointF
    Dim Uboundp(0 To 100) As PointF
    Dim parea(0 To 100) As Single
    Dim Is_Inside(0 To 100) As Boolean
    Dim scx As Integer = 1
    Dim scy As Integer = 1
    Dim dx, dx_sc As Single
    Dim dy, dy_sc As Single
    Dim text_size As Integer = 8
    Dim pp As PointF
    Dim coord(0 To 5000) As PointF
    Dim right_direct(0 To 100) As Boolean
    Dim bend_display As Boolean = False
    Dim offs As Double
    Dim start_bend_of(0 To 500) As Integer
    Dim seg_bend(0 To 500) As Integer
    Dim coordsBackUp(0 To 500) As PointF
    Dim Processing As Boolean
    Dim offsX, offsY As Single
    Dim w, h As Single
    Dim scale_factor(0 To 20) As Single
    Dim is_selected(0 To 100) As Boolean
    Dim start_drag, end_drag As Point
    Dim is_dragging As Boolean = False
    Dim first_drag As Boolean = True
    Dim load_successed As Boolean = False
    Dim p1, p2, p3, p4 As PointF
    Private mySerialPort As New SerialPort
    Dim Data_Recive As String
    Dim point_num As Integer
    Private Delegate Sub UpdateFormDelegate()
    Private UpdateFormDelegate1 As UpdateFormDelegate

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
        gpanel = Panel1.CreateGraphics
        scale_factor(2) = 0.16777216
        scale_factor(3) = 0.2097152
        scale_factor(4) = 0.262144
        scale_factor(5) = 0.32768
        scale_factor(6) = 0.4096
        scale_factor(7) = 0.512
        scale_factor(8) = 0.64
        scale_factor(9) = 0.8
        scale_factor(10) = 1
        scale_factor(11) = 1.25
        scale_factor(12) = 1.5625
        scale_factor(13) = 1.953125
        scale_factor(14) = 2.44140625
        scale_factor(15) = 3.0517578125
        scale_factor(16) = 3.814697265625
        scale_factor(17) = 4.76837158203125
        scale_factor(18) = 5.9604644775390634
        load_successed = True

        Dim objIniFile As New clsIni("E:\WorldInfo.ini")
        objIniFile.WriteString("Software_Config", "COM", "fjsdfdskjfhsd")
        Label3.Text = objIniFile.GetString("Machine_Config", "OffSet", "")

        gpanel = Panel1.CreateGraphics
        AddHandler mySerialPort.DataReceived, AddressOf mySerialPort_DataReceived
        For Each AvailableSerialPorts As String In SerialPort.GetPortNames()
            ComboBox_AvailableSerialPorts.Items.Add(AvailableSerialPorts)
            SerialPort1.ReadTimeout = 2000
            ComboBox_AvailableSerialPorts.Text = AvailableSerialPorts
            Button_Connect.Visible = True

        Next
    End Sub
    
    Sub export_coords()
        For i1 = 1 To numofpoints(1)
            mssg.Text += i1.ToString + ": " + coords(1, i1).X.ToString + " " + coords(1, i1).Y.ToString + " "
        Next
    End Sub
    Sub get_the_coords_dxf()
        Dim i As Integer = 0
        Dim j As Integer
        Do
            i = i + 1
            numofpoly += 1
            numofpoints(i) = 0
            j = 1
            Do
                If str <> "EOF" Then
                    str = sr.ReadLine()
                End If
            Loop Until str = "POLYLINE" Or str = "EOF" Or sr Is Nothing
            Do
                '----------------------------------
                Do
                    If str <> "EOF" Then
                        str = sr.ReadLine()
                    End If
                Loop Until str = "VERTEX" Or str = "SEQEND" Or str = "EOF" Or sr Is Nothing
                If str = "VERTEX" Then
                    ''read coords(i,j)
                    'enter line until get code "10"
                    Do
                        str = sr.ReadLine()
                    Loop Until str = " 10"
                    str = sr.ReadLine()
                    If str.Length - str.IndexOf("."c) - 4 > 0 And str.IndexOf("."c) > 0 Then
                        str = str.Remove(str.IndexOf("."c) + 4, str.Length - str.IndexOf("."c) - 4)
                    End If
                    coords(i, j).X = CSng(str)
                    'tm = CInt(CSng(str) * 1000)
                    'coords(i, j).X = CSng(tm / 1000)
                    'get the Lbound X of the coords
                    If coords(i, j).X < minx Then
                        minx = coords(i, j).X
                    End If
                    'get the Ubound X of the coords
                    If coords(i, j).X > maxx Then
                        maxx = coords(i, j).X
                    End If
                    ''read coords(j+1)
                    'enter line until get code "20"
                    Do
                        str = sr.ReadLine()
                    Loop Until str = " 20"
                    str = sr.ReadLine()
                    If str.Length - str.IndexOf("."c) - 4 > 0 And str.IndexOf("."c) > 0 Then
                        str = str.Remove(str.IndexOf("."c) + 4, str.Length - str.IndexOf("."c) - 4)
                    End If
                    coords(i, j).Y = CSng(str)
                    'get the Lbound Y of the coords
                    If coords(i, j).Y < miny Then
                        miny = coords(i, j).Y
                    End If
                    'get the Ubound Y of the coords
                    If coords(i, j).Y > maxy Then
                        maxy = coords(i, j).Y
                    End If
                    j = j + 1
                End If
            Loop Until str = "SEQEND" Or str = "EOF" Or sr Is Nothing
            numofpoints(i) = j - 1
            '-------------------------------------------
        Loop Until str = "EOF" Or sr Is Nothing
        numofpoly -= 1
    End Sub
    Sub get_the_coords_AI()
        Dim i, j As Integer
        i = 0
        'str = ""
        str = sr.ReadLine
        Dim substr As String = ""
        While str <> "%%EOF" And sr IsNot Nothing
            i += 1
            j = 1

            While str <> "f" And str <> "%%EOF" And sr IsNot Nothing
                'str = sr.ReadLine
                If str.Length > 2 Then
                    substr = str.Substring(str.Length - 2)

                End If
                If substr = " m" Or substr = " l" Or substr = " L" Then

                    substr = str.Remove(str.IndexOf(" "c), str.Length - str.IndexOf(" "c))
                    coords(i, j).X = CSng(substr)
                    str = str.Remove(0, str.IndexOf(" "c) + 1)
                    substr = str.Remove(str.IndexOf(" "c), str.Length - str.IndexOf(" "c))
                    coords(i, j).Y = CSng(substr)
                    Label1.Text = substr
                    str = str.Remove(0, str.Length - 1)
                    j += 1
                ElseIf substr = " c" Or substr = " C" Then
                    'first direction point p1 is the previous point of coordinate system
                    p1.X = coords(i, j - 1).X
                    p1.Y = coords(i, j - 1).Y
                    'p2,p3,p4 is read from the line
                    'p2
                    substr = str.Remove(str.IndexOf(" "c), str.Length - str.IndexOf(" "c))
                    p2.X = CSng(substr)
                    str = str.Remove(0, str.IndexOf(" "c) + 1)
                    substr = str.Remove(str.IndexOf(" "c), str.Length - str.IndexOf(" "c))
                    p2.Y = CSng(substr)
                    str = str.Remove(0, str.IndexOf(" "c) + 1)
                    'p3
                    substr = str.Remove(str.IndexOf(" "c), str.Length - str.IndexOf(" "c))
                    p3.X = CSng(substr)
                    str = str.Remove(0, str.IndexOf(" "c) + 1)
                    substr = str.Remove(str.IndexOf(" "c), str.Length - str.IndexOf(" "c))
                    p3.Y = CSng(substr)
                    str = str.Remove(0, str.IndexOf(" "c) + 1)
                    'p4
                    substr = str.Remove(str.IndexOf(" "c), str.Length - str.IndexOf(" "c))
                    p4.X = CSng(substr)
                    str = str.Remove(0, str.IndexOf(" "c) + 1)
                    substr = str.Remove(str.IndexOf(" "c), str.Length - str.IndexOf(" "c))
                    p4.Y = CSng(substr)
                    str = str.Remove(0, str.IndexOf(" "c) + 1)
                    'B=P0*(1-t)^3 + 3*P1*t*(1-t)^2 + 3*P2*(1-t)*t^2 + P3*t^3
                    'gpanel.DrawBezier(myPen, p1, p2, p3, p4)
                    extrude_bezierpath(i, j)

                    j += point_num
                End If
                str = sr.ReadLine

            End While
            numofpoints(i) = j - 1
            'check tình trạng trùng lặp của 2 đỉnh cuối
            
            If str <> "%%EOF" Then
                str = sr.ReadLine
            End If
        End While
        numofpoly = i - 1

    End Sub
    Sub extrude_bezierpath(ByVal i As Integer, ByVal j As Integer)
        Dim k As Integer
        Dim coord_temp(0 To 100) As PointF
        k = j
        Dim seg_length As Double = 0
        point_num = 60
        Dim s As Integer = point_num
        For t = 1 To s
            coord_temp(t).X = CSng(p1.X * (1 - t / s) ^ 3 + 3 * p2.X * (t / s) * (1 - t / s) ^ 2 + 3 * p3.X * (1 - t / s) * (t / s) ^ 2 + p4.X * (t / s) ^ 3)
            coord_temp(t).Y = CSng(p1.Y * (1 - t / s) ^ 3 + 3 * p2.Y * (t / s) * (1 - t / s) ^ 2 + 3 * p3.Y * (1 - t / s) * (t / s) ^ 2 + p4.Y * (t / s) ^ 3)
        Next
        'calculate segment length
        For t = 2 To s
            seg_length += len(coord_temp(t - 1), coord_temp(t))
        Next

        Label11.Text += " " + Math.Round(seg_length, 2).ToString
        If seg_length < 10 Then
            point_num = 3
        ElseIf seg_length > 10 And seg_length < 20 Then
            point_num = 4
        Else
            point_num = CInt(seg_length / 14)
        End If
        'record points to coords array
        s = point_num
        For t = 1 To s
            coords(i, k).X = CSng(p1.X * (1 - t / s) ^ 3 + 3 * p2.X * (t / s) * (1 - t / s) ^ 2 + 3 * p3.X * (1 - t / s) * (t / s) ^ 2 + p4.X * (t / s) ^ 3)
            coords(i, k).Y = CSng(p1.Y * (1 - t / s) ^ 3 + 3 * p2.Y * (t / s) * (1 - t / s) ^ 2 + 3 * p3.Y * (1 - t / s) * (t / s) ^ 2 + p4.Y * (t / s) ^ 3)
            k += 1
        Next
    End Sub
    Sub find_boundary()
        For i1 = 1 To numofpoly
            For j1 = 1 To numofpoints(i1)
                If coords(i1, j1).X < minx Then
                    minx = coords(i1, j1).X
                End If
                If coords(i1, j1).X > maxx Then
                    maxx = coords(i1, j1).X
                End If
                If coords(i1, j1).Y < miny Then
                    miny = coords(i1, j1).Y
                End If
                If coords(i1, j1).Y > maxy Then
                    maxy = coords(i1, j1).Y
                End If
            Next
        Next
    End Sub
    Sub scale_coords()
        Dim sc_factor As Single = 2.83464
        Label9.Text = coords(1, 1).ToString
        For i1 = 1 To numofpoly
            For j1 = 1 To numofpoints(i1)
                If coords(i1, j1).X > minx Then
                    coords(i1, j1).X = (coords(i1, j1).X - minx * (1 - sc_factor)) / sc_factor
                End If
                If coords(i1, j1).Y > miny Then
                    coords(i1, j1).Y = (coords(i1, j1).Y - miny * (1 - sc_factor)) / sc_factor
                End If
            Next
        Next
        Label9.Text += " " + coords(1, 1).ToString
        Label10.Text += minx.ToString + " " + miny.ToString
    End Sub
    Sub initialize_file()
        Panel1.Refresh()
        gpanel.ResetTransform()
        gpanel.PageScale = 1
        offsX = 0
        offsY = 0
        scx = 10
        scy = 10
        dx = 0
        dy = 0
        dx_sc = 0
        dy_sc = 0
        pp.X = 0
        pp.Y = 0
        ToolStripButton14.Text = "Select"
        'reset str
        str = ""
        'reset minx & miny
        minx = Single.MaxValue
        miny = Single.MaxValue
        'reset maxx & maxy
        maxx = Single.MinValue
        maxy = Single.MinValue
        'clear array=0
        Array.Clear(numofpoints, 0, 100)
        'reset len_poly
        For i1 = 1 To 100
            poly_len(i1) = 0
        Next
        For i1 = 0 To 100
            For j1 = 0 To 500
                coords(i1, j1).X = 0
                coords(i1, j1).Y = 0
            Next
        Next
        'reset in_out size property
        For i1 = 1 To 100
            Is_Inside(i1) = False
        Next
        'reset numofpoly
        numofpoly = 0
        i = 0
        j = 0
        start_select = False
        is_dragging = False
        For i1 = 0 To 100
            is_selected(i1) = False
        Next
    End Sub
    Sub draw_poly_panel()
        Dim k As Integer
        Dim p As Integer
        myPen = New Pen(Brushes.White, 1 / scale_factor(scx))
        If numofpoly > 0 Then
            For k = 1 To numofpoly
                If is_selected(k) = True Then
                    myPen.Color = Color.Red
                Else
                    myPen.Color = Color.White
                End If
                For p = 1 To numofpoints(k) - 1
                    gpanel.DrawLine(myPen, coords(k, p), coords(k, p + 1))
                Next
            Next
        End If
        
    End Sub
    Sub extrude_poly()
        Dim p As PointF
        Dim fm As New StringFormat
        p.Y = maxy + 50
        p.X = minx
        Dim rect As Rectangle
        myPen.Color = Color.White
        myPen.Width = 1 / scale_factor(scx)
        Dim pb As PointF
        Dim ang_thres As Single = 20
        Dim k As Integer
        Dim f As New Font("Arial", text_size)
        Dim br As New SolidBrush(Color.Aqua)
        fm.Alignment = StringAlignment.Center
        If scx > 4 And scx < 16 Then
            For i1 = 1 To numofpoly
                fm.Alignment = StringAlignment.Center
                For j1 = 2 To numofpoints(i1)
                    If ang_out(i1, j1) >= ang_thres Or ang_out(i1, j1) <= -ang_thres Then
                        k = j1
                        pb.X = p.X
                        pb.Y = p.Y
                        Do
                            k -= 1
                        Loop Until ang_out(i1, k) >= ang_thres Or ang_out(i1, k) <= -ang_thres Or k = 1
                        'count segment lengh
                        p.X += CSng(len_seg(i1, k, j1))
                        'draw rect 
                        rect = New Rectangle(CInt(pb.X), CInt(pb.Y), CInt(len_seg(i1, k, j1)), 15)
                        gpanel.DrawRectangle(myPen, rect)
                        'draw angle 
                        myPen.Width = 1
                        gpanel.DrawString(CStr(ang_out(i1, j1)), f, Brushes.Brown, p.X, p.Y - text_size - 6, fm)
                        'draw segment lengh value
                        gpanel.DrawString(CStr(len_seg(i1, k, j1)), f, Brushes.White, (pb.X + p.X) / 2, p.Y + text_size - 6, fm)
                    End If
                Next
                'leng of whole segment
                p.X = minx + CSng(poly_len(i1))
                'draw whole segment
                rect = New Rectangle(CInt(minx), CInt(p.Y), CInt(poly_len(i1)), 15)
                gpanel.DrawRectangle(myPen, rect)
                fm.Alignment = StringAlignment.Near
                'draw lengh value of whole segment
                gpanel.DrawString(" lengh=" + CStr(poly_len(i1)), f, Brushes.Brown, p.X, p.Y + text_size - 2, fm)
                p.Y += 35
                p.X = minx
            Next
        End If
    End Sub
    Function len_seg(ByVal i1 As Integer, ByVal aa As Integer, ByVal bb As Integer) As Double
        len_seg = 0
        For j1 = aa To bb - 1
            len_seg += len(coords(i1, j1), coords(i1, j1 + 1))
        Next
    End Function
    Function len(ByVal p1 As PointF, ByVal p2 As PointF) As Double
        Dim tmp As Double = CDbl((p2.X - p1.X) * (p2.X - p1.X) + (p2.Y - p1.Y) * (p2.Y - p1.Y))
        len = tmp ^ 0.5
    End Function
    Sub draw_points()
        For i1 = 1 To numofpoly
            myPen.Width = 6
            myPen.Color = Color.Yellow
            gpanel.DrawLine(myPen, coords(i1, 1).X - 3, coords(i1, 1).Y, coords(i1, 1).X + 3, coords(i1, 1).Y)
            myPen.Color = Color.White
            For j1 = 2 To numofpoints(i1) - 1
                'If ang_out(i1, j1) >= 30 Or ang_out(i1, j1) <= -30 Then
                myPen.Width = 3
                myPen.Color = Color.Blue
                gpanel.DrawLine(myPen, coords(i1, j1).X - 1, coords(i1, j1).Y, coords(i1, j1).X + 1, coords(i1, j1).Y)
                'End If
            Next
            myPen.Width = 4
            myPen.Color = Color.Blue
        Next
    End Sub
    Sub angle_measure()
        Dim p11, p22, p33 As PointF
        For i1 = 1 To numofpoly
            For j1 = 1 To numofpoints(i1) - 1
                If j1 = 1 Then
                    'ni-1
                    p11 = coords(i1, numofpoints(i1) - 1)
                    'ni
                    p22 = coords(i1, 1)
                    'ji=1
                    p33 = coords(i1, 2)
                Else
                    p11 = coords(i1, j1 - 1)
                    'i, j + 1
                    p22 = coords(i1, j1)
                    'i,j+2
                    p33 = coords(i1, j1 + 1)
                End If
                'remove vertex
                For k = 1 To j1 - 1
                    coord(k) = coords(i1, k)
                Next
                For k = j1 + 1 To numofpoints(i1)
                    coord(k - 1) = coords(i1, k)
                Next
                coord(numofpoints(i1) - 1) = coord(1)
                'check vertex
                If area(coord, numofpoints(i1) - 1) > parea(i1) Then
                    ang_out(i1, j1) = -angle3points(p11, p22, p33)
                Else
                    ang_out(i1, j1) = angle3points(p11, p22, p33)
                End If
                If Is_Inside(i1) = True Then
                    ang_out(i1, j1) = -ang_out(i1, j1)
                End If
                ang_out(i1, numofpoints(i1)) = ang_out(i1, 1)
            Next
        Next
    End Sub
    Function angle3points(ByVal p1 As PointF, ByVal p2 As PointF, ByVal p3 As PointF) As Single
        Dim xba As Single
        Dim yba As Single
        Dim xbc As Single
        Dim ybc As Single
        Dim ba As Double
        Dim bc As Double
        Dim dot_prd As Single
        Dim arcc As Double
        Dim temp_ang As Double
        'xba
        xba = p1.X - p2.X
        'yba
        yba = p1.Y - p2.Y
        'xbc
        xbc = p3.X - p2.X
        'ybc
        ybc = p3.Y - p2.Y
        'ba
        str = CStr((xba * xba + yba * yba) ^ 0.5)
        If str.Length - str.IndexOf("."c) - 3 > 0 And str.IndexOf("."c) > 0 Then
            str = str.Remove(str.IndexOf("."c) + 3, str.Length - str.IndexOf("."c) - 3)
        End If
        ba = (xba * xba + yba * yba) ^ 0.5
        'bc
        str = CStr((xbc * xbc + ybc * ybc) ^ 0.5)
        If str.Length - str.IndexOf("."c) - 3 > 0 And str.IndexOf("."c) > 0 Then
            str = str.Remove(str.IndexOf("."c) + 3, str.Length - str.IndexOf("."c) - 3)
        End If
        bc = (xbc * xbc + ybc * ybc) ^ 0.5
        'dot_prd
        str = CStr(xba * xbc + yba * ybc)
        If str.Length - str.IndexOf("."c) - 3 > 0 And str.IndexOf("."c) > 0 Then
            str = str.Remove(str.IndexOf("."c) + 3, str.Length - str.IndexOf("."c) - 3)
        End If
        dot_prd = xba * xbc + yba * ybc
        'arcc
        arcc = dot_prd / (ba * bc)
        'angle3points
        str = CStr(180 - Math.Acos(arcc) * 180 / 3.14159265359)
        If str.Length - str.IndexOf("."c) - 3 > 0 And str.IndexOf("."c) > 0 Then
            str = str.Remove(str.IndexOf("."c) + 3, str.Length - str.IndexOf("."c) - 3)
        End If
        temp_ang = 180 - Math.Acos(arcc) * 180 / 3.14159265359
        temp_ang = Math.Round(temp_ang, 2)
        angle3points = CSng(temp_ang)
    End Function
    Sub len_measure()
        Dim tmpx, tmpy As Double
        For i1 = 1 To numofpoly
            'len between j,j+1
            poly_len(i1) = 0
            For j1 = 1 To numofpoints(i1) - 1
                tmpx = CDbl((coords(i1, j1).X - coords(i1, j1 + 1).X) * (coords(i1, j1).X - coords(i1, j1 + 1).X))
                tmpy = CDbl((coords(i1, j1).Y - coords(i1, j1 + 1).Y) * (coords(i1, j1).Y - coords(i1, j1 + 1).Y))
                poly_len(i1) += (tmpx + tmpy) ^ 0.5
            Next
            str = CStr(poly_len(i1))
            If str.Length - str.IndexOf("."c) - 3 > 0 And str.IndexOf("."c) > 0 Then
                str = str.Remove(str.IndexOf("."c) + 3, str.Length - str.IndexOf("."c) - 3)
            End If
            poly_len(i1) = CDbl(str)
        Next
    End Sub
    Sub area_and_direct()
        Dim s As Single
        For i1 = 1 To numofpoly
            right_direct(i1) = False
            s = 0
            For j1 = 1 To numofpoints(i1) - 1
                s = s + (coords(i1, j1).X - coords(i1, j1 + 1).X) * (coords(i1, j1).Y + coords(i1, j1 + 1).Y)
            Next
            If s < 0 Then
                parea(i1) = -s / 2
                right_direct(i1) = True
            Else
                parea(i1) = s / 2

            End If
        Next
    End Sub
    Function area(ByVal coord() As PointF, ByVal nv As Integer) As Single
        area = 0
        For j1 = 1 To nv - 1
            area += (coord(j1).X - coord(j1 + 1).X) * (coord(j1).Y + coord(j1 + 1).Y)
        Next
        If area < 0 Then
            area = -area / 2
        Else
            area = area / 2
        End If
    End Function
    Sub printtext_poly()
        Dim f As New Font("Arial", 2)
        Dim br As New SolidBrush(Color.Aqua)
        For i1 = 1 To numofpoly
            For j1 = 1 To numofpoints(i1) - 1
                'If ang_out(i1, j1) > 20 Or ang_out(i1, j1) < -20 Then
                gpanel.DrawString(" " + CStr(ang_out(i1, j1)), f, br, coords(i1, j1).X, coords(i1, j1).Y - 4)
                'End If
            Next
        Next

    End Sub
    Sub get_Lbound()
        For i1 = 1 To numofpoly
            Lboundp(i1) = coords(i1, 1)
            For j1 = 2 To numofpoints(i1)
                If coords(i1, j1).X < Lboundp(i1).X Then
                    Lboundp(i1).X = coords(i1, j1).X
                End If
                If coords(i1, j1).X < minx Then
                    minx = coords(i1, j1).X
                End If
                If coords(i1, j1).Y < Lboundp(i1).Y Then
                    Lboundp(i1).Y = coords(i1, j1).Y
                End If
                If coords(i1, j1).Y < miny Then
                    miny = coords(i1, j1).Y
                End If
            Next
        Next
    End Sub
    Sub get_Ubound()
        For i1 = 1 To numofpoly
            Uboundp(i1) = coords(i1, 1)
            For j1 = 2 To numofpoints(i1)
                If coords(i1, j1).X > Uboundp(i1).X Then
                    Uboundp(i1).X = coords(i1, j1).X
                End If
                If coords(i1, j1).X > maxx Then
                    maxx = coords(i1, j1).X
                End If
                If coords(i1, j1).Y > Uboundp(i1).Y Then
                    Uboundp(i1).Y = coords(i1, j1).Y
                End If
                If coords(i1, j1).Y > maxy Then
                    maxy = coords(i1, j1).Y
                End If
            Next
        Next
    End Sub
    Sub in_outside()
        For i1 = 1 To numofpoly
            For j1 = 1 To numofpoly
                If i1 <> j1 And Lboundp(i1).X > Lboundp(j1).X And Uboundp(i1).X < Uboundp(j1).X And Lboundp(i1).Y > Lboundp(j1).Y And Uboundp(i1).Y < Uboundp(j1).Y Then
                    Is_Inside(i1) = True
                End If
            Next
        Next
    End Sub

   
    Private Sub transform_coords()
        Dim i1 As Integer
        Dim j1 As Integer
        If numofpoly > 0 Then
            get_Lbound()
            get_Ubound()
            dx = (minx + maxx) / 2
            dy = (miny + maxy) / 2
            'gpanel.TranslateTransform(dx - CSng(Panel1.Size.Width) / 2, dy - CSng(Panel1.Size.Height) / 2)
            For i1 = 1 To numofpoly
                For j1 = 1 To numofpoints(i1)
                    coords(i1, j1).Y = maxy + miny - coords(i1, j1).Y
                Next
            Next
            'Re-check min-max of Y coords
            miny = Single.MaxValue
            maxy = Single.MinValue
            For i1 = 1 To numofpoly
                For j1 = 1 To numofpoints(i1)
                    If coords(i1, j1).Y < miny Then
                        miny = coords(i1, j1).Y
                    ElseIf coords(i1, j1).Y > maxy Then
                        maxy = coords(i1, j1).Y
                    End If
                Next
            Next
        End If
        dx = (minx + maxx) / 2 - CSng(Panel1.Size.Width / 2)
        dy = (miny + maxy) / 2 - CSng(Panel1.Size.Height / 2)

    End Sub
    Sub movecoords(ByVal dx As Single, ByVal dy As Single)
        Dim i, j As Integer
        For i = 1 To numofpoly
            For j = 1 To numofpoints(i)
                coords(i, j).X -= dx
                coords(i, j).Y -= dy
            Next
        Next
    End Sub

    Sub corect_direct()
        Dim j1 As Integer
        Dim tmp As PointF
        For i1 = 1 To numofpoly
            If (Is_Inside(i1) = False And right_direct(i1) = False) Or (Is_Inside(i1) = True And right_direct(i1) = True) Then
                For j1 = 2 To CInt(numofpoints(i1) / 2)
                    tmp = coords(i1, j1)
                    coords(i1, j1) = coords(i1, numofpoints(i1) - j1 + 1)
                    coords(i1, numofpoints(i1) - j1 + 1) = tmp
                Next
            End If
        Next
    End Sub
 
    Private Sub LeftButt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If numofpoly > 0 Then
            dx -= 40
            movecoords(40, 0)
            Me.Panel1.Refresh()
            draw_poly_panel()
            draw_points()
            printtext_poly()
            extrude_poly()
            'export_coords()
        End If
    End Sub
    Private Sub RightButt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If numofpoly > 0 Then
            dx += 40
            movecoords(-40, 0)
            Me.Panel1.Refresh()
            draw_poly_panel()
            draw_points()
            printtext_poly()
            extrude_poly()
            'export_coords()
        End If
    End Sub
    Public Sub move_down()
        If numofpoly > 0 Then
            dy -= 40
            Me.gpanel.TranslateTransform(0, 40)
            Me.Panel1.Refresh()
            draw_poly_panel()
            draw_points()
            printtext_poly()
            extrude_poly()
        End If
    End Sub

    Public Sub move_up()
        If numofpoly > 0 Then
            dy += 40
            Me.gpanel.TranslateTransform(0, -40)
            Me.Panel1.Refresh()
            draw_poly_panel()
            draw_points()
            printtext_poly()
            extrude_poly()
        End If
    End Sub
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim f As New Font("Arial", text_size - 4)
        Dim br As New SolidBrush(Color.Aqua)
        If bend_display = False And numofpoly > 0 Then
            For i1 = 1 To numofpoly
                For j1 = 1 To numofpoints(i1) - 1
                    If ang_out(i1, j1) > 2 Or ang_out(i1, j1) < -2 Then
                        gpanel.DrawString(" " + CStr(ang_out(i1, j1)), f, br, coords(i1, j1).X, coords(i1, j1).Y - 4)
                    End If
                Next
            Next
            bend_display = True

        Else
            Me.Panel1.Refresh()
            draw_poly_panel()
            draw_points()
            printtext_poly()
            extrude_poly()
            bend_display = False

        End If
    End Sub

    Function points_aligned(ByVal p1 As PointF, ByVal p2 As Point, ByVal p3 As PointF) As Boolean
        Dim xba As Single
        Dim yba As Single
        Dim xbc As Single
        Dim ybc As Single
        Dim ba As Double
        Dim bc As Double
        Dim dot_prd As Single
        Dim arcc As Double
        Dim temp_ang As Double
        Dim epsilon As Single = 5
        points_aligned = False
        'xba
        xba = p1.X - p2.X
        'yba
        yba = p1.Y - p2.Y
        'xbc
        xbc = p3.X - p2.X
        'ybc
        ybc = p3.Y - p2.Y
        'ba
        ba = (xba * xba + yba * yba) ^ 0.5
        'bc        
        bc = (xbc * xbc + ybc * ybc) ^ 0.5
        'dot_prd        
        dot_prd = xba * xbc + yba * ybc
        'arcc
        arcc = dot_prd / (ba * bc)
        '
        temp_ang = Math.Acos(arcc) * 180 / 3.14159265359
        If temp_ang > 180 - epsilon Then
            points_aligned = True
        End If
    End Function

    Private Sub Form1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        
    End Sub

    Private Sub Set_lengh_per_pulse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub


    Sub long_process(ByVal i2 As Integer)
        Dim excess As Double = 5
        Dim seg As Double
        Dim j1, j2 As Integer
        Dim seg1 As Double
        'tiền xử lí
        reset_bendid(i2)
        BackUpCoords(i2)
        pre_process_coords(i2)
        'xử lí
        Dim id As Integer
        Dim quantity As Integer
        For j1 = 2 To numofpoints(i2) + 1
            If start_bend_of(j1) = 0 Then
                fwd(len(coords(i2, j1 - 1), coords(i2, j1)))
            Else
                id = start_bend_of(j1)
                quantity = seg_bend(j1)
                'get length from id->j1-1
                seg = get_length_from(i2, id, j1 - 1)
                seg1 = seg
                'mssg.Text += " " + seg.ToString
                'fwd(offs-seg)
                fwd(offs - seg)
                If id <= numofpoints(i2) - 1 Then
                    bend(ang_out(i2, id))
                End If
                quantity -= 1
                'bend hết các đoạn trong j1-1,j1
                For j2 = 1 To quantity
                    seg = len(coords(i2, id), coords(i2, id + 1))
                    fwd(seg)
                    id += 1
                    If id <= numofpoints(i2) - 1 Then
                        bend(ang_out(i2, id))
                    End If
                Next
                'fwd đoạn thừa nếu ko phải ở đỉnh sau cuối
                If j1 <= numofpoints(i2) Then
                    seg = get_length_from(i2, id, j1)
                    'mssg.Text += " " + seg.ToString
                    fwd(seg - offs)
                End If
            End If
            If j1 <= numofpoints(i2) Then
                cut(ang_out(i2, j1))
            End If
        Next
        'đỉnh cần bend cuối cùng nằm tại vị trí bend
        RestoreCoords(i2)

    End Sub
    Function get_length_from(ByVal k1 As Integer, ByVal k2 As Integer, ByVal k3 As Integer) As Double
        Dim tmp As Double = 0
        For j3 = k2 To k3 - 1
            tmp += len(coords(k1, j3), coords(k1, j3 + 1))
        Next
        Return tmp
    End Function
    Sub pre_process_coords(ByVal i2 As Integer)
        Dim j1 As Integer
        Dim j2 As Integer
        Dim seg As Double
        coords(i2, numofpoints(i2 + 1)).X += CSng(offs)
        coords(i2, numofpoints(i2 + 1)).Y += CSng(offs)
        j1 = 2
        Do While j1 <= numofpoints(i2)      'out-most loop
            j2 = j1
            seg = len(coords(i2, j1), coords(i2, j1 + 1))

            Do While seg < offs And j2 <= numofpoints(i2)
                j2 += 1
                seg += len(coords(i2, j2), coords(i2, j2 + 1))
            Loop
            'sau khi xử lí thì seg>offs
            If start_bend_of(j2 + 1) = 0 Then
                start_bend_of(j2 + 1) = j1
                seg_bend(j2 + 1) = 1
            ElseIf start_bend_of(j2 + 1) > 0 Then
                seg_bend(j2 + 1) += 1
            End If
            j1 += 1
        Loop 'out-most loop
        'restore coords

    End Sub
    Sub reset_bendid(ByVal i1 As Integer)
        For j1 = 0 To 500
            start_bend_of(j1) = 0
            seg_bend(j1) = 0
        Next
    End Sub
    Sub BackUpCoords(ByVal i1 As Integer)
        For j1 = 0 To 500
            coordsBackUp(j1) = coords(i1, j1)
        Next
    End Sub
    Sub RestoreCoords(ByVal i1 As Integer)
        For j1 = 0 To 500
            coords(i1, j1) = coordsBackUp(j1)
        Next
    End Sub
    Sub short_process(ByVal i2 As Integer)
        Dim seg As Double
        Dim j1 As Integer
        seg = 0
        cut(1)
        For j1 = 1 To numofpoints(i2) - 1
            seg += len(coords(i2, j1), coords(i2, j1 + 1))
            fwd(len(coords(i2, j1), coords(i2, j1 + 1)))
            cut(ang_out(i2, j1 + 1))
        Next
        seg -= len(coords(i2, 1), coords(i2, 2))
        fwd(offs - seg)
        bend(ang_out(i2, 2))
        For j1 = 2 To numofpoints(i2) - 1
            fwd(len(coords(i2, j1), coords(i2, j1 + 1)))
            bend(ang_out(i2, j1 + 1))
        Next
    End Sub
    Sub fwd(ByVal leng As Double)
        Dim str As String
        str = leng.ToString
        If str.Length - str.IndexOf("."c) - 3 > 0 And str.IndexOf("."c) > 0 Then
            str = str.Remove(str.IndexOf("."c) + 3, str.Length - str.IndexOf("."c) - 3)
        End If
        leng = CDbl(str)
        mss2send.Text += "F" + leng.ToString + "-"
    End Sub
    Sub prv(ByVal leng As Double)
        Dim str As String
        str = leng.ToString
        If str.Length - str.IndexOf("."c) - 3 > 0 And str.IndexOf("."c) > 0 Then
            str = str.Remove(str.IndexOf("."c) + 3, str.Length - str.IndexOf("."c) - 3)
        End If
        leng = CDbl(str)
        mss2send.Text += "P" + leng.ToString + "-"
    End Sub
    Sub cut(ByVal angle As Double)
        Dim str As String
        str = angle.ToString
        If str.Length - str.IndexOf("."c) - 3 > 0 And str.IndexOf("."c) > 0 Then
            str = str.Remove(str.IndexOf("."c) + 3, str.Length - str.IndexOf("."c) - 3)
        End If
        angle = CDbl(str)
        mss2send.Text += "C" + angle.ToString + "-"
    End Sub
    Sub bend(ByVal angle As Double)
        Dim str As String
        str = angle.ToString
        If str.Length - str.IndexOf("."c) - 3 > 0 And str.IndexOf("."c) > 0 Then
            str = str.Remove(str.IndexOf("."c) + 3, str.Length - str.IndexOf("."c) - 3)
        End If
        angle = CDbl(str)
        mss2send.Text += "B" + angle.ToString + "-"
    End Sub

    Private Sub Button_Open_port_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub Form1_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseWheel
        Dim p As PointF
        p = Panel1.PointToClient(MousePosition)
        w = p.X
        h = p.Y

        If e.Delta > 0 Then

            Trace.WriteLine("Scrolled up!")
            If numofpoly > 0 And scx < 18 Then
                scx = scx + 1
                scy = scy + 1
                gpanel.ScaleTransform(1.25, 1.25)
                Label1.Text = scx.ToString
                dx_sc = (w - w * 5 / 4) / CSng(1.25)
                dy_sc = (h - h * 5 / 4) / CSng(1.25)
                gpanel.TranslateTransform(dx_sc, dy_sc)
                offsX = offsX / CSng(1.25) + dx_sc 'cthuc dung
                offsY = offsY / CSng(1.25) + dy_sc ' cthuc dung
                Label5.Text = offsX.ToString + " " + offsY.ToString
                Label3.Text = p.X.ToString + " " + p.Y.ToString
                Label6.Text = "sc" + dx_sc.ToString + "  " + dy_sc.ToString
                Label2.Text = "w h " + w.ToString + " " + h.ToString
                Label1.Text = gpanel.ClipBounds.Height.ToString
                Me.Panel1.Refresh()
                draw_poly_panel()
                draw_points()
                printtext_poly()
                extrude_poly()

            End If
        Else
            Trace.WriteLine("Scrolled down!")
            If numofpoly > 0 And scx > 2 Then
                scx -= 1
                scy -= 1
                Label1.Text = scx.ToString
                Me.gpanel.ScaleTransform(4 / 5, 4 / 5)
                dx_sc = (w - w * 4 / 5) / CSng(0.8)
                dy_sc = (h - h * 4 / 5) / CSng(0.8)
                gpanel.TranslateTransform(dx_sc, dy_sc)
                offsX = offsX / CSng(0.8) + dx_sc
                offsY = offsY / CSng(0.8) + dy_sc
                Label2.Text = w.ToString + " " + h.ToString
                Label6.Text = "dxsc" + dx_sc.ToString + "  " + dy_sc.ToString
                Label1.Text = gpanel.ClipBounds.Height.ToString
                Me.Panel1.Refresh()
                draw_poly_panel()
                draw_points()
                printtext_poly()
                extrude_poly()
            End If
        End If
    End Sub
    Private Sub Panel1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseMove
        Dim p As Point
        If is_dragging And numofpoly > 0 Then
            p.X = CInt((e.X - start_drag.X) / scale_factor(scx))
            p.Y = CInt((e.Y - start_drag.Y) / scale_factor(scx))
            start_drag.X = e.X
            start_drag.Y = e.Y
            gpanel.TranslateTransform(p.X, p.Y)
            offsX += p.X
            offsY += p.Y
            Panel1.Refresh()
            draw_poly_panel()
            draw_points()
            printtext_poly()
            extrude_poly()
        End If
    End Sub
    Private Sub Panel1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseUp
        is_dragging = False
        Label1.Text = gpanel.ClipBounds.Height.ToString
    End Sub



    Private Sub Butt_extrude_previous_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Butt_extrude_previous.Click
        If Processing = False Then
            If IsNumeric(Text_Ext_Previous) Then
                prv(CDbl(Text_Ext_Previous.Text))
            End If
        End If
    End Sub

    Private Sub Butt_extrude_fwd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Butt_extrude_fwd.Click
        If Processing = False Then
            If IsNumeric(Text_Ext_Forward) Then
                prv(CDbl(Text_Ext_Forward.Text))
            End If
        End If
    End Sub
 

    Private Sub Panel1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseClick
        Dim p As PointF
        Dim checked As Boolean = False
        Dim epsilon As Double = 1.5
        Dim len1, len2, len3 As Double
        myPen.Color = Color.White
        myPen.Width = 1 / scale_factor(scx)
        If start_select = True And e.Button = Windows.Forms.MouseButtons.Left Then
            p = Panel1.PointToClient(MousePosition)
            p.X = p.X / scale_factor(scx) - offsX
            p.Y = p.Y / scale_factor(scy) - offsY
            Label4.Text = e.X.ToString + " " + e.Y.ToString
            Label3.Text = p.X.ToString + " " + p.Y.ToString
            For i1 = 1 To numofpoly
                For j1 = 1 To numofpoints(i1) - 1
                    len1 = len(coords(i1, j1), p)
                    len2 = len(p, coords(i1, j1 + 1))
                    len3 = len(coords(i1, j1), coords(i1, j1 + 1))
                    If len1 + len2 - len3 <= epsilon Then
                        gpanel.DrawLine(myPen, p.X - 2, p.Y, p.X + 2, p.Y)
                        is_selected(i1) = Not is_selected(i1)
                        Label7.Text = "selected " + is_selected(i1).ToString

                    End If
                Next
            Next
            draw_poly_panel()
        End If
    End Sub
    Function ischecked(ByVal i As Integer, ByVal j As Integer, ByVal p As PointF) As Boolean
        ischecked = False
        Dim epsilon As Double = 1.5
        Dim len1, len2, len3 As Double
        len1 = len(coords(i, j), p)
        len2 = len(p, coords(i, j + 1))
        len3 = len(coords(i, j), coords(i, j + 1))
        If len1 + len2 - len3 <= epsilon Then
            ischecked = True
        End If
    End Function
    Private Sub Form1_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint

        draw_poly_panel()

        draw_points()
        printtext_poly()
        extrude_poly()
    End Sub



    Private Sub Panel1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then
            start_drag = e.Location
            is_dragging = True
            Label1.Text = e.X.ToString + " " + e.Y.ToString
        End If
    End Sub


    Sub set_all_selected()
        For i1 = 1 To numofpoly
            is_selected(i1) = True
        Next
    End Sub
    Sub set_all_unselected()
        For i1 = 1 To numofpoly
            is_selected(i1) = False
        Next
    End Sub

    Private Sub Button_Connect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Connect.Click
        If (Button_Connect.Text = "Connect") Then
            mySerialPort.BaudRate = CInt("250000")
            mySerialPort.PortName = CStr(ComboBox_AvailableSerialPorts.SelectedItem)
            If mySerialPort.IsOpen = False Then
                mySerialPort.Open()
            End If
            Button_Connect.Text = "Disconnect"
            Button_Connect.ForeColor = Color.Red
        Else
            If mySerialPort.IsOpen = True Then
                mySerialPort.Close()
            End If
            Button_Connect.Text = "Connect"
            Button_Connect.ForeColor = Color.Blue
        End If
    End Sub
    Private Sub mySerialPort_DataReceived(ByVal sender As Object, ByVal e As SerialDataReceivedEventArgs)
        'Handles serial port data received events
        UpdateFormDelegate1 = New UpdateFormDelegate(AddressOf UpdateDisplay)
        Data_Recive = mySerialPort.ReadTo("/n") 'read data from the buffer
        Me.Invoke(UpdateFormDelegate1) 'call the delegate
    End Sub
    Private Sub UpdateDisplay()
        TextBox1.Text = Data_Recive.Trim()
        If (CStr(Data_Recive.Trim().CompareTo(mssg.Text.Trim())) = "0") Then
            txt_lich_su_1.ForeColor = Color.LimeGreen
            txt_lich_su_1.Text = "==>  " + DateTime.Now.ToString("dd-MM-yyyy  |  hh:mm:ss tt") + " : Send Data Successfull" + vbNewLine + txt_lich_su_1.Text
            Dim ans As DialogResult = MessageBox.Show("Data Send Succesfull. Do you want to run the machine ???", "Run Machine", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If ans = DialogResult.Yes Then
                mySerialPort.Write("R")
                txt_lich_su_1.Text = "==>  " + DateTime.Now.ToString("dd-MM-yyyy  |  hh:mm:ss tt") + " : Machine is Running" + vbNewLine + txt_lich_su_1.Text
            ElseIf ans = DialogResult.No Then
                mySerialPort.Write("S")
                txt_lich_su_1.Text = "==>  " + DateTime.Now.ToString("dd-MM-yyyy  |  hh:mm:ss tt") + " : Machine is Not Running" + vbNewLine + txt_lich_su_1.Text
                'Do nothing
            End If
        ElseIf (CStr(Data_Recive.Trim().CompareTo("OK")) = "0") Then
            txt_lich_su_1.Text = "==>  " + DateTime.Now.ToString("dd-MM-yyyy  |  hh:mm:ss tt") + " : Machine Run Done" + vbNewLine + txt_lich_su_1.Text
            Button_Send.Enabled = True
        Else
            txt_lich_su_1.Text = "==>  " + DateTime.Now.ToString("dd-MM-yyyy  |  hh:mm:ss tt") + " : Send Data Error" + vbNewLine + txt_lich_su_1.Text
            txt_lich_su_1.ForeColor = Color.Red
            Button_Send.Enabled = True
        End If


    End Sub

    Private Sub Button_Send_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Send.Click
        Dim Length_Data As Int32 = mssg.Text.Length()
        Dim Data_0 As String
        Data_0 = mssg.Text.Trim()
        Static Dim First_Push As Int32 = 0
        mySerialPort.WriteLine(Data_0 + "/n")
        ' Label8.Text = Data_0
        Button_Send.Enabled = False

        Label9.Text = CStr(Length_Data)
        txt_lich_su_1.Text = "==>  " + DateTime.Now.ToString("dd-MM-yyyy  |  hh:mm:ss tt") + " : Send Data to Arduino" + vbNewLine + txt_lich_su_1.Text
    End Sub

    Private Sub ToolStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ToolStrip1.ItemClicked

    End Sub

    Private Sub ToolStripLabel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel2.Click

    End Sub

    Private Sub OpenToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripButton.Click
        Dim OpenfileDialog1 As New OpenFileDialog
        Dim extension_str As Char
        If OpenfileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            sr = New System.IO.StreamReader(OpenfileDialog1.FileName)
            Dim fi As New IO.FileInfo(OpenfileDialog1.FileName)
            extension_str = fi.ToString.Last
            Label7.Text = extension_str
            initialize_file()

            If extension_str = "f" Then
                get_the_coords_dxf()
                transform_coords()
            ElseIf extension_str = "i" Then
                get_the_coords_AI()
                find_boundary()
                Label3.Text = numofpoly.ToString
                Label8.Text = ang_out(1, 2).ToString + " len " + len(coords(1, 1), coords(1, 6)).ToString
                transform_coords()
                scale_coords()
                find_boundary()
            End If
            trim_excess_vertex()
            movecoords(dx, dy)
            draw_poly_panel()
            in_outside()
            area_and_direct()
            corect_direct()
            len_measure()
            angle_measure()
            Text_scale.Text = "100"
            mssg.Text = ""
            export_coords()
            draw_points()
            printtext_poly()
            extrude_poly()
            End If
    End Sub
    Sub trim_excess_vertex()
        Label11.Text = ""
        For i1 = 1 To numofpoly
            If coords(i1, 1) = coords(i1, 2) Then
                For j1 = 1 To numofpoints(i1) - 1
                    coords(i1, j1) = coords(i1, j1 + 1)
                Next
                numofpoints(i1) -= 1
                Label11.Text += "trimaa" + i1.ToString
            End If
            If coords(i1, numofpoints(i1)) = coords(i1, numofpoints(i1) - 1) Then
                numofpoints(i1) -= 1
                Label11.Text += "trim" + i1.ToString
            End If
        Next
    End Sub
    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        If numofpoly > 0 Then
            scx = 10
            scy = 10
            offsX = 0
            offsY = 0
            dx_sc = 0
            dy_sc = 0
            Panel1.Refresh()
            gpanel.ResetTransform()
            Me.Panel1.Refresh()
            draw_poly_panel()
            draw_points()
            printtext_poly()
            extrude_poly()
            'export_coords()
        End If
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ToolStripButton14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton14.Click
        If start_select = True Then
            start_select = False
            ToolStripButton14.Text = "Select"
        Else
            start_select = True
            ToolStripButton14.Text = "Selecting..."
        End If
    End Sub

    Private Sub ToolStripButton15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton15.Click
        set_all_unselected()


    End Sub

    Private Sub ToolStripButton17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton17.Click
        set_all_selected()
        Panel1.Refresh()
        draw_poly_panel()
        draw_points()
        printtext_poly()
    End Sub

    Private Sub Text_scale_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Text_scale.TextChanged

    End Sub

    Private Sub ToolStripButton_Setup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton_Setup.Click
        Form_Setup.Text = "SET UP"
        Form_Setup.Show()
    End Sub

    Private Sub ToolStripButton19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton19.Click
        offs = 324
        Processing = True
        mss2send.Text = ""
        For i1 = 1 To numofpoly
            If poly_len(i1) - len(coords(i1, 1), coords(i1, 2)) > offs Then
                long_process(i1)
            Else
                short_process(i1)
            End If
        Next
        'len_measure()
        Processing = False
    End Sub
End Class





