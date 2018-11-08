Option Strict On
Imports System
Imports System.IO.Ports
Imports System.Drawing.Drawing2D
Public Class Form1
    Private ppoint As Point
    Dim poly_len(100) As Double
    Dim coords(100, 500) As PointF
    Dim ang_out(100, 500) As Single
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
    Dim coord(0 To 500) As PointF
    Dim right_direct(0 To 100) As Boolean
    Dim bend_display As Boolean = False
    Dim offs As Double
    Dim start_bend_of(0 To 500) As Integer
    Dim seg_bend(0 To 500) As Integer
    Dim coordsBackUp(0 To 500) As PointF
    Dim Poly_scale As Double
    Dim Processing As Boolean
    Dim offsX, offsY As Single
    Dim w, h As Single
    Dim scale_factor(0 To 20) As Single
    Dim is_selected(0 To 100) As Boolean
    Dim start_drag, end_drag As Point
    Dim is_dragging As Boolean = False
    Dim first_drag As Boolean = True
    Dim load_successed As Boolean = False
    Private Sub Open_butt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Open_butt.Click
        Dim OpenfileDialog1 As New OpenFileDialog
        If OpenfileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            sr = New System.IO.StreamReader(OpenfileDialog1.FileName)
            initialize_file()
            get_the_coords()
            transform_coords()
            movecoords(dx, dy)
            draw_poly_panel()
            in_outside()
            area_and_direct()
            corect_direct()
            len_measure()
            angle_measure()
            Poly_scale = 1
            Text_scale.Text = "100"
            mssg.Text = ""
            export_coords()
            draw_points()
            printtext_poly()
            'extrude_poly()
        End If

    End Sub
    Sub export_coords()
        For i1 = 1 To numofpoints(1)
            mssg.Text += i1.ToString + ": " + coords(1, i1).X.ToString + " " + coords(1, i1).Y.ToString + " "
        Next
    End Sub
    Sub get_the_coords()
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
        Butt_select.Text = "Select"
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
                gpanel.DrawString(" lengh=" + CStr(poly_len(i1)), f, Brushes.Brown, p.X, p.Y + text_size - 6, fm)
                p.Y += 35
                p.X = minx
            Next
        End If
    End Sub
    Function len_seg(ByVal i1 As Integer, ByVal aa As Integer, ByVal bb As Integer) As Double
        len_seg = 0
        For j1 = aa To bb - 1
            len_seg += Poly_scale * len(coords(i1, j1), coords(i1, j1 + 1))
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
                If ang_out(i1, j1) >= 30 Or ang_out(i1, j1) <= -30 Then
                    myPen.Width = 6
                    myPen.Color = Color.Blue
                    gpanel.DrawLine(myPen, coords(i1, j1).X - 3, coords(i1, j1).Y, coords(i1, j1).X + 3, coords(i1, j1).Y)
                End If
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
                poly_len(i1) += Poly_scale * (tmpx + tmpy) ^ 0.5
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
        Dim f As New Font("Arial", text_size)
        Dim br As New SolidBrush(Color.Aqua)
        For i1 = 1 To numofpoly
            For j1 = 1 To numofpoints(i1) - 1
                If ang_out(i1, j1) > 20 Or ang_out(i1, j1) < -20 Then
                    gpanel.DrawString(" " + CStr(ang_out(i1, j1)), f, br, coords(i1, j1).X, coords(i1, j1).Y - 4)
                End If
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
    Private Sub Zinbutt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Zinbutt.Click
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
            'extrude_poly()
            'export_coords()
        End If
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
    Private Sub Start_points_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Start_points.Click

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
            'extrude_poly()
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
            'extrude_poly()
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

    Private Sub Start_bending_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Start_bending.Click
        offs = 25
        Processing = True
        mssg.Text = poly_len(1).ToString
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
                mssg.Text += " " + seg.ToString
                'fwd(offs-seg)
                fwd(offs - seg)
                If id <= numofpoints(i2) - 1 Then
                    bend(id)
                End If
                quantity -= 1
                'bend hết các đoạn trong j1-1,j1
                For j2 = 1 To quantity
                    seg = len(coords(i2, id), coords(i2, id + 1))
                    fwd(seg)
                    id += 1
                    If id <= numofpoints(i2) - 1 Then
                        bend(id)
                    End If
                Next
                'fwd đoạn thừa nếu ko phải ở đỉnh sau cuối
                If j1 <= numofpoints(i2) Then
                    seg = get_length_from(i2, id, j1)
                    mssg.Text += " " + seg.ToString
                    fwd(seg - offs)
                End If
            End If
            If j1 <= numofpoints(i2) Then
                cut(j1)
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
            cut(j1 + 1)
        Next
        seg -= len(coords(i2, 1), coords(i2, 2))
        fwd(offs - seg)
        bend(2)
        For j1 = 2 To numofpoints(i2) - 1
            fwd(len(coords(i2, j1), coords(i2, j1 + 1)))
            bend(j1 + 1)
        Next
    End Sub
    Sub fwd(ByVal leng As Double)
        mssg.Text += " fwd" + leng.ToString

    End Sub
    Sub prv(ByVal leng As Double)

    End Sub
    Sub cut(ByVal angle As Double)
        mssg.Text += " cut" + angle.ToString
    End Sub
    Sub bend(ByVal angle As Double)
        mssg.Text += " bend" + angle.ToString
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
                'extrude_poly()

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
                'extrude_poly()
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
        End If
    End Sub
    Private Sub Panel1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseUp
        is_dragging = False
        Label1.Text = gpanel.ClipBounds.Height.ToString
    End Sub
    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Setup.Click
        Dim form2 As New Form
        Dim Butt_set_offset As New Button
        form2.Text = "SET UP"
        Butt_set_offset.Text = "SET OFFSET"

        Butt_set_offset.Location = New Point(20, 20)
        Butt_set_offset.Size = New Size(70, 20)
        Butt_set_offset.FlatStyle = New FlatStyle
        Butt_set_offset.FlatAppearance.BorderSize = 1


        form2.Controls.Add(Butt_set_offset)
        form2.Show()

    End Sub

    Private Sub Butt_set_Scale_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Butt_set_Scale.Click
        If Processing = False Then
            If IsNumeric(Text_scale.Text) Then
                Poly_scale = CDbl(Text_scale.Text) / 100
            End If
        End If
        mssg.Text = Poly_scale.ToString
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
    Private Sub Butt_select_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Butt_select.Click
        If start_select = True Then
            start_select = False
            Butt_select.Text = "Select"
        Else
            start_select = True
            Butt_select.Text = "Selecting..."
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
        'Panel1.Refresh()
        draw_poly_panel()
        'gpanel.DrawLine(myPen, 150, 150, 150, 155)
        draw_points()
        printtext_poly()
        'extrude_poly()
    End Sub

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

    End Sub

    Private Sub Panel1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then
            start_drag = e.Location
            is_dragging = True
            Label1.Text = e.X.ToString + " " + e.Y.ToString
        End If
    End Sub
    
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True And start_select = True Then
            set_all_selected()
        ElseIf CheckBox1.Checked = False And start_select = True Then
            set_all_unselected()
        End If
        Panel1.Refresh()
        draw_poly_panel()
        draw_points()
        printtext_poly()
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
End Class



