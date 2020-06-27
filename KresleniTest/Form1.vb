Imports System.Drawing.Drawing2D

Public Class Form1

    Public Structure tObject
        Dim oCenter As Point
        Dim oPosun As Point
        Dim intSize As Integer
        Dim pPen As Pen
        Dim blnLeftToRight As Boolean
        Dim blnDown As Boolean
        Dim intRotationSpeed As Integer
        Dim blnRotateRight As Boolean
        Dim sinAngle As Single
        'body pro hvezdu
        Dim aPoints() As Point
        Dim intPoints As Integer
        Dim intObjectType As Integer
    End Structure

    Const constintMaxPoints As Integer = 7
    Const constintMaxSpeeed As Integer = 6
    Const constintMaxRotation As Integer = 3
    Const constintMinObjects As Integer = 15
    Const constintMaxObjects As Integer = 30
    Const constintMinSize As Integer = 30
    Const constintMaxSize As Integer = 60
    Const constblnChangeBackground As Boolean = False

    Dim blnIsInitialized As Boolean = False

    Dim pts() As Point
    Dim lPts As Long

    Dim aObjects() As tObject
    Dim intObjects As Integer

    Dim blnMouseDown As Boolean = False
    Dim sinAngle As Single = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim i As Long

        Randomize()

        intObjects = CInt(constintMinObjects + Int(Rnd() * constintMaxObjects))

        SetStyle(ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.OptimizedDoubleBuffer, True)

        SetStyle(ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.OptimizedDoubleBuffer, True)

        SetPosition(True)

        Refresh()

        Timer1.Start()
        Timer2.Stop()

    End Sub

    Private Sub SetPosition(Optional ByVal blnStart As Boolean = False)
        Dim i As Integer

        If blnStart Then

            ReDim aObjects(intObjects)
            For i = 0 To intObjects

                With aObjects(i)

                    If (i <= (intObjects \ 4)) Then
                        .intObjectType = 1 'CInt(1 + Int(Rnd() * 2))
                    ElseIf (i > (intObjects \ 4)) And (i <= 2 * (intObjects \ 4)) Then
                        .intObjectType = 2 'CInt(1 + Int(Rnd() * 2))
                    ElseIf (i > 2 * (intObjects \ 4)) And (i <= 3 * (intObjects \ 4)) Then
                        .intObjectType = 3 'CInt(1 + Int(Rnd() * 2))
                    Else
                        .intObjectType = 4 'CInt(1 + Int(Rnd() * 2))
                    End If
                    '.intObjectType = 1

                    .oCenter = New Point(CInt(200 + Int(Rnd() * (Me.ClientSize.Width - 400))), CInt(200 + Int(Rnd() * (Me.ClientSize.Height - 400))))
                    .blnLeftToRight = IIf(CInt(1 + Int(Rnd() * 2)) = 1, True, False)
                    .blnDown = IIf(CInt(1 + Int(Rnd() * 2)) = 1, True, False)
                    .blnRotateRight = IIf(CInt(1 + Int(Rnd() * 2)) = 1, True, False)
                    .intRotationSpeed = CInt(1 + Int(Rnd() * constintMaxRotation))
                    .pPen = New Pen(Color.Red)
                    .pPen.Width = CInt(2 + Int(Rnd() * 10))
                    .pPen.DashStyle = DashStyle.Solid
                    .pPen.Color = Color.FromArgb(CInt(50 + Rnd() * 200), CInt(50 + Rnd() * 200), CInt(50 + Rnd() * 200))
                    .sinAngle = 0
                    .intSize = CInt(constintMinSize + Int(Rnd() * constintMaxSize))

                    If .blnLeftToRight Then

                        If .blnDown Then

                            .oPosun = New Point(CInt(2 + Int(Rnd() * constintMaxSpeeed)), CInt(2 + Int(Rnd() * constintMaxSpeeed)))

                        Else

                            .oPosun = New Point(CInt(2 + Int(Rnd() * constintMaxSpeeed)), -1 * CInt(2 + Int(Rnd() * constintMaxSpeeed)))

                        End If

                    Else

                        If .blnDown Then

                            .oPosun = New Point(-1 * CInt(2 + Int(Rnd() * constintMaxSpeeed)), CInt(2 + Int(Rnd() * constintMaxSpeeed)))

                        Else

                            .oPosun = New Point(-1 * CInt(2 + Int(Rnd() * constintMaxSpeeed)), -1 * CInt(2 + Int(Rnd() * constintMaxSpeeed)))

                        End If

                    End If


                End With

                Call Hvezda(aObjects(i))

            Next

            blnIsInitialized = True

        Else

            For i = 0 To intObjects

                With aObjects(i)

                    If .blnLeftToRight Then

                        If .blnDown Then

                            If ((.oCenter.X + .intSize) >= Me.ClientSize.Width) Then

                                .oPosun.X = -1 * CInt(2 + Int(Rnd() * constintMaxSpeeed))
                                .blnLeftToRight = False
                                .intRotationSpeed = CInt(1 + Int(Rnd() * constintMaxRotation))
                                .pPen.Width = CInt(2 + Int(Rnd() * 10))
                                .pPen.Color = Color.FromArgb(CInt(50 + Rnd() * 200), CInt(50 + Rnd() * 200), CInt(50 + Rnd() * 200))
                                '.intSize = CInt(30 + Int(Rnd() * 70))
                                If constblnChangeBackground Then Me.BackColor = Color.FromArgb(CInt(50 + Rnd() * 100), CInt(50 + Rnd() * 100), CInt(50 + Rnd() * 100))
                                '.intObjectType = .intObjectType + 1
                                'If .intObjectType > 4 Then .intObjectType = 1

                            End If

                            If ((.oCenter.Y + .intSize) >= Me.ClientSize.Height) Then

                                .oPosun.Y = -1 * CInt(2 + Int(Rnd() * constintMaxSpeeed))
                                .blnDown = False
                                .intRotationSpeed = CInt(1 + Int(Rnd() * constintMaxRotation))
                                .pPen.Width = CInt(2 + Int(Rnd() * 10))
                                .pPen.Color = Color.FromArgb(CInt(50 + Rnd() * 200), CInt(50 + Rnd() * 200), CInt(50 + Rnd() * 200))
                                If constblnChangeBackground Then Me.BackColor = Color.FromArgb(CInt(50 + Rnd() * 100), CInt(50 + Rnd() * 100), CInt(50 + Rnd() * 100))
                                '.intObjectType = .intObjectType + 1
                                'If .intObjectType > 4 Then .intObjectType = 1

                            End If

                        Else 'blnDown

                            If ((.oCenter.X + .intSize) >= Me.ClientSize.Width) Then

                                .oPosun.X = -1 * CInt(2 + Int(Rnd() * constintMaxSpeeed))
                                .blnLeftToRight = False
                                .intRotationSpeed = CInt(1 + Int(Rnd() * constintMaxRotation))
                                .pPen.Width = CInt(2 + Int(Rnd() * 10))
                                .pPen.Color = Color.FromArgb(CInt(50 + Rnd() * 200), CInt(50 + Rnd() * 200), CInt(50 + Rnd() * 200))
                                If constblnChangeBackground Then Me.BackColor = Color.FromArgb(CInt(50 + Rnd() * 100), CInt(50 + Rnd() * 100), CInt(50 + Rnd() * 100))
                                '.intObjectType = .intObjectType + 1
                                'If .intObjectType > 4 Then .intObjectType = 1

                            End If

                            If ((.oCenter.Y - .intSize) <= 0) Then

                                .oPosun.Y = CInt(2 + Int(Rnd() * constintMaxSpeeed))
                                .blnDown = True
                                .intRotationSpeed = CInt(1 + Int(Rnd() * constintMaxRotation))
                                .pPen.Width = CInt(2 + Int(Rnd() * 10))
                                .pPen.Color = Color.FromArgb(CInt(50 + Rnd() * 200), CInt(50 + Rnd() * 200), CInt(50 + Rnd() * 200))
                                If constblnChangeBackground Then Me.BackColor = Color.FromArgb(CInt(50 + Rnd() * 100), CInt(50 + Rnd() * 100), CInt(50 + Rnd() * 100))
                                '.intObjectType = .intObjectType + 1
                                'If .intObjectType > 4 Then .intObjectType = 1

                            End If

                        End If 'blnDown

                    Else

                        If .blnDown Then

                            If ((.oCenter.X - .intSize) <= 0) Then

                                .oPosun.X = CInt(2 + Int(Rnd() * constintMaxSpeeed))
                                .blnLeftToRight = True
                                .intRotationSpeed = CInt(1 + Int(Rnd() * constintMaxRotation))
                                .pPen.Width = CInt(2 + Int(Rnd() * 10))
                                .pPen.Color = Color.FromArgb(CInt(50 + Rnd() * 200), CInt(50 + Rnd() * 200), CInt(50 + Rnd() * 200))
                                If constblnChangeBackground Then Me.BackColor = Color.FromArgb(CInt(50 + Rnd() * 100), CInt(50 + Rnd() * 100), CInt(50 + Rnd() * 100))
                                '.intObjectType = .intObjectType + 1
                                'If .intObjectType > 4 Then .intObjectType = 1

                            End If

                            If ((.oCenter.Y + .intSize) >= Me.ClientSize.Height) Then

                                .oPosun.Y = -1 * CInt(2 + Int(Rnd() * constintMaxSpeeed))
                                .blnDown = False
                                .intRotationSpeed = CInt(1 + Int(Rnd() * constintMaxRotation))
                                .pPen.Width = CInt(2 + Int(Rnd() * 10))
                                .pPen.Color = Color.FromArgb(CInt(50 + Rnd() * 200), CInt(50 + Rnd() * 200), CInt(50 + Rnd() * 200))
                                If constblnChangeBackground Then Me.BackColor = Color.FromArgb(CInt(50 + Rnd() * 100), CInt(50 + Rnd() * 100), CInt(50 + Rnd() * 100))
                                '.intObjectType = .intObjectType + 1
                                'If .intObjectType > 4 Then .intObjectType = 1

                            End If

                        Else 'blnDown

                            If ((.oCenter.X - .intSize) <= 0) Then

                                .oPosun.X = CInt(2 + Int(Rnd() * constintMaxSpeeed))
                                .blnLeftToRight = True
                                .intRotationSpeed = CInt(1 + Int(Rnd() * constintMaxRotation))
                                .pPen.Width = CInt(2 + Int(Rnd() * 10))
                                .pPen.Color = Color.FromArgb(CInt(50 + Rnd() * 200), CInt(50 + Rnd() * 200), CInt(50 + Rnd() * 200))
                                If constblnChangeBackground Then Me.BackColor = Color.FromArgb(CInt(50 + Rnd() * 100), CInt(50 + Rnd() * 100), CInt(50 + Rnd() * 100))
                                '.intObjectType = .intObjectType + 1
                                'If .intObjectType > 4 Then .intObjectType = 1

                            End If

                            If ((.oCenter.Y - .intSize) <= 0) Then

                                .oPosun.Y = CInt(2 + Int(Rnd() * constintMaxSpeeed))
                                .blnDown = True
                                .intRotationSpeed = CInt(1 + Int(Rnd() * constintMaxRotation))
                                .pPen.Width = CInt(2 + Int(Rnd() * 10))
                                .pPen.Color = Color.FromArgb(CInt(50 + Rnd() * 200), CInt(50 + Rnd() * 200), CInt(50 + Rnd() * 200))
                                If constblnChangeBackground Then Me.BackColor = Color.FromArgb(CInt(50 + Rnd() * 100), CInt(50 + Rnd() * 100), CInt(50 + Rnd() * 100))
                                '.intObjectType = .intObjectType + 1
                                'If .intObjectType > 4 Then .intObjectType = 1

                            End If

                        End If 'blnDown

                    End If

                    .oCenter.X = .oCenter.X + .oPosun.X
                    .oCenter.Y = .oCenter.Y + .oPosun.Y

                    Call Hvezda(aObjects(i))
                    If (.intObjectType = 2) Then .pPen.Width = 2

                End With

            Next

        End If

    End Sub

    Private Sub Form1_Paint(ByVal sender As Object, ByVal e As PaintEventArgs) Handles MyBase.Paint
        'Dim i As Long, lHvezdy As Long

        'lHvezdy = (10 + Int(Rnd() * 20))

        'For i = 0 To lHvezdy - 1

        '    Dim byVolba As Byte
        '    byVolba = Int(Rnd() * 2) + 1

        '    Call Hvezda(byVolba, e)

        'Next i

        If blnIsInitialized Then Call Rotace(e)

    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        Me.Refresh()
        Timer1.Stop()
        Timer2.Start()

        If Not (e.Modifiers = Keys.Alt And e.KeyCode = Keys.F4) Then
            e.SuppressKeyPress = True
        Else
            Application.Exit()
        End If

    End Sub

    Public Sub Rotace(ByVal e As PaintEventArgs)
        Dim myMatrix As New Matrix
        Dim i As Integer

        SetPosition()

        For i = 0 To intObjects

            With aObjects(i)

                If .blnRotateRight Then
                    myMatrix.RotateAt(.sinAngle, .oCenter, MatrixOrder.Append)
                Else
                    myMatrix.RotateAt(- .sinAngle, .oCenter, MatrixOrder.Append)
                End If
                e.Graphics.Transform = myMatrix

                Select Case .intObjectType
                    Case 1
                        e.Graphics.DrawRectangle(.pPen, .oCenter.X - .intSize, .oCenter.Y - .intSize, (2 * .intSize), (2 * .intSize))
                        e.Graphics.DrawLine(.pPen, .oCenter.X, .oCenter.Y - 10, .oCenter.X, .oCenter.Y + 10)
                        If .blnRotateRight Then
                            e.Graphics.DrawEllipse(.pPen, .oCenter.X - (3 * (.intSize \ 4)), .oCenter.Y - (.intSize \ 2), .intSize \ 2, .intSize \ 4)
                            e.Graphics.DrawEllipse(.pPen, .oCenter.X + (.intSize \ 4), .oCenter.Y - (.intSize \ 2), .intSize \ 2, .intSize \ 4)
                            e.Graphics.DrawArc(.pPen, .oCenter.X - (.intSize \ 4), .oCenter.Y + (.intSize \ 2), (.intSize \ 2), (.intSize \ 2), 0, -180)
                        Else
                            e.Graphics.DrawEllipse(.pPen, .oCenter.X - (.intSize \ 2), .oCenter.Y - (.intSize \ 2), .intSize \ 4, .intSize \ 2)
                            e.Graphics.DrawEllipse(.pPen, .oCenter.X + (.intSize \ 4), .oCenter.Y - (.intSize \ 2), .intSize \ 4, .intSize \ 2)
                            e.Graphics.DrawArc(.pPen, .oCenter.X - (.intSize \ 4), .oCenter.Y, (.intSize \ 2), (.intSize \ 2), 0, 180)
                        End If

                    Case 2
                        e.Graphics.DrawPolygon(.pPen, .aPoints)

                    Case 3
                        e.Graphics.DrawEllipse(.pPen, .oCenter.X - .intSize, .oCenter.Y - .intSize, .intSize, (2 * .intSize))

                    Case Else
                        e.Graphics.DrawRectangle(.pPen, .oCenter.X - .intSize, .oCenter.Y - .intSize, (2 * .intSize), .intSize)

                End Select
                If .intObjectType = 1 Then



                Else


                End If

                If .blnRotateRight Then
                    myMatrix.RotateAt(- .sinAngle, .oCenter, MatrixOrder.Append)
                Else
                    myMatrix.RotateAt(.sinAngle, .oCenter, MatrixOrder.Append)
                End If
                e.Graphics.Transform = myMatrix

                .sinAngle = .sinAngle + .intRotationSpeed '1
                If .sinAngle >= 359 Then .sinAngle = 0

            End With

        Next

    End Sub

    Public Sub Hvezda(ByRef oObject As tObject)

        With oObject

            If Not blnIsInitialized Then

                Do
                    .intPoints = CInt(5 + Int(Rnd() * constintMaxPoints))
                Loop While .intPoints < 5

                If (.intPoints Mod 2) = 0 Then
                    If .intPoints + 1 >= constintMaxPoints Then
                        .intPoints = .intPoints - 1
                    Else
                        .intPoints = .intPoints + 1
                    End If
                End If

                ReDim .aPoints(.intPoints)

            End If

            Dim theta As Double
            Dim dtheta As Double = Math.PI * 0.8

            Dim j As Long = 0

            For i As Integer = 0 To .intPoints

                j = (j + .intPoints \ 2) Mod .intPoints
                theta = (j * (Math.PI * 2) / .intPoints)

                .aPoints(i).X = CInt(.oCenter.X + .intSize * Math.Cos(theta))
                .aPoints(i).Y = CInt(.oCenter.Y + .intSize * Math.Sin(theta))

            Next i

        End With

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        Timer2.Stop()
        Refresh()

    End Sub

    Private Sub Form1_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown

        blnMouseDown = True
        Refresh()
        Timer1.Stop()
        Timer2.Start()

    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        Timer1.Start()

    End Sub

    Private Sub Form1_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove

        If blnMouseDown Then

            Refresh()
            Timer1.Stop()
            Timer2.Start()

        End If
    End Sub

    Private Sub Form1_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
        blnMouseDown = False
    End Sub

    Private Sub Form1_MouseWheel(sender As Object, e As MouseEventArgs) Handles Me.MouseWheel

        Refresh()
        Timer1.Stop()
        Timer2.Start()

    End Sub
End Class