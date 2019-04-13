Public Class colors
    Public Structure RGB
        Private _r As Byte
        Private _g As Byte
        Private _b As Byte

        Public Sub New(r As Byte, g As Byte, b As Byte)
            Me._r = r
            Me._g = g
            Me._b = b
        End Sub

        Public Property R() As Byte
            Get
                Return Me._r
            End Get
            Set(value As Byte)
                Me._r = value
            End Set
        End Property

        Public Property G() As Byte
            Get
                Return Me._g
            End Get
            Set(value As Byte)
                Me._g = value
            End Set
        End Property

        Public Property B() As Byte
            Get
                Return Me._b
            End Get
            Set(value As Byte)
                Me._b = value
            End Set
        End Property

        Public Overloads Function Equals(rgb As RGB) As Boolean
            Return (Me.R = rgb.R) AndAlso (Me.G = rgb.G) AndAlso (Me.B = rgb.B)
        End Function
    End Structure

    Public Structure HSV
        Private _h As Double
        Private _s As Double
        Private _v As Double

        Public Sub New(h As Double, s As Double, v As Double)
            Me._h = h
            Me._s = s
            Me._v = v
        End Sub

        Public Property H() As Double
            Get
                Return Me._h
            End Get
            Set(value As Double)
                Me._h = value
            End Set
        End Property

        Public Property S() As Double
            Get
                Return Me._s
            End Get
            Set(value As Double)
                Me._s = value
            End Set
        End Property

        Public Property V() As Double
            Get
                Return Me._v
            End Get
            Set(value As Double)
                Me._v = value
            End Set
        End Property

        Public Overloads Function Equals(hsv As HSV) As Boolean
            Return (Me.H = hsv.H) AndAlso (Me.S = hsv.S) AndAlso (Me.V = hsv.V)
        End Function
    End Structure

    Public Shared Function HSVToRGB(hsv As HSV) As RGB
        Dim r As Double = 0, g As Double = 0, b As Double = 0

        If hsv.S = 0 Then
            r = hsv.V
            g = hsv.V
            b = hsv.V
        Else
            Dim i As Integer
            Dim f As Double, p As Double, q As Double, t As Double

            If hsv.H = 360 Then
                hsv.H = 0
            Else
                hsv.H = hsv.H / 60
            End If

            i = CInt(Math.Truncate(hsv.H))
            f = hsv.H - i

            p = hsv.V * (1.0 - hsv.S)
            q = hsv.V * (1.0 - (hsv.S * f))
            t = hsv.V * (1.0 - (hsv.S * (1.0 - f)))

            Select Case i
                Case 0
                    r = hsv.V
                    g = t
                    b = p
                    Exit Select

                Case 1
                    r = q
                    g = hsv.V
                    b = p
                    Exit Select

                Case 2
                    r = p
                    g = hsv.V
                    b = t
                    Exit Select

                Case 3
                    r = p
                    g = q
                    b = hsv.V
                    Exit Select

                Case 4
                    r = t
                    g = p
                    b = hsv.V
                    Exit Select
                Case Else

                    r = hsv.V
                    g = p
                    b = q
                    Exit Select

            End Select
        End If

        Return New RGB(CByte(Math.Truncate(r * 255)), CByte(Math.Truncate(g * 255)), CByte(Math.Truncate(b * 255)))
    End Function

End Class