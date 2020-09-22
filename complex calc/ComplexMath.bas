Attribute VB_Name = "ComplexMath"
Public Type Complex 'this data type holds the variable (a+bi)
    Real As Double
    Imag As Double 'sqr(-1)
End Type
'n1 + n2
Public Function C_ADD(ByRef n1 As Complex, ByRef n2 As Complex) As Complex
C_ADD.Real = n1.Real + n2.Real
C_ADD.Imag = n1.Imag + n2.Imag
End Function
'n1-n2
Public Function C_sub(ByRef n1 As Complex, ByRef n2 As Complex) As Complex
C_sub.Real = n1.Real - n2.Real
C_sub.Imag = n1.Imag - n2.Imag
End Function
'finds the square root of a complex and or negative number
Public Function C_SQR(ByRef n1 As Complex) As Complex
C_SQR.Imag = Sqr(-0.5 * n1.Real + 0.5 * Sqr(n1.Imag ^ 2 + n1.Real ^ 2))
C_SQR.Real = Sqr(n1.Real + C_SQR.Imag ^ 2)
End Function

Public Function MakeComplex(NReal As Double, Optional NImag As Double = 0) As Complex
MakeComplex.Real = NReal
MakeComplex.Imag = NImag
End Function
'used in devision
Public Function conj(ByRef n1 As Complex) As Complex
conj.Real = n1.Real
conj.Imag = -n1.Imag
End Function
'n1*n2
Public Function Cmult(ByRef n1 As Complex, ByRef n2 As Complex) As Complex
Cmult.Real = n1.Real * n2.Real - n1.Imag * n2.Imag
Cmult.Imag = n1.Real * n2.Imag + n2.Real * n1.Imag
End Function
'n1/n2
Public Function CDev(ByRef n1 As Complex, ByRef n2 As Complex) As Complex
With CDev
    .Real = Cmult(conj(n2), n1).Real / Cmult(conj(n2), n2).Real
    .Imag = Cmult(conj(n2), n1).Imag / Cmult(conj(n2), n2).Real
End With
End Function

Public Function GenerateString(n1 As Complex) As String
Dim nstring As String
nstring = n1.Real
If n1.Imag <> 0 Then
    If n1.Real <> 0 Then
        nstring = n1.Real & "+" & n1.Imag & "i"
    Else
        nstring = n1.Imag & "i"
    End If
End If
GenerateString = nstring
End Function
