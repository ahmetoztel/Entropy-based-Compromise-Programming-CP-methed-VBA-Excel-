Attribute VB_Name = "Module1"
Sub compromise_programming()
x = InputBox("Alternatif Say�s� Giriniz")
y = InputBox("Kriter Say�s� Giriniz")
Dim etop() As Double
ReDim etop(y) As Double

Dim renk() As Double
ReDim renk(x, y) As Double
Dim We() As Double
ReDim We(y) As Double

Dim ext()
Dim AA() As Double
Dim BB() As Double
ReDim AA(y) As Double
ReDim BB(y) As Double
Dim R As Long
Dim c As Long
ReDim ext(y)
Dim fp() As Double
ReDim fp(y) As Double

Dim fn() As Double
ReDim fn(y) As Double

Dim S() As Double
ReDim S(x) As Double









For c = 2 To y + 1
R = 2
AA(c - 1) = Cells(R, c)
BB(c - 1) = Cells(R, c)
For R = 2 To x + 1
If Cells(R, c) > AA(c - 1) Then
AA(c - 1) = Cells(R, c)
End If
If Cells(R, c) < BB(c - 1) Then
BB(c - 1) = Cells(R, c)
End If
Next R
Next c


For c = 2 To y + 1
For R = 2 To x + 1
If Cells(R, c) < 0 Then
Cells(R, c) = (AA(c - 1) - BB(c - 1)) * (Cells(R, c) - Int(BB(c - 1))) / ((Sgn(AA(c - 1)) * Int(Abs(AA(c - 1)))) - (Sgn(BB(c - 1)) * Int(Abs(BB(c - 1)))))

End If

Next R
Next c
For c = 2 To y + 1
For R = 2 To x + 1
If Cells(R, c) = 0 Then
Cells(R, c) = 0.000001
End If
Next R
Next c



R = 2
 c = 1
For R = 2 To x + 1

Cells(R, c).Select
Cells(R, c).Value = "A" & CStr(R - 1)

Next R
R = 1


For c = 2 To y + 1

Cells(R, c).Select
Cells(R, c).Value = "C" & CStr(c - 1)
geri_:
ext(c - 1) = InputBox("Bu kriter i�in ideal de�er minimum ise 1 maximum ise 2 giriniz.")
If ext(c - 1) < 1 Or ext(c - 1) > 2 Then
MsgBox "Yanl�� de�er girdiniz"
GoTo geri_

End If
Next c

Cells(x + 2, 1).Value = "We"

For c = 2 To y + 1
etop(c - 1) = 0

For R = 2 To x + 1
n = Cells(R, c)


etop(c - 1) = etop(c - 1) + n
Next R



Next c

For c = 2 To y + 1

For R = 2 To x + 1

renk(R - 1, c - 1) = Cells(R, c) / etop(c - 1)




Next R


Next c

P = 0
For c = 2 To y + 1
t = 0
For R = 2 To x + 1


k = -renk(R - 1, c - 1) * Log(renk(R - 1, c - 1)) / Log(x)
t = t + k
Next R
P = P + 1 - t
We(c - 1) = t



Next c



For c = 2 To y + 1
We(c - 1) = (1 - We(c - 1)) / P

Cells(x + 2, c).Value = We(c - 1)

Next c
Cells(x + 2, 1).Select
MsgBox "We ile ba�layan sat�rda herbir kriter i�in Entropy y�ntemiyle hesaplanm�� olan a��rl�k de�erleri yer almaktad�r"

For j = 1 To y

If ext(j) = 2 Then
fp(j) = AA(j)
fn(j) = BB(j)
Else
fp(j) = BB(j)
fn(j) = AA(j)
End If
Next j

'Her bir alternatifin her bir kritere g�re faydas�
Dim Ut() As Double
ReDim Ut(x, y) As Double
Dim Ut�oklu() As Double
ReDim Ut�oklu(x) As Double
Dim Ups() As Double
ReDim Ups(x) As Double
Dim sonup As Double
Dim ��z�m() As Double
ReDim ��z�m(x) As Double
Cells(x + 4, 1) = "Tekli fayda fonksiyon de�erleri "
Cells(1, y + 4) = "p=1 i�in  �oklu fayda fonksiyon de�erleri "
Cells(x + 2, y + 4) = "p=sonsuz i�in  �oklu fayda fonksiyon de�erleri "
Cells(1, y + 9) = "Optimal ��z�m de�erleri "



For i = 1 To x
toplam = 0
sonup = 0
For j = 1 To y
Ut(i, j) = (fp(j) - Cells(i + 1, j + 1)) / (fp(j) - fn(j))
Ut�oklu(i) = We(j) * Ut(i, j)
Ups(i) = We(j) * Ut(i, j)
If sonup <= Ups(i) Then
sonup = Ups(i)
End If

toplam = toplam + Ut�oklu(i)
Cells(i + x + 4, 1 + j) = Ut(i, j)
Next j
Ups(i) = sonup

Ut�oklu(i) = toplam
Cells(i + 1, y + 4) = Ut�oklu(i)
Cells(x + i + 2, y + 4) = Ups(i)
��z�m(i) = 0.5 * (Ut�oklu(i) + Ups(i))
Cells(i + 1, y + 9) = ��z�m(i)

Next i

'��z�mlerin k���kten b�y��e s�ralanmas�

Dim s�rra() As Double
ReDim s�rra(x) As Double
Dim tut As Integer
Cells(1, y + 12) = "S�ralama "
Cells(1, y + 13) = "De�erler "
For i = 1 To x
kasa = ��z�m(i)
 For j = 1 To x
 If kasa >= ��z�m(j) Then
 kasa = ��z�m(j)
 tut = j
 End If
 Next j
 s�rra(i) = kasa
 Cells(i + 1, y + 12) = "A" & Str(tut)
 Cells(i + 1, y + 13) = Str(kasa)
 ��z�m(tut) = 5000000
 
 Next i
 
 





End Sub
