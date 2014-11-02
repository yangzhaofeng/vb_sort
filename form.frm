Public Sub sort(list() As Integer, ByVal min As Integer, ByVal max As Integer)
    Dim med_value As Integer
    Dim hi As Integer
    Dim lo As Integer
    Dim i As Integer

' If the list has no more than CutOff elements,
' finish it off with SelectionSort.
    If max <= min Then Return

' Pick the dividing value.
    i = Int((max - min + 1) * Rnd + min)
    med_value = list(i)

' Swap it to the front.
    list(i) = list(min)

    lo = min
    hi = max
    Do
' Look down from hi for a value < med_value.
        Do While list(hi) >= med_value
            hi = hi - 1
            If hi <= lo Then Exit Do
        Loop
        If hi <= lo Then
            list(lo) = med_value
            Exit Do
        End If

' Swap the lo and hi values.
        list(lo) = list(hi)

' Look up from lo for a value >= med_value.
        lo = lo + 1
        Do While list(lo) < med_value
            lo = lo + 1
            If lo >= hi Then Exit Do
        Loop
        If lo >= hi Then
            lo = hi
            list(hi) = med_value
            Exit Do
        End If

' Swap the lo and hi values.
        list(hi) = list(lo)
    Loop

' Sort the two sublists.
    QuicksortInt list(), min, lo - 1
    QuicksortInt list(), lo + 1, max
End Sub
Private Sub QuicksortSingle(list() As Single, ByVal min As Integer, ByVal max As Integer)
    Dim med_value As Single
    Dim hi As Integer
    Dim lo As Integer
    Dim i As Integer

' If the list has no more than CutOff elements,
' finish it off with SelectionSort.
    If max <= min Then Exit Sub

' Pick the dividing value.
    i = Int((max - min + 1) * Rnd + min)
    med_value = list(i)

' Swap it to the front.
    list(i) = list(min)

    lo = min
    hi = max
    Do
' Look down from hi for a value < med_value.
        Do While list(hi) >= med_value
            hi = hi - 1
            If hi <= lo Then Exit Do
        Loop
        If hi <= lo Then
            list(lo) = med_value
            Exit Do
        End If

' Swap the lo and hi values.
        list(lo) = list(hi)

' Look up from lo for a value >= med_value.
        lo = lo + 1
        Do While list(lo) < med_value
            lo = lo + 1
            If lo >= hi Then Exit Do
        Loop
        If lo >= hi Then
            lo = hi
            list(hi) = med_value
            Exit Do
        End If

' Swap the lo and hi values.
list(hi) = list(lo)
Loop

' Sort the two sublists.
QuicksortSingle list(), min, lo - 1
QuicksortSingle list(), lo + 1, max
End Sub

Private Sub QuicksortDouble(list() As Double, ByVal min As Integer, ByVal max As Integer)
Dim med_value As Double
Dim hi As Integer
Dim lo As Integer
Dim i As Integer

' If the list has no more than CutOff elements,
' finish it off with SelectionSort.
If max <= min Then Exit Sub

' Pick the dividing value.
i = Int((max - min + 1) * Rnd + min)
med_value = list(i)

' Swap it to the front.
list(i) = list(min)

lo = min
hi = max
Do
' Look down from hi for a value < med_value.
Do While list(hi) >= med_value
hi = hi - 1
If hi <= lo Then Exit Do
Loop
If hi <= lo Then
list(lo) = med_value
Exit Do
End If

' Swap the lo and hi values.
list(lo) = list(hi)

' Look up from lo for a value >= med_value.
lo = lo + 1
Do While list(lo) < med_value
lo = lo + 1
If lo >= hi Then Exit Do
Loop
If lo >= hi Then
lo = hi
list(hi) = med_value
Exit Do
End If

' Swap the lo and hi values.
list(hi) = list(lo)
Loop

' Sort the two sublists.
QuicksortDouble list(), min, lo - 1
QuicksortDouble list(), lo + 1, max
End Sub
Private Sub QuicksortString(list() As String, ByVal min As Integer, ByVal max As Integer)
Dim med_value As String
Dim hi As Integer
Dim lo As Integer
Dim i As Integer

' If the list has no more than CutOff elements,
' finish it off with SelectionSort.
If max <= min Then Exit Sub

' Pick the dividing value.
i = Int((max - min + 1) * Rnd + min)
med_value = list(i)

' Swap it to the front.
list(i) = list(min)

lo = min
hi = max
Do
' Look down from hi for a value < med_value.
Do While list(hi) >= med_value
hi = hi - 1
If hi <= lo Then Exit Do
Loop
If hi <= lo Then
list(lo) = med_value
Exit Do
End If

' Swap the lo and hi values.
list(lo) = list(hi)

' Look up from lo for a value >= med_value.
lo = lo + 1
Do While list(lo) < med_value
lo = lo + 1
If lo >= hi Then Exit Do
Loop
If lo >= hi Then
lo = hi
list(hi) = med_value
Exit Do
End If

' Swap the lo and hi values.
list(hi) = list(lo)
Loop

' Sort the two sublists.
QuicksortString list(), min, lo - 1
QuicksortString list(), lo + 1, max
End Sub
Private Sub QuicksortVariant(list() As Variant, ByVal min As Integer, ByVal max As Integer)
Dim med_value As Variant
Dim hi As Integer
Dim lo As Integer
Dim i As Integer

' If the list has no more than CutOff elements,
' finish it off with SelectionSort.
If max <= min Then Exit Sub

' Pick the dividing value.
i = Int((max - min + 1) * Rnd + min)
med_value = list(i)

' Swap it to the front.
list(i) = list(min)

lo = min
hi = max
Do
' Look down from hi for a value < med_value.
Do While list(hi) >= med_value
hi = hi - 1
If hi <= lo Then Exit Do
Loop
If hi <= lo Then
list(lo) = med_value
Exit Do
End If

' Swap the lo and hi values.
list(lo) = list(hi)

' Look up from lo for a value >= med_value.
lo = lo + 1
Do While list(lo) < med_value
lo = lo + 1
If lo >= hi Then Exit Do
Loop
If lo >= hi Then
lo = hi
list(hi) = med_value
Exit Do
End If

' Swap the lo and hi values.
list(hi) = list(lo)
Loop

' Sort the two sublists.
QuicksortVariant list(), min, lo - 1
QuicksortVariant list(), lo + 1, max
End Sub

Private Sub QuicksortLong(list() As Long, ByVal min As Integer, ByVal max As Integer)
Dim med_value As Long
Dim hi As Integer
Dim lo As Integer
Dim i As Integer

' If the list has no more than CutOff elements,
' finish it off with SelectionSort.
If max <= min Then Exit Sub

' Pick the dividing value.
i = Int((max - min + 1) * Rnd + min)
med_value = list(i)

' Swap it to the front.
list(i) = list(min)

lo = min
hi = max
Do
' Look down from hi for a value < med_value.
Do While list(hi) >= med_value
hi = hi - 1
If hi <= lo Then Exit Do
Loop
If hi <= lo Then
list(lo) = med_value
Exit Do
End If

' Swap the lo and hi values.
list(lo) = list(hi)

' Look up from lo for a value >= med_value.
lo = lo + 1
Do While list(lo) < med_value
lo = lo + 1
If lo >= hi Then Exit Do
Loop
If lo >= hi Then
lo = hi
list(hi) = med_value
Exit Do
End If

' Swap the lo and hi values.
list(hi) = list(lo)
Loop

' Sort the two sublists.
    QuicksortLong list(), min, lo - 1
    QuicksortLong list(), lo + 1, max
End Sub

' Sort an array of integers.
Public Sub SortIntArray(list() As Integer)
    QuicksortInt list, LBound(list), UBound(list)
End Sub
' Sort an array of longs.
Public Sub SortLongArray(list() As Long)
    QuicksortLong list, LBound(list), UBound(list)
End Sub
' Sort an array of singles.
Public Sub SortSingleArray(list() As Single)
    QuicksortSingle list, LBound(list), UBound(list)
End Sub
' Sort an array of doubles.
Public Sub SortDoubleArray(list() As Double)
    QuicksortDouble list, LBound(list), UBound(list)
End Sub
' Sort an array of strings.
Public Sub SortStringArray(list() As String)
    QuicksortString list, LBound(list), UBound(list)
End Sub
' Sort an array of variants.
Public Sub SortVariantArray(list() As Variant)
    QuicksortVariant list, LBound(list), UBound(list)
End Sub
