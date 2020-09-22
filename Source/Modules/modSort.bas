Attribute VB_Name = "modSort"
'---------------------------------------------------------------------------------------
' Procedure : quickSortRand
' Author    : Not me
' Date      : 5/14/2012
' Purpose   : Uses a recursive quicksort algorithm to sort the udttype array sent in based on
'             the value of the rndValue property.
' Note      : This is based on a comglomeration of several other submissions found on the
'             Planet site.
'---------------------------------------------------------------------------------------
Public Sub quickSortRand(iRay() As udtType, lo0, hi0)
    Dim lo As Long
    Dim hi As Long
    Dim midp As Long
    Dim mid1(1) As udtType
    Dim sortP As udtType
    lo = lo0
    hi = hi0
    If hi0 > lo0 Then
        midp = (lo0 + hi0) / 2
        mid1(0) = iRay(midp)
        Do While (lo <= hi)
            Do While ((lo < hi0) And (iRay(lo).rndValue < mid1(0).rndValue))
                lo = lo + 1
            Loop
            Do While ((hi > lo0) And (iRay(hi).rndValue > mid1(0).rndValue))
                hi = hi - 1
            Loop
            If lo <= hi Then
                sortP = iRay(lo)
                iRay(lo) = iRay(hi)
                iRay(hi) = sortP
                lo = lo + 1
                hi = hi - 1
            End If
            DoEvents
        Loop
        If lo0 < hi Then
            quickSortRand iRay, lo0, hi
        End If
        If lo < hi0 Then
            quickSortRand iRay, lo, hi0
        End If
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : quickSortLast
' Author    : Not me
' Date      : 5/14/2012
' Purpose   : Uses a recursive quicksort algorithm to sort the udttype array sent in based on
'             the value of the lastName property.
' Note      : This is based on a comglomeration of several other submissions found on the
'             Planet site.
'---------------------------------------------------------------------------------------
Public Sub quickSortLast(iRay() As udtType, lo0, hi0)
    Dim lo As Long
    Dim hi As Long
    Dim midp As Long
    Dim mid1(1) As udtType
    Dim sortP As udtType
    lo = lo0
    hi = hi0
    If hi0 > lo0 Then
        midp = (lo0 + hi0) / 2
        mid1(0) = iRay(midp)
        Do While (lo <= hi)
            Do While ((lo < hi0) And (iRay(lo).lastName < mid1(0).lastName))
                lo = lo + 1
            Loop
            Do While ((hi > lo0) And (iRay(hi).lastName > mid1(0).lastName))
                hi = hi - 1
            Loop
            If lo <= hi Then
                sortP = iRay(lo)
                iRay(lo) = iRay(hi)
                iRay(hi) = sortP
                lo = lo + 1
                hi = hi - 1
            End If
            DoEvents
        Loop
        If lo0 < hi Then
            quickSortLast iRay, lo0, hi
        End If
        If lo < hi0 Then
            quickSortLast iRay, lo, hi0
        End If
    End If
End Sub


