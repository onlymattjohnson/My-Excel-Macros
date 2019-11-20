Function Combine(WorkRng As Range, Optional Sign As String = ",", Optional IgnoreEmpty As Boolean = True) As String
    'This function combines text with a delimeter provided
    Dim Rng As Range
    Dim OutStr As String
    OutStr = ""
    For Each Rng In WorkRng
        If IgnoreEmpty Then
            If Rng.Text <> "," And Not IsEmpty(Rng) Then
                OutStr = OutStr & Rng.Text & Sign
            End If
        Else
            If Rng.Text <> "," Then
                OutStr = OutStr & Rng.Text & Sign
            End If
        End If
    Next
    If Len(OutStr) = 0 Then
        Combine = OutStr
    Else
        Combine = Left(OutStr, Len(OutStr) - 1)
    End If
End Function