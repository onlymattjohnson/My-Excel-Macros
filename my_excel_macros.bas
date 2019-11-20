Function Combine(WorkRng As Range, Optional Sign As String = ",", Optional IgnoreEmpty As Boolean = True) As String
    'This function combines text with a delimeter provided
    Dim Rng As Range
    Dim OutStr As String
    For Each Rng In WorkRng
        If IgnoreEmpty Then
            If Rng.Text <> "," And Len(Rng.Text) > 1 Then
                OutStr = OutStr & Rng.Text & Sign
            End If
        Else
            If Rng.Text <> "," Then
                OutStr = OutStr & Rng.Text & Sign
            End If
        End If
    Next
    Combine = Left(OutStr, Len(OutStr) - 1)
End Function

