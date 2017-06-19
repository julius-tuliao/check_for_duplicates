Attribute VB_Name = "checkdups"
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
'Modified by jindon http://www.ozgrid.com/forum/showthread.php?t=157244
        Dim rng As Range, r As Range, msg As String, x As Range
    Set rng = Intersect(Columns(29), Target)
    If Not rng Is Nothing Then
        Application.EnableEvents = False
        For Each r In rng
            If Not IsEmpty(r.Value) Then
                If Application.CountIf(Columns(29), r.Value) > 1 Then
                    msg = msg & vbLf & r.Address(0, 0) & vbTab & r.Value
                    If x Is Nothing Then
                        r.Activate
                        Set x = r
                    Else
                        Set x = Union(x, r)
                    End If
                End If
            End If
        Next
        If Len(msg) Then
            MsgBox "Duplicate Entry" & msg
            x.ClearContents
            x.Select
        End If
        Set rng = Nothing
        Set x = Nothing
        Application.EnableEvents = True
    End If
End Sub
