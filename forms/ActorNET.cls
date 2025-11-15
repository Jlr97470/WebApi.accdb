Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Function ListeActor() As Boolean
    Dim DicResult As Collection
    Dim DicName As Variant
    Dim DicData As Variant
    Dim DicDataLigne As Variant
    Dim Ligne As Variant
    Dim ColName As Variant
    Dim IntNbCol As Integer
    Dim StrItem As String
    Dim StrName(1) As String
    Dim StrData(1) As String

    Me.CmbActor.RowSource = ""

    Me.lblStatus = "Liste..."

    DoEvents

    Set DicResult = GetActorApiNET("LIST", StrName, StrData)

    If Not DicResult Is Nothing Then

        If DicResult(1)("status") = "OK" Then

            Set DicName = DicResult(1)

            For Each ColName In DicName

                Select Case ColName
                    Case "status", "status_description"

                    Case Else

                        Me.CmbActor.RowSource = Me.CmbActor.RowSource & ColName & ";"

                        IntNbCol = IntNbCol + 1

                End Select

            Next

            Me.CmbActor.ColumnCount = IntNbCol

            For Each DicDataLigne In DicResult

                StrItem = ""

                For Each ColName In DicDataLigne

                    Select Case ColName
                        Case "status", "status_description"

                        Case Else

                            StrItem = StrItem & Chr(34) & Replace(Nz(DicDataLigne(ColName), ""), ";", Chr(34) & ";" & Chr(34)) & Chr(34) & ";"

                    End Select

                Next

                Me.CmbActor.RowSource = Me.CmbActor.RowSource & StrItem

            Next

            Me.lblStatus = "Done"

        Else

            Me.lblStatus = DicResult(1)("status_description")

        End If

    Else

        Me.lblStatus = "Probléme connection web service !!!!"

    End If

    Set DicResult = Nothing

End Function

Private Sub CmdAdd_Click()
    Dim DicResult As Collection
    Dim DicName As Variant
    Dim DicData As Variant
    Dim DicDataLigne As Variant
    Dim Ligne As Variant
    Dim ColName As Variant
    Dim IntNbCol As Integer
    Dim StrItem As String
    Dim StrName() As String
    Dim StrData() As String

    If Not IsNull(Me.TxtFirst_NameAdd) And Not IsNull(Me.TxtLast_nameAdd) And Not IsNull(Me.TxtLast_UpdateAdd) Then

        Me.lblStatusAdd = "Ajout..."

        DoEvents

        ReDim Preserve StrName(2)

        ReDim Preserve StrData(2)

        StrName(0) = "first_name"

        StrData(0) = Me.TxtFirst_NameAdd

        StrName(1) = "last_name"

        StrData(1) = Me.TxtLast_nameAdd

        StrName(2) = "last_update"

        StrData(2) = Me.TxtLast_UpdateAdd

        Set DicResult = GetActorApiNET("ADD", StrName, StrData)

        If Not DicResult Is Nothing Then

            If DicResult(1)("status") = "OK" Then

                Me.lblStatusAdd = "Done"

                ListeActor

            Else

                Me.lblStatusAdd = DicResult(1)("status_description")

            End If

        Else

            Me.lblStatusAdd = "Probléme connection web service !!!!"

        End If

    End If

End Sub

Private Sub CmdDelete_Click()
    Dim DicResult As Collection
    Dim DicName As Variant
    Dim DicData As Variant
    Dim DicDataLigne As Variant
    Dim Ligne As Variant
    Dim ColName As Variant
    Dim IntNbCol As Integer
    Dim StrItem As String
    Dim StrName() As String
    Dim StrData() As String

    If Not IsNull(Me.CmbActor.Value) Then

        Me.lblStatusDelete = "Suppression..."

        DoEvents

        ReDim Preserve StrName(0)

        ReDim Preserve StrData(0)

        StrName(0) = "actor_id"

        StrData(0) = Me.CmbActor.Value

        Set DicResult = GetActorApiNET("DELETE", StrName, StrData)

        If Not DicResult Is Nothing Then

            If DicResult(1)("status") = "OK" Then

                Me.lblStatusDelete = "Done"

                ListeActor

            Else

                Me.lblStatusDelete = DicResult(1)("status_description")

            End If

        Else

            Me.lblStatusDelete = "Probléme connection web service !!!!"

        End If

    End If

End Sub

Private Sub CmdListe_Click()

    ListeActor

End Sub

Private Sub CmbActor_Click()

    If Me.CmbActor.Column(0) <> "" Then

        Me.TxtActor_id = Me.CmbActor.Column(0)

        Me.TxtFirst_Name = Me.CmbActor.Column(1)

        Me.TxtLast_name = Me.CmbActor.Column(2)

        'Me.TxtLast_Update = Format(Me.CmbActor.Column(3), "dd/mm/yyyy hh:mm:ss")

    End If
End Sub

Private Sub CmdUpdate_Click()
    Dim DicResult As Collection
    Dim DicName As Variant
    Dim DicData As Variant
    Dim DicDataLigne As Variant
    Dim Ligne As Variant
    Dim ColName As Variant
    Dim IntNbCol As Integer
    Dim StrItem As String
    Dim StrName() As String
    Dim StrData() As String

    If Not IsNull(Me.TxtFirst_Name) And Not IsNull(Me.TxtLast_name) And Not IsNull(Me.TxtLast_Update) Then

        Me.lblStatusUpdate = "Mise a jour..."

        DoEvents

        ReDim Preserve StrName(3)

        ReDim Preserve StrData(3)

        StrName(0) = "actor_id"

        StrData(0) = Me.TxtActor_id

        StrName(1) = "first_name"

        StrData(1) = Me.TxtFirst_Name

        StrName(2) = "last_name"

        StrData(2) = Me.TxtLast_name

        StrName(3) = "last_update"

        StrData(3) = Me.TxtLast_Update

        Set DicResult = GetActorApiNET("UPDATE", StrName, StrData)

        If Not DicResult Is Nothing Then

            If DicResult(1)("status") = "OK" Then

                Me.lblStatusUpdate = "Done"

                ListeActor

            Else

                Me.lblStatusUpdate = DicResult(1)("status_description")

            End If

        Else

            Me.lblStatusUpdate = "Probléme connection web service !!!!"

        End If

    End If

End Sub


Private Sub TxtLast_Update_BeforeUpdate(Cancel As Integer)

    Cancel = Not IsDate(TxtLast_Update)


End Sub

Private Sub TxtLast_UpdateAdd_BeforeUpdate(Cancel As Integer)

    Cancel = Not IsDate(TxtLast_UpdateAdd)

End Sub
