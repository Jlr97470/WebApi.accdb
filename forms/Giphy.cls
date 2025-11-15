Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnResearch_Click()
    Dim Data As Variant
    Dim DicResult As Dictionary

    If IsNull(Me.TxtGiphy.Value) Then
        MsgBox "Please select a value"
        Exit Sub
    End If

    Me.lblStatus.Caption = "Recherche..."

    DoEvents

    Set DicResult = GetGiphyApi(Me.TxtGiphy.Value)

    Me.CmbGiphyImages.RowSource = ""

    If Not DicResult Is Nothing Then

        For Each Data In DicResult("data")

            Me.CmbGiphyImages.AddItem Data("embed_url")

        Next

    End If

    Me.lblStatus.Caption = "Done"

    Set Data = Nothing

    Set DicResult = Nothing

End Sub


Private Sub CmbGiphyImages_Click()

    Me.WebBrowser.Navigate (Me.CmbGiphyImages)
End Sub
