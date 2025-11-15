Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnConvert_Click()
    Dim DicResult As Dictionary

    If IsNull(Me.TxtLanguage.Value) Or IsNull(Me.CmbFromLanguageName.Value) Or IsNull(Me.CmbToLanguageName.Value) Then
        MsgBox "Please select a value"
        Exit Sub
    End If

    Me.TxtLanguageC.Value = ""

    Me.lblStatus.Caption = "Convert..."

    DoEvents

    Set DicResult = GetLanguageApi(Me.TxtLanguage.Value, Me.CmbFromLanguageName.Value, Me.CmbToLanguageName.Value)

    If Not DicResult Is Nothing Then

        Me.TxtLanguageC.Value = DicResult("data")("translations")(1)("translatedText")

    End If

    Me.lblStatus.Caption = "Done"

    Set DicResult = Nothing

End Sub
