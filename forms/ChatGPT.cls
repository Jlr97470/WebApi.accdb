Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnSend_Click()
    Dim Vox As Object ' SpeechLib.SpVoice
    Dim DicResult As Dictionary

    Set Vox = CreateObject("SAPI.SpVoice") ' New SpeechLib.SpVoice

    Set Vox.Voice = Vox.GetVoices.Item(0)

    If IsNull(Me.TxtQuestion.Value) Then
        MsgBox "Please select a value"
        Exit Sub
    End If

    Me.TxtReponse.Value = ""

    Me.lblStatus.Caption = "Envoyé..."

    DoEvents

    Set DicResult = GetChatGPTApi(Me.TxtQuestion.Value)

    If Not DicResult Is Nothing Then

        Me.TxtReponse.Value = DicResult("choices")(1)("message")("content")

        Me.TxtQuestion.Value = ""

        DoEvents

        Vox.Speak Me.TxtReponse.Value

    End If

    Me.lblStatus.Caption = "Done"

    Set Vox = Nothing

    Set DicResult = Nothing
End Sub
