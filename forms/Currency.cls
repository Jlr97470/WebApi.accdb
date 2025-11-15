Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnConvert_Click()
    Dim Result As Dictionary

    If IsNull(Me.TxtCurrency.Value) Or IsNull(Me.CmbFromCurrencyName.Value) Or IsNull(Me.CmbToCurrencyName.Value) Then
        MsgBox "Please select a value"
        Exit Sub
    End If

    Me.TxtCurrencyC.Value = ""

    Me.lblStatus.Caption = "Convert..."

    DoEvents


    Set Result = GetConvertCurrencyApi(Me.CmbFromCurrencyName.Value, Me.CmbToCurrencyName.Value, Me.TxtCurrency.Value)

    If Not Result Is Nothing Then

    Me.TxtCurrencyC.Value = Result("to_amount")

    End If

    Me.lblStatus.Caption = "Done"

    Set Result = Nothing

End Sub
