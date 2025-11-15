Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Sub GetProducts()
    Dim products As Object
    Set products = GetProductListFromApi()

    While Me.lstProducts.ListCount > 0
        Me.lstProducts.RemoveItem (0)
    Wend

    If Not products Is Nothing Then
        For i = 1 To products.Count
            Me.lstProducts.AddItem (products(i)("ID") & ";" & products(i)("ProductName") & ";" & products(i)("UnitPrice"))
        Next
    End If
End Sub

Private Sub cmdGetProducts_Click()
    GetProducts
End Sub

Private Sub lstProducts_Click()
    Dim product As Dictionary
    Set product = GetProductFromApi(Me.lstProducts.Value)

    If Not product Is Nothing Then
        Me.txtID.Value = product("ID")
        Me.txtProductName.Value = product("ProductName")
        Me.txtUnitPrice.Value = product("UnitPrice")
    End If
End Sub

Private Sub cmdSaveProduct_Click()
    If Len(Me.txtProductName.Value & vbNullString) = 0 Or Len(Me.txtUnitPrice.Value & vbNullString) = 0 Then
        MsgBox "Please select or enter a product to save"
        Exit Sub
    End If

    Dim response As Object
    If Len(Me.txtID.Value & vbNullString) = 0 Then
        Set response = CreateProductFromAPI(Me.txtProductName.Value, Me.txtUnitPrice.Value)
    Else
        Set response = UpdateProductFromAPI(Me.txtID.Value, Me.txtProductName.Value, Me.txtUnitPrice.Value)
    End If

    ClearTextBoxes
    DoEvents
    GetProducts
End Sub

Private Sub cmdNewProduct_Click()
    ClearTextBoxes
End Sub

Private Sub cmdDeleteProduct_Click()
    If Len(Me.txtID.Value & vbNullString) = 0 Then
        MsgBox "Please select a product to delete"
        Exit Sub
    End If

    Dim response As Object
    Set response = DeleteProductFromAPI(txtID.Value)

    ClearTextBoxes
    DoEvents
    GetProducts
End Sub

Private Sub ClearTextBoxes()
    With Me
        .txtID.Value = ""
        .txtProductName.Value = ""
        .txtUnitPrice.Value = ""
    End With
End Sub
