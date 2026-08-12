Attribute VB_Name = "ProductsApi"
Option Compare Database
Option Explicit

Function GetProductListFromApi() As Object

    Dim webClient As New webClient
    webClient.BaseUrl = "https://localhost:44340"

    ' URL :  BaseUrl & webRequest.Resource
    ' https://localhost:44340/Api/ProductsApi/

    Dim webRequest As New webRequest
    webRequest.Resource = "/Api/ProductsApi/"
    webRequest.Method = WebMethod.Httpget
    webRequest.Format = WebFormat.json

    Dim webResponse As webResponse
    Set webResponse = webClient.Execute(webRequest)

    If webResponse.StatusCode <> WebStatusCode.Ok Then
        Call ShowStatusError(webResponse)
        Exit Function
    End If

    Call StandardDebugOutput(webClient, webRequest, webResponse)

    Dim root As Object
    Set root = webResponse.Data()

    Set GetProductListFromApi = root

End Function

Function GetProductFromApi(ByVal productID As Integer) As Dictionary

    Dim webClient As New webClient
    webClient.BaseUrl = "https://localhost:44340/"

    Dim webRequest As New webRequest
    webRequest.Resource = "Api/ProductsApi/"
    webRequest.Method = WebMethod.Httpget
    webRequest.Format = WebFormat.json

    webRequest.AddQuerystringParam "ID", productID

    ' URL :  BaseUrl & webRequest.Resource & querystringparam
    ' https://localhost:44340/Api/ProductsApi/?search=chocolate

    Dim webResponse As webResponse
    Set webResponse = webClient.Execute(webRequest)

    If webResponse.StatusCode <> WebStatusCode.Ok Then
        Call ShowStatusError(webResponse)
        Exit Function
    End If

    Call StandardDebugOutput(webClient, webRequest, webResponse)

    Dim root As Object
    Set root = webResponse.Data()

    Set GetProductFromApi = root

End Function

Function UpdateProductFromAPI(ByVal productID As Integer, ByVal productName As String, ByVal unitPrice As Currency) As Object

    Dim webClient As New webClient
    webClient.BaseUrl = "https://localhost:44340/"

    Dim webRequest As New webRequest
    webRequest.Resource = "Api/ProductsApi/"
    webRequest.Method = WebMethod.Httpput
    webRequest.Format = WebFormat.json

    webRequest.AddQuerystringParam "ID", productID

    webRequest.AddBodyParameter "ID", productID
    webRequest.AddBodyParameter "ProductName", productName
    webRequest.AddBodyParameter "UnitPrice", unitPrice

    Dim webResponse As webResponse
    Set webResponse = webClient.Execute(webRequest)

    If webResponse.StatusCode <> WebStatusCode.Ok Then
        Call ShowStatusError(webResponse)
        Exit Function
    End If

    Call StandardDebugOutput(webClient, webRequest, webResponse)

    Dim root As Object
    Set root = webResponse.Data()

    Set UpdateProductFromAPI = webResponse

End Function

Function CreateProductFromAPI(ByVal productName As String, ByVal unitPrice As Currency) As Object

    Dim webClient As New webClient
    webClient.BaseUrl = "https://localhost:44340/"

    Dim webRequest As New webRequest
    webRequest.Resource = "Api/ProductsApi/"
    webRequest.Method = WebMethod.Httppost
    webRequest.Format = WebFormat.json

    webRequest.AddBodyParameter "ProductName", productName
    webRequest.AddBodyParameter "UnitPrice", unitPrice

    Dim webResponse As webResponse
    Set webResponse = webClient.Execute(webRequest)

    If webResponse.StatusCode <> WebStatusCode.Ok Then
        Call ShowStatusError(webResponse)
        Exit Function
    End If

    Call StandardDebugOutput(webClient, webRequest, webResponse)

    Dim root As Object
    Set root = webResponse.Data()

    Set CreateProductFromAPI = webResponse

End Function

Function DeleteProductFromAPI(ByVal productID As Integer) As Object

    Dim webClient As New webClient
    webClient.BaseUrl = "https://localhost:44340/"

    Dim webRequest As New webRequest
    webRequest.Resource = "Api/ProductsApi/"
    webRequest.Method = WebMethod.HttpDelete
    webRequest.Format = WebFormat.json
    webRequest.AddQuerystringParam "ID", productID

    Dim webResponse As webResponse
    Set webResponse = webClient.Execute(webRequest)

    If webResponse.StatusCode <> WebStatusCode.Ok Then
        Call ShowStatusError(webResponse)
        Exit Function
    End If

    Call StandardDebugOutput(webClient, webRequest, webResponse)

    Dim root As Object
    Set root = webResponse.Data()

    Set DeleteProductFromAPI = webResponse

End Function

Sub GetProductList()
    Dim products As Object
    Dim LngProd As Long
    Set products = GetProductListFromApi()

    For LngProd = 1 To products.Count
        Debug.Print products(LngProd)("ID"), products(LngProd)("ProductName"), products(LngProd)("UnitPrice")
    Next
End Sub

Sub ShowStatusError(ByRef response As webResponse)
    MsgBox "Error Code " & response.StatusCode & vbCrLf & response.StatusDescription, vbCritical
End Sub

Sub StandardDebugOutput(ByRef client As webClient, ByRef request As webRequest, ByRef response As webResponse)
    Debug.Print client.BaseUrl & request.FormattedResource
    Debug.Print request.Method
    Debug.Print response.StatusCode, response.StatusDescription
    Debug.Print response.Content
End Sub


Sub GetProduct()
    Dim product As Dictionary
    Set product = GetProductFromApi(1)
End Sub

Sub UpdateProduct()
    Dim response As webResponse
    Set response = UpdateProductFromAPI(3, "Candy Sprinkles", 1.99)
End Sub

Sub DeleteProduct()
    Dim response As webResponse
    Set response = DeleteProductFromAPI(3)
End Sub
Sub CreateProduct()
    Dim response As webResponse
    Set response = CreateProductFromAPI("A new product", 9.99)
End Sub
