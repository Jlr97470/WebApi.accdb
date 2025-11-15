Attribute VB_Name = "Web"
Option Compare Database
Option Explicit

Public Function CreateWebJson() As webJSON

    Set CreateWebJson = New webJSON

End Function

Public Function CreateWebClient() As webClient

    Set CreateWebClient = New webClient

End Function

Public Function CreateWebRequest() As webRequest

    Set CreateWebRequest = New webRequest

End Function

Public Function CreateWebResponse() As webResponse

    Set CreateWebResponse = New webResponse

End Function

Public Function CreateIWebAuthenticator() As IWebAuthenticator

    Set CreateIWebAuthenticator = New IWebAuthenticator

End Function
