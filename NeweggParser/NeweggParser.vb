Imports System.Net
Imports System.Runtime.InteropServices

<Guid("E68ADC37-B09D-451A-96B7-8928700BD0D8")>
<InterfaceType(ComInterfaceType.InterfaceIsIDispatch)>
Public Interface _NeweggParser
    <DispId(1)> Sub LoadNeweggNumber(ByVal NeweggNumber As String)
    <DispId(2)> Sub Load(ByVal FilenameOrURL As String)
    <DispId(3)> Function GetProperty(ByVal propertyName As String) As Object
    <DispId(4)> Function GetProperty(ByVal propertyName As String, ByVal jsonObject As Object) As Object
    <DispId(5)> Function GetKeys() As String()
    <DispId(6)> Function GetKeys(ByVal jsonObject As Object) As String()
    <DispId(7)> Function GetTypeOf(ByVal jsonObject As Object) As Type
End Interface

<Guid("997121B3-A76E-47A4-B27B-4229D46A3C30")>
<ClassInterface(ClassInterfaceType.None)>
<ProgId("NeweggParser.NeweggParser")>
Public Class NeweggParser
    Implements _NeweggParser

    Private WithEvents _scripter As MSScriptControl.ScriptControl
    Private _contents As String = ""
    Private _jsonObject As Object = Nothing
    Private _source As String = ""
    Private _types As Dictionary(Of String, Type) = Nothing

    Public Sub New()
        _scripter = New MSScriptControl.ScriptControl
        _scripter.Language = "JScript"
        Dim w As New WebClient
        Dim json2 As String = w.DownloadString("http://ajax.cdnjs.com/ajax/libs/json2/20110223/json2.js")
        _scripter.AddCode(json2)
        w = Nothing
        Dim getProperty As String = _
            "function getProperty(jsonObj, propertyName) {" & _
            "   return jsonObj[propertyName]; " & _
            "}"

        Dim getKeys As String = _
            "function getKeys(jsonObj) {" & _
            "   var keys = new Array();" & _
            "   for (var i in jsonObj) {" & _
            "       keys.push(i);" & _
            "   }" & _
            "   return keys;" & _
            "}"

        Dim getTypeOf As String = _
            "function getTypeOf(jsonObj) {" & _
            "   return ({}).toString.call(jsonObj).match(/\s([a-zA-Z]+)/)[1].toLowerCase();" & _
            "}"

        _scripter.AddCode(getProperty)
        _scripter.AddCode(getKeys)
        _scripter.AddCode(getTypeOf)

        _types = New Dictionary(Of String, Type)
        _types.Add("undefined", GetType(Object))
        _types.Add("number", GetType(Integer))
        _types.Add("boolean", GetType(Boolean))
        _types.Add("string", GetType(String))
        _types.Add("array", GetType(List(Of Object)))
        _types.Add("date", GetType(Date))
    End Sub

    Private Sub ScriptError() Handles _scripter.Error
        Throw New ApplicationException(_scripter.Error.Description & " in " & _scripter.Error.Source & " - line " & _scripter.Error.Line & ", column " & _scripter.Error.Column)
    End Sub

    Public Sub LoadNeweggNumber(ByVal NeweggNumber As String) Implements _NeweggParser.LoadNeweggNumber
        Load("http://www.ows.newegg.com/Products.egg/" & NeweggNumber & "/")
    End Sub

    Public Sub Load(ByVal FilenameOrURL As String) Implements _NeweggParser.Load
        _source = FilenameOrURL
        Dim m As New System.Text.RegularExpressions.Regex("^http(s?)\://.*")
        If m.IsMatch(FilenameOrURL) Then
            'url
            Dim w As New WebClient
            _contents = w.DownloadString(FilenameOrURL)
            w = Nothing
        Else
            'filename
            _contents = New System.IO.StreamReader(FilenameOrURL).ReadToEnd
        End If
        m = Nothing
        _jsonObject = _scripter.Eval("(" & _contents & ")")
    End Sub

    Public Function GetProperty(ByVal propertyName As String) As Object Implements _NeweggParser.GetProperty
        Return _scripter.Run("getProperty", _jsonObject, propertyName)
    End Function

    Public Function GetProperty(ByVal propertyName As String, ByVal jsonObject As Object) As Object Implements _NeweggParser.GetProperty
        Return _scripter.Run("getProperty", jsonObject, propertyName)
    End Function

    Public Function GetKeys() As String() Implements _NeweggParser.GetKeys
        Return GetKeys(_jsonObject)
    End Function

    Public Function GetKeys(ByVal jsonObject As Object) As String() Implements _NeweggParser.GetKeys
        Dim keysObject As Object = _scripter.Run("getKeys", jsonObject)
        Dim length As Integer = GetProperty("length", keysObject)
        Dim keys As New List(Of String)
        For Each key In keysObject
            keys.Add(key)
        Next
        Return keys.ToArray
    End Function

    Public Function GetTypeOf(ByVal jsonObject As Object) As Type Implements _NeweggParser.GetTypeOf
        Return _types(_scripter.Run("getTypeOf", jsonObject))
    End Function
End Class
