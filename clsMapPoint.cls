VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMapPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Built from Scratch
' Created by Shad4Ever@hotmail.com
' Based on MSDN Samples

Private Enum MapGeo
  MapAmerica = 1
  MapEurope = 2
End Enum

Private MapObj As Object


' This function is used to map multiple ZipCodes
' I made it simple - no multi zip search results
' Since i needed some way to map all surrounding customers from my database with minimal user input (none)
' So i rely on the zip code (since mostly it is enough)
' And i put the adress on the balloon
' You can pass array or not into it
' And that's it !
' ZipCode= The zip code to be found
' ZipName= The name to be displayed
' ZipDetail= The detail to be displayed on the ballon
Public Sub MapZipCodes(ZipCode As Variant, ZipName As Variant, ZipDetail As Variant, Optional MapFile As String = "")
  MapObj.Visible = True
  
  If MapObj.ActiveMap Is Nothing Then
    If MapFile = "" Then
      Screen.MousePointer = 11
      MapObj.NewMap MapAmerica
      Screen.MousePointer = 0
    Else
      Screen.MousePointer = 11
      MapObj.OpenMap MapFile, True
      Screen.MousePointer = 0
    End If
  End If
  
  Dim i As Integer
  Dim oLoc As Object, oPush As Object
  
  If IsArray(ZipCode) Then
    For i = LBound(ZipCode) To UBound(ZipCode)
      If TypeName(ZipCode(i)) = "String" And TypeName(ZipName(i)) = "String" And TypeName(ZipDetail(i)) = "String" Then
        With MapObj.ActiveMap
          Set oLoc = .FindResults(ZipCode(i))(1)
          oLoc.GoTo
          
          Set oPush = .AddPushpin(oLoc, ZipName(i))
          oPush.BalloonState = geoDisplayBalloon
          oPush.Note = ZipDetail(i)
          oPush.Highlight = True
          
        End With
      End If
    Next
  Else
    If TypeName(ZipCode) = "String" And TypeName(ZipName) = "String" And TypeName(ZipDetail) = "String" Then
      With MapObj.ActiveMap
        Set oLoc = .FindResults(ZipCode)(1)
        oLoc.GoTo
        
        Set oPush = .AddPushpin(oLoc, ZipName)
        oPush.BalloonState = geoDisplayBalloon
        oPush.Note = ZipDetail
        oPush.Highlight = True
      End With
    End If
  End If
  
  ' Avoid message that say the map wasn't saved
  MapObj.ActiveMap.Saved = True
  
  Set oLoc = Nothing
  Set oPush = Nothing
End Sub


Private Sub Class_Initialize()
  Set MapObj = CreateObject("MapPoint.Application")
  
  If MapObj Is Nothing Then
    Err.Raise 101, "clsMapPoint.cls", "MapPoint COM Object not registered"
  End If
End Sub

Private Sub Class_Terminate()
  Set MapObj = Nothing
End Sub
