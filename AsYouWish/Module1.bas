Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function DlPortReadPortUchar Lib "dlportio.dll" (ByVal Port As Long) As Byte
Public Declare Sub DlPortWritePortUchar Lib "dlportio.dll" (ByVal Port As Long, ByVal Value As Byte)

 Type DeviceData
 Device1 As String * 15
 Device2 As String * 15
 Device3 As String * 15
 Device4 As String * 15
 Device5 As String * 15
 Device6 As String * 15
 Device7 As String * 15
 Device8 As String * 15

 Device9 As String * 15
 Device10 As String * 15
 Device11 As String * 15
 Device12 As String * 15
 Device13 As String * 15
 Device14 As String * 15
 Device15 As String * 15
 Device16 As String * 15

 Device17 As String * 15
 Device18 As String * 15
 Device19 As String * 15
 Device20 As String * 15
 Device21 As String * 15
 Device22 As String * 15
 Device23 As String * 15
 Device24 As String * 15
 
 End Type

Type ConfigData

 SpeakerName As String * 20
 PportAddress As String * 4
 ThresholdLevel As String * 3
 MicName As String * 15
End Type

