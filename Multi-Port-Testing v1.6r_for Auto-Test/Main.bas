Attribute VB_Name = "MMain"
Option Explicit
'Public Declare Function GetTickCount Lib "Kernel32" () As Long
Declare Function QueryPerformanceCounter Lib "kernel32" (X As Currency) As Boolean 'replace gettickcount
Declare Function QueryPerformanceFrequency Lib "kernel32" (X As Currency) As Boolean 'replace gettickcount
Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function GetIpAddrTable_API Lib "IpHlpApi" Alias "GetIpAddrTable" (pIPAddrTable As Any, pdwSize As Long, ByVal bOrder As Long) As Long


Public Type SYSTEMTIME
       wYear As Integer
       wMonth As Integer
       wDayOfWeek As Integer
      wDay As Integer
      wHour As Integer
       wMinute As Integer
       wSecond As Integer
       wMilliseconds As Integer
End Type

Public Const MAX_COM_PORT As Integer = 32

Public Port_Status(MAX_COM_PORT) As Integer
Public Const DisConnect As Integer = 0
Public Const Waiting_Connect As Integer = 1
Public Const Connecting As Integer = 2
Public Const Modem_Link_OK As Integer = 3

Public Send_Status(MAX_COM_PORT) As Integer
Public Const Stop_Send As Integer = 0
Public Const Repeat_Send As Integer = 1

Public iLoop_Time(MAX_COM_PORT) As Long                 'Unit is Sec
Public bTCP_Receive(MAX_COM_PORT) As Boolean            'determine if bytes received
Public iReceived_Count(MAX_COM_PORT) As Long
Public iSend_Buf_index(MAX_COM_PORT) As Long
Public iReceive_Buf_index(MAX_COM_PORT) As Long
Public iShow_Port As Integer
Public iRx(MAX_COM_PORT) As Long
Public iTx(MAX_COM_PORT) As Long
Public iLoss(MAX_COM_PORT) As Long
Public iError(MAX_COM_PORT) As Long
Public iFirstTick(MAX_COM_PORT) As Currency
Public iLastTick(MAX_COM_PORT) As Currency
Public Freq As Currency
Public iLatency(MAX_COM_PORT) As Currency
Public iDuration(MAX_COM_PORT) As Currency
Public Loopback_DataG(MAX_COM_PORT) As Variant         'Global array holding loopback data array (array of array)

Public Com_Port_Status As String
Public Const Com_Port_Open As String = "OPEN"
Public Const Com_Port_Close As String = "CLOSE"
Public Com_Send_Status As Integer
'Public str_buffer(MAX_COM_PORT) As Variant
Public no_send_count As Integer
Public no_receive_count(MAX_COM_PORT) As Integer
Public Custom_Send() As Byte
Public Custom_Return() As Byte

Sub Main()
Dim i As Integer
  For i = 1 To 2
    Port_Status(i) = DisConnect
    Send_Status(i) = Stop_Send
  Next i
  Com_Port_Status = Com_Port_Close
  FMain.Show
End Sub


