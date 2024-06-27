VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstServerMACs 
      Height          =   1815
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtResult 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Проверить MAC адрес"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetMACAddress Lib "iphlpapi.dll" Alias "GetAdaptersInfo" (ByVal pAdapterInfo As Long, ByRef pOutBufLen As Long) As Long

Private Type IP_ADAPTER_INFO
    Next As Long
    ComboIndex As Long
    AdapterName As String * 260
    Description As String * 132
    AddressLength As Long
    Address(0 To 7) As Byte
    Index As Long
    Type As Long
    DhcpEnabled As Long
    CurrentIpAddress As Long
    IpAddressList As Long
    GatewayList As Long
    DhcpServer As Long
    HaveWins As Long
    PrimaryWinsServer As Long
    SecondaryWinsServer As Long
    LeaseObtained As Long
    LeaseExpires As Long
End Type

Dim WithEvents Server As OPCServer
Attribute Server.VB_VarHelpID = -1
Dim Group As OPCGroup
Dim Items() As OPCItem

Private Sub cmdCheck_Click()
    Dim localMAC As String
    localMAC = GetLocalMACAddress()
    
    txtResult.Text = "Локальный MAC адрес: " & localMAC & vbCrLf
    
    ConnectToOPCServer
    
    Dim serverMACs As Variant
    serverMACs = GetServerMACAddresses()
    
    lstServerMACs.Clear
    If IsNull(serverMACs) Then
        txtResult.Text = txtResult.Text & "Не удалось получить MAC адреса с сервера"
    Else
        Dim i As Integer
        Dim found As Boolean
        found = False
        
        txtResult.Text = txtResult.Text & "MAC адреса на сервере:" & vbCrLf
        For i = LBound(serverMACs) To UBound(serverMACs)
            lstServerMACs.AddItem serverMACs(i)
            txtResult.Text = txtResult.Text & serverMACs(i) & vbCrLf
            If serverMACs(i) = localMAC Then
                found = True
            End If
        Next i
        
        If found Then
            txtResult.Text = txtResult.Text & vbCrLf & "MAC адрес этого компьютера прописан на сервере"
        Else
            txtResult.Text = txtResult.Text & vbCrLf & "MAC адрес этого компьютера НЕ прописан на сервере"
        End If
    End If
    
    DisconnectFromOPCServer
End Sub

Private Function GetLocalMACAddress() As String
    Dim AdapterInfo As IP_ADAPTER_INFO
    Dim lngBufLen As Long
    Dim lngStatus As Long
    
    lngBufLen = Len(AdapterInfo)
    lngStatus = GetMACAddress(VarPtr(AdapterInfo), lngBufLen)
    
    If lngStatus = 0 Then
        GetLocalMACAddress = ByteArrayToString(AdapterInfo.Address)
    Else
        GetLocalMACAddress = ""
    End If
End Function

Private Function ByteArrayToString(arr() As Byte) As String
    Dim i As Integer
    For i = 0 To 5
        ByteArrayToString = ByteArrayToString & Right$("0" & Hex$(arr(i)), 2)
        If i < 5 Then ByteArrayToString = ByteArrayToString & "-"
    Next i
End Function

Private Sub ConnectToOPCServer()
    Set Server = New OPCServer
    Server.Connect "OPCServer.WinCC", "192.168.0.102"
    
    On Error Resume Next
    Set Group = Server.OPCGroups.GetOPCGroup("Group1")
    If Err.Number <> 0 Then
        Set Group = Server.OPCGroups.Add("Group1")
    End If
    On Error GoTo 0
    
    Group.UpdateRate = 1000
    Group.IsActive = True
    
    ReDim Items(0)
    Set Items(0) = Group.OPCItems.AddItem("MAC_Addresses", 1)
End Sub

Private Function GetServerMACAddresses() As Variant
    Dim value As Variant
    Items(0).Read 1, value
    
    If VarType(value) = vbString Then
        GetServerMACAddresses = Split(value, ",")
    Else
        GetServerMACAddresses = Null
    End If
End Function

Private Sub DisconnectFromOPCServer()
    Server.OPCGroups.RemoveAll
    Server.Disconnect
    Set Server = Nothing
End Sub
