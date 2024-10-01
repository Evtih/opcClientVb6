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
   Begin VB.TextBox txtTimestamp 
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtQuality 
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtValue 
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Timestamp"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Quality"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Value"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents OPCServer As OPCServer
Attribute OPCServer.VB_VarHelpID = -1
Private WithEvents OPCGroup As OPCGroup
Attribute OPCGroup.VB_VarHelpID = -1
Private OPCItems As OPCItems
Private OPCItem As OPCItem

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    ' Создаем экземпляр OPC сервера
    Set OPCServer = New OPCServer
    
    ' Подключаемся к серверу
    OPCServer.Connect "OPCServer.WinCC", "192.168.0.102"
    ' OPCServer.Connect "opcserver.wincc", "192.168.0.101"
    ' OPCServer.Connect "OPCServer.WinCC", "192.168.0.101"
    
    ' Добавляем группу
    Set OPCGroup = OPCServer.OPCGroups.Add("Group1")
    OPCGroup.UpdateRate = 1000 ' Обновление каждую секунду
    OPCGroup.IsActive = True
    
    ' Получаем коллекцию элементов группы
    Set OPCItems = OPCGroup.OPCItems
    
    ' Добавляем тег в группу
    Set OPCItem = OPCItems.AddItem("FIR1_121_PV_act", 1)
    
    ' Запускаем асинхронное чтение
    OPCGroup.IsSubscribed = True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка: " & Err.Description, vbExclamation
End Sub

Private Sub OPCGroup_DataChange(ByVal TransactionID As Long, ByVal NumItems As Long, ClientHandles() As Long, ItemValues() As Variant, Qualities() As Long, TimeStamps() As Date)
    ' Обработка изменения данных
    Dim i As Long
    For i = 1 To NumItems
        If ClientHandles(i) = 1 Then ' Проверяем, что это наш тег
            txtValue.Text = ItemValues(i)
            txtQuality.Text = Qualities(i)
            txtTimestamp.Text = TimeStamps(i)
        End If
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Отключаемся от сервера при закрытии формы
    If Not OPCServer Is Nothing Then
        OPCServer.Disconnect
        Set OPCServer = Nothing
    End If
End Sub
