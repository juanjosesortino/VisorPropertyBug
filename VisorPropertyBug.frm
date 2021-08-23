VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Visor PropertyBug"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   5085
   Icon            =   "VisorPropertyBug.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   5085
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtBusqueda 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1470
      TabIndex        =   5
      Top             =   30
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      DownPicture     =   "VisorPropertyBug.frx":0A02
      Height          =   420
      Left            =   1050
      Picture         =   "VisorPropertyBug.frx":1404
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Buscar"
      Top             =   30
      Width           =   390
   End
   Begin VB.CommandButton Command1 
      Height          =   420
      Left            =   540
      Picture         =   "VisorPropertyBug.frx":1E06
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Copiar Todo"
      Top             =   30
      Width           =   390
   End
   Begin VB.CommandButton cmd1 
      Height          =   420
      Left            =   30
      Picture         =   "VisorPropertyBug.frx":2808
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Texto para Inmediato"
      Top             =   30
      Width           =   390
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   6645
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   11721
      _Version        =   393217
      Indentation     =   1235
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   3300
      Top             =   3540
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   7140
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4057
            MinWidth        =   4057
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Visor PropertyBag
' DateTime  : 08/2017
' Author    : Juan José Sortino
' Purpose   : Examinar contenido de PropertyBag's
'---------------------------------------------------------------------------------------

Option Explicit
  
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" _
     (ByVal hwnd As Long, _
     ByVal hWndInsertAfter As Long, _
     ByVal x As Long, _
     ByVal y As Long, _
     ByVal cx As Long, _
     ByVal cy As Long, _
     ByVal wFlags As Long) As Long
        
Private strAction    As String
Private dFileDate    As Date
Private dFileDateNew As Date
Private nodX         As Node
Private strContenido As String

Private Sub Form_Load()
   'SetTopMostWindow Form1.hwnd, True ' TOP TOP

End Sub

Private Sub Timer_Timer()
Dim fs           As Object
Dim ts           As TextStream
Dim strContenido As String
Dim pb1          As PropertyBag

   On Error GoTo GestErr
   
   Set fs = CreateObject("Scripting.FileSystemObject")

   If Not fs.FileExists("c:\PB.dat") Then
      TreeView1.Nodes.Clear
      StatusBar.Panels(2).Text = ""
      Exit Sub
   End If
   
   DoEvents
   dFileDateNew = FileDateTime("c:\PB.dat")
   StatusBar.Panels(2).Text = "PB.dat " & dFileDateNew

   If dFileDateNew <> dFileDate Then
      StatusBar.Panels(1).Text = "Actualizando"
      dFileDate = dFileDateNew
      TreeView1.Nodes.Clear
      strContenido = ""
      
      Set pb1 = New PropertyBag
         
   '   On Error Resume Next
       
      pb1.Contents = LoadContents("c:\PB.dat")
      VerProperty pb1.Contents, 1
      Colorear_Nodos TreeView1.Nodes.Item(1)
      StatusBar.Panels(1).Text = ""
   End If
      
   Exit Sub

GestErr:
   strAction = "Timer_Timer " & Err.Description & Erl
   MsgBox strAction
End Sub
Private Function LoadContents(FilePath As String) As Variant
Dim FileNum As Integer
Dim tempContents As Variant
  
   On Error GoTo GestErr
   
   FileNum = FileSystem.FreeFile()
   Open FilePath For Binary As FileNum
   Get #FileNum, , tempContents
   Close FileNum
   LoadContents = tempContents
   
   Exit Function

GestErr:
   If Err.Number = 458 Then
      strAction = "Archivo PB.dat no válido"
   Else
      strAction = "LoadContents " & Err.Description & Erl
   End If
   MsgBox strAction
End Function

Private Sub VerProperty(ByVal pbParametros As String, ByVal iIteracion As Integer)
Dim pb1           As PropertyBag
Dim strContenido  As String
Dim strContenido2 As String
Dim ch            As String
Dim bBandera      As Boolean
Dim aCampos()     As String
Dim i             As Long
Dim x             As Long
Dim aByte()       As Byte

   Set pb1 = New PropertyBag
   aByte = pbParametros
   pb1.Contents = aByte
   strContenido = pb1.Contents

   strContenido2 = ""
   For i = 1 To Len(strContenido)
      ch = Mid(strContenido, i, 1)
      If Asc(ch) = 36 Or Asc(ch) = 4 Or Asc(ch) = 63 Or Asc(ch) = 0 Or Asc(ch) = 9 Then
         strContenido2 = strContenido2 & "#"
      End If
      If (ch >= " " And ch <= "~") Then
         If Asc(ch) <> 4 Then '$
            strContenido2 = strContenido2 & ch
         End If
      End If
   Next i
    
   bBandera = False
   Do While bBandera = False
      strContenido2 = Replace(strContenido2, "##", "#")
      strContenido2 = Replace(strContenido2, "#@#", "#")
      strContenido2 = Replace(strContenido2, "#$#", "#")
      
      If InStr(strContenido2, "##") = 0 Then bBandera = True
   Loop
   
   On Error Resume Next
   
   strContenido2 = Mid(strContenido2, 2)
   aCampos = Split(strContenido2, "#")
   For i = LBound(aCampos) To UBound(aCampos)
      If Len(aCampos(i)) > 1 Then
         strContenido = pb1.ReadProperty(aCampos(i), "")
         If Len(strContenido) = 0 Then
            aCampos(i) = Left(aCampos(i), Len(aCampos(i)) - 1)
            strContenido = pb1.ReadProperty(aCampos(i), "")
         End If
         bBandera = False
         For x = 1 To Len(strContenido)
            ch = Mid(strContenido, x, 1)
            If Asc(ch) = 36 Or Asc(ch) = 4 Or Asc(ch) = 63 Or Asc(ch) = 0 Or Asc(ch) = 9 Then
               bBandera = True
               Exit For
            End If
         Next x
   
         If Len(strContenido) > 0 Then
            If bBandera Then
               Select Case iIteracion
                  Case 1
                     pvAddPath TreeView1, "Nivel: " & iIteracion & "\" & aCampos(i)
                  Case 2
                     pvAddPath TreeView1, "Nivel: " & iIteracion - 1 & "\" & "Nivel: " & iIteracion & "\" & aCampos(i)
                  Case 3
                     pvAddPath TreeView1, "Nivel: " & iIteracion - 2 & "\" & "Nivel: " & iIteracion - 1 & "\" & "Nivel: " & iIteracion & "\" & aCampos(i)
                  Case 4
                     pvAddPath TreeView1, "Nivel: " & iIteracion - 3 & "\" & "Nivel: " & iIteracion - 2 & "\" & "Nivel: " & iIteracion - 1 & "\" & "Nivel: " & iIteracion & "\" & aCampos(i)
                  Case 5
                     pvAddPath TreeView1, "Nivel: " & iIteracion - 4 & "\" & "Nivel: " & iIteracion - 3 & "\" & "Nivel: " & iIteracion - 2 & "\" & "Nivel: " & iIteracion - 1 & "\" & "Nivel: " & iIteracion & "\" & aCampos(i)
                  Case 6
                     pvAddPath TreeView1, "Nivel: " & iIteracion - 5 & "\" & "Nivel: " & iIteracion - 4 & "\" & "Nivel: " & iIteracion - 3 & "\" & "Nivel: " & iIteracion - 2 & "\" & "Nivel: " & iIteracion - 1 & "\" & "Nivel: " & iIteracion & "\" & aCampos(i)
                  Case 7
                     pvAddPath TreeView1, "Nivel: " & iIteracion - 6 & "\" & "Nivel: " & iIteracion - 5 & "\" & "Nivel: " & iIteracion - 4 & "\" & "Nivel: " & iIteracion - 3 & "\" & "Nivel: " & iIteracion - 2 & "\" & "Nivel: " & iIteracion - 1 & "\" & "Nivel: " & iIteracion & "\" & aCampos(i)
                  Case 8
                     pvAddPath TreeView1, "Nivel: " & iIteracion - 7 & "\" & "Nivel: " & iIteracion - 6 & "\" & "Nivel: " & iIteracion - 5 & "\" & "Nivel: " & iIteracion - 4 & "\" & "Nivel: " & iIteracion - 3 & "\" & "Nivel: " & iIteracion - 2 & "\" & "Nivel: " & iIteracion - 1 & "\" & "Nivel: " & iIteracion & "\" & aCampos(i)
                  Case 9
                     pvAddPath TreeView1, "Nivel: " & iIteracion - 8 & "\" & "Nivel: " & iIteracion - 7 & "\" & "Nivel: " & iIteracion - 6 & "\" & "Nivel: " & iIteracion - 5 & "\" & "Nivel: " & iIteracion - 4 & "\" & "Nivel: " & iIteracion - 3 & "\" & "Nivel: " & iIteracion - 2 & "\" & "Nivel: " & iIteracion - 1 & "\" & "Nivel: " & iIteracion & "\" & aCampos(i)
               End Select
               VerProperty pb1.ReadProperty(aCampos(i)), iIteracion + 1 'Recursiva
            Else
               Select Case iIteracion
                  Case 1
                     pvAddPath TreeView1, "Nivel: " & iIteracion & "\" & aCampos(i) & " : """ & pb1.ReadProperty(aCampos(i)) & """"
                  Case 2
                     pvAddPath TreeView1, "Nivel: " & iIteracion - 1 & "\" & "Nivel: " & iIteracion & "\" & aCampos(i) & " : """ & pb1.ReadProperty(aCampos(i)) & """"
                  Case 3
                     pvAddPath TreeView1, "Nivel: " & iIteracion - 2 & "\" & "Nivel: " & iIteracion - 1 & "\" & "Nivel: " & iIteracion & "\" & aCampos(i) & " : """ & pb1.ReadProperty(aCampos(i)) & """"
                  Case 4
                     pvAddPath TreeView1, "Nivel: " & iIteracion - 3 & "\" & "Nivel: " & iIteracion - 2 & "\" & "Nivel: " & iIteracion - 1 & "\" & "Nivel: " & iIteracion & "\" & aCampos(i) & " : """ & pb1.ReadProperty(aCampos(i)) & """"
                  Case 5
                     pvAddPath TreeView1, "Nivel: " & iIteracion - 4 & "\" & "Nivel: " & iIteracion - 3 & "\" & "Nivel: " & iIteracion - 2 & "\" & "Nivel: " & iIteracion - 1 & "\" & "Nivel: " & iIteracion & "\" & aCampos(i) & " : """ & pb1.ReadProperty(aCampos(i)) & """"
                  Case 6
                     pvAddPath TreeView1, "Nivel: " & iIteracion - 5 & "\" & "Nivel: " & iIteracion - 4 & "\" & "Nivel: " & iIteracion - 3 & "\" & "Nivel: " & iIteracion - 2 & "\" & "Nivel: " & iIteracion - 1 & "\" & "Nivel: " & iIteracion & "\" & aCampos(i) & " : """ & pb1.ReadProperty(aCampos(i)) & """"
                  Case 7
                     pvAddPath TreeView1, "Nivel: " & iIteracion - 6 & "\" & "Nivel: " & iIteracion - 5 & "\" & "Nivel: " & iIteracion - 4 & "\" & "Nivel: " & iIteracion - 3 & "\" & "Nivel: " & iIteracion - 2 & "\" & "Nivel: " & iIteracion - 1 & "\" & "Nivel: " & iIteracion & "\" & aCampos(i) & " : """ & pb1.ReadProperty(aCampos(i)) & """"
                  Case 8
                     pvAddPath TreeView1, "Nivel: " & iIteracion - 7 & "\" & "Nivel: " & iIteracion - 6 & "\" & "Nivel: " & iIteracion - 5 & "\" & "Nivel: " & iIteracion - 4 & "\" & "Nivel: " & iIteracion - 3 & "\" & "Nivel: " & iIteracion - 2 & "\" & "Nivel: " & iIteracion - 1 & "\" & "Nivel: " & iIteracion & "\" & aCampos(i) & " : """ & pb1.ReadProperty(aCampos(i)) & """"
                  Case 9
               End Select
            End If
         End If
      End If
   Next i
End Sub

Private Sub Form_Resize()
   If Me.WindowState = vbMinimized Or Me.WindowState = vbMaximized Then Exit Sub
   If WindowState <> vbMinimized Then
      If Me.Width > 5200 Then
         Me.Width = 5200
      End If
      If Me.Height > 8000 Then
         Me.Height = 8000
      End If
   End If
End Sub
 
Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) As Long
   If Topmost = True Then 'Make the window topmost
      SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   Else
      SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
      SetTopMostWindow = False
   End If
End Function

Private Sub pvAddPath(oCtl As TreeView, ByVal sPath As String)
    Dim lNext           As Long
    Dim lStart          As Long
   
    If oCtl.Nodes.Count = 0 Then
        oCtl.Indentation = 0
    End If
    Do While lStart < Len(sPath)
        lNext = InStr(lStart + 1, sPath, "\")
        If lNext = 0 Then
            lNext = Len(sPath) + 1
        End If
        On Error Resume Next
        If lStart = 0 Then
            oCtl.Nodes.Add(, , Left$(sPath, lNext), Left$(sPath, lNext)).Expanded = True
            strContenido = strContenido & sPath & vbCrLf
        Else
            oCtl.Nodes.Add(Left$(sPath, lStart), tvwChild, Left$(sPath, lNext), Mid$(sPath, lStart + 1, lNext - lStart - 1)).Expanded = True
        End If
        
        On Error GoTo 0
        lStart = lNext
    Loop
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
   Clipboard.Clear
   Clipboard.SetText Node.Text
End Sub
Private Sub Command1_Click()
   Clipboard.Clear
   Clipboard.SetText strContenido
End Sub
Private Sub cmd1_Click()
Dim strTexto As String
   Clipboard.Clear
   ' Ejecutar este codigo desde el inmediato:
   strTexto = "FileNum = FileSystem.FreeFile()" & vbCrLf
   strTexto = strTexto & "Open " & Chr(34) & "C:\PB.DAT" & Chr(34) & " For Binary As FileNum" & vbCrLf
   strTexto = strTexto & "Put #FileNum, , pb.Contents" & vbCrLf
   strTexto = strTexto & "Close FileNum" & vbCrLf
   Clipboard.SetText strTexto
End Sub
Private Sub Command2_Click()
   If Len(txtBusqueda.Text) = 0 Then Exit Sub
   Recorrer_Nodos TreeView1.Nodes.Item(1), txtBusqueda.Text
End Sub
Private Sub Recorrer_Nodos(objNode As Node, ByVal strText As String)
      
   Dim Nodo As Node
      
   Set Nodo = objNode
   
   ' recorre los nodos en forma recursiva hasta que no haya mas nodos
   Do
      Nodo.ForeColor = vbBlack
      ' muestra el resultado en el textbox
      If InStr(Nodo.Text, strText) > 0 Then
         Nodo.ForeColor = vbRed
      End If
          
      If Not Nodo.Child Is Nothing Then
         ' Nodos que cuelgan del actual
         Call Recorrer_Nodos(Nodo.Child, strText)
      End If
          
      ' siguiente nodo
      Set Nodo = Nodo.Next
     
   Loop While Not Nodo Is Nothing
  
End Sub
  
Private Sub Colorear_Nodos(objNode As Node)
      
   Dim Nodo As Node
   Dim iColor          As Single
   
   Set Nodo = objNode
   
   If InStr(Nodo.FullPath, "Nivel: 1") Then
      iColor = vbBlack
   End If
   If InStr(Nodo.FullPath, "Nivel: 2") Then
      iColor = vbBlue
   End If
   If InStr(Nodo.FullPath, "Nivel: 3") Then
      iColor = vbMagenta
   End If
   If InStr(Nodo.FullPath, "Nivel: 4") Then
      iColor = RGB(255, 100, 0)
   End If
   If InStr(Nodo.FullPath, "Nivel: 5") Then
      iColor = RGB(100, 100, 150)
   End If
   If InStr(Nodo.FullPath, "Nivel: 6") Then
      iColor = RGB(55, 155, 55)
   End If
   
   
   ' recorre los nodos en forma recursiva hasta que no haya mas nodos
   Do
      Nodo.ForeColor = iColor
      
      If Not Nodo.Child Is Nothing Then
         ' Nodos que cuelgan del actual
         Call Colorear_Nodos(Nodo.Child)
      End If
          
      ' siguiente nodo
      Set Nodo = Nodo.Next
     
   Loop While Not Nodo Is Nothing
  
End Sub
