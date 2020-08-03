VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mengambil Data Nama File"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function StripPath(T$) As String
Dim x%, ct%
  StripPath$ = T$
  x% = InStr(T$, "\")
  Do While x%
     ct% = x%
     x% = InStr(ct% + 1, T$, "\")
  Loop
  If ct% > 0 Then StripPath$ = Mid$(T$, ct% + 1)
End Function

Private Sub Form_Load()
  'Ganti dengan nama lengkap file (beserta path-nya)
  'yang ingin Anda ambil nama file-nya.
  MsgBox StripPath(App.Path + "\mydir\myfile.exe") 'Contoh ini
  'menghasilkan: 'myfile.exe'
End Sub

