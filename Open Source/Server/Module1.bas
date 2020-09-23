Attribute VB_Name = "Module122"
'****************************************
'* Al-Mokhtar For Networks Server       *
'*  By : Mokhtar saied saleh            *
'*      Syria - Abokamal                *
'*      WWW.ABOKAMAL.COM                *
'*  MOKHTAR_SS@HOTMAIL.COM              *
'*       0096394467547                  *
'****************************************
Public nowcompnum As Integer
Public usrcommand As String
Public usrcommand2 As String
Public opentimes(40) As String
Public cafprice(40) As Integer
Public toolbarstate As Boolean
Public statebarstate As Boolean
Public runinstartup As Boolean
Public Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
'ÏÇáÉÏæãÇð Ýí ÇáãÞÏãÉ
Public Sub StayOnTop(frm As Form)
  SetWindowPos frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Function CState(StNumber As Integer, CompNum)
Rem äÇã Êáæíä ÇáÍæÇÓíÈ ÍÓÈ ÇáÍÇáÉ
Rem áÇ ÊäÓì ÇáÑÞã 1 áÍÇáÉ  (Ýí ÇáÅäÊÙÇÑ)
Rem ÇáÑÞã 2 áÍÇáÉ Çááæä ÇáÃÍãÑ (ÇäÞØÚ ÇáÇÊÕÇá)
Rem ÇáÑÞã3 áÍÇáÉ  (ãÊÕá)
Rem &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&77
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&7

Pst1 = App.Path & ("\ps2.smg")
Pst2 = App.Path & ("\ps3.smg")
Pst3 = App.Path & ("\ps1.smg")
Pst4 = App.Path & ("\ps4.smg")
Select Case StNumber
Case 1
frm_main.comp(CompNum).Visible = False
frm_main.Clabel(CompNum).Visible = False
frm_main.labelopentimes.Visible = False
frm_main.SkinLabel1.Visible = False
frm_main.SkinLabel2.Visible = False
frm_main.SkinLabel3.Visible = False

Case 2
frm_main.comp(CompNum).Picture = LoadPicture(Pst2)
frm_main.comp(CompNum).Visible = True
frm_main.Clabel(CompNum).Visible = True
frm_main.Clabel(CompNum).Caption = CompNum
Case 3
If frm_main.isClosed(CompNum).Text = "true" Then
frm_main.comp(CompNum).Picture = LoadPicture(Pst3)
frm_main.comp(CompNum).Visible = True
frm_main.Clabel(CompNum).Visible = True
Else
frm_main.comp(CompNum).Picture = LoadPicture(Pst1)
frm_main.comp(CompNum).Visible = True
frm_main.Clabel(CompNum).Visible = True
End If
Case 4
frm_main.comp(CompNum).Picture = LoadPicture(Pst4)
frm_main.comp(CompNum).Visible = True
frm_main.Clabel(CompNum).Visible = True

End Select

End Function
