VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRTB 
   Caption         =   "Rich Text Box with Background"
   ClientHeight    =   2925
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlBrowse 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pctRTB 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1635
      Left            =   60
      ScaleHeight     =   1635
      ScaleWidth      =   5355
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Width           =   5355
      Begin RichTextLib.RichTextBox rtbTrans 
         Height          =   1575
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2778
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"frmRTB.frx":0000
      End
   End
End
Attribute VB_Name = "frmRTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private sBGImg As String

Private Sub Form_Load()
  'A bit of test poetry
  rtbTrans.TextRTF = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fnil\fprq1\fcharset0 Arial;}{\f1\fnil\fprq1\fcharset0 Times New Roman;}{\f2\fnil\fcharset0 Calibri;}}" & vbNewLine & _
  "{\colortbl ;\red255\green0\blue0;\red0\green176\blue80;\red0\green77\blue187;\red255\green255\blue0;\red155\green0\blue211;\red255\green192\blue0;}" & vbNewLine & "{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\pard\sa200\sl276\slmult1\cf1\highlight0\lang9\b\f0\fs32 'Twas brillig and the slithy toves\par" & vbNewLine & "\cf2\b0\i Did gyre and gimble in the wabe\par" & vbNewLine & "\cf3\ul\i0 All Mimsy were the borogroves\par" & vbNewLine & "\cf4\ulnone\strike And the mome raths outgrabe\cf0\strike0\par" & vbNewLine & "\par" & vbNewLine & "\cf5\f1\fs44 ""Beware the Jabberwock my son!\par" & vbNewLine & "\cf6 The jaws that bite, the claws that catch!\par" & vbNewLine & "\cf2\b Beware the Jub-Jub bird, and shun\par" & vbNewLine & "\cf1\b0\i The frumious bandersnatch!""\cf0\highlight0\i0\f2\fs22\par" & vbNewLine & "}" & vbNewLine
  'Set the default background image
  sBGImg = App.Path & "\Test.jpg"
  'Call the resize and draw code
  Form_Resize
End Sub

Private Sub Form_Resize()
  'Just setting the text box and its container picturebox to the size of the window
  pctRTB.Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight
  rtbTrans.Move pctRTB.ScaleLeft, pctRTB.ScaleTop, pctRTB.ScaleWidth, pctRTB.ScaleHeight
  'Call the drawing code
  DrawBackground
End Sub

Private Sub rtbTrans_DblClick()
  'Change the background image
  cdlBrowse.Filter = "Image Files|*.bmp;*.jpg;*.gif"
  cdlBrowse.Flags = cdlOFNHideReadOnly
  cdlBrowse.ShowOpen
  If LenB(cdlBrowse.FileName) & LenB(cdlBrowse.FileTitle) Then
    'Set the image
    sBGImg = cdlBrowse.FileName
    'Call the drawing code
    DrawBackground
  Else
    'Disable the image
    sBGImg = vbNullString
    'Call the drawing code
    DrawBackground
  End If
End Sub

Private Sub DrawBackground()
Dim iBG     As IPictureDisp
  'Clear the image
  pctRTB.Cls
  If LenB(sBGImg) > 0 And LenB(Dir$(sBGImg, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0 Then
    'Make the rich text box background transparent
    SetTransparentRTB rtbTrans.hWnd, True
    'Load the picture into an IPictureDisp for easy painting
    Set iBG = LoadPicture(sBGImg)
    'Paint it centered
    pctRTB.PaintPicture iBG, (pctRTB.ScaleWidth - pctRTB.ScaleX(iBG.Width, vbHimetric)) / 2, (pctRTB.ScaleHeight - pctRTB.ScaleY(iBG.Height, vbHimetric)) / 2
  Else
    'Disable the image
    SetTransparentRTB rtbTrans.hWnd, False
  End If
End Sub

Private Sub SetTransparentRTB(ByVal hWnd As Long, ByVal Enable As Boolean)
  If Enable Then
    SetWindowLongA hWnd, (-20), &H20&
  Else
    SetWindowLongA hWnd, (-20), &H0&
  End If
End Sub
