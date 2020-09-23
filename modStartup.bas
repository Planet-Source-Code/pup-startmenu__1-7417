Attribute VB_Name = "modStartup"
Option Explicit

Public Const StartMenuPath = "C:\windows\start menu\programs"

Sub Main()
  Load M0
  M0.Top = 0 'Me.Top - Me.Height
  M0.Left = 0 'Me.Left + Me.Width - 200
  M0.GetMenu StartMenuPath
  'SubShown(3) = True
End Sub


