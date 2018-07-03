Attribute VB_Name = "Module1"
Public PrmPath As String

Public strAnalog_ As String
Public strDeviceType As String
Public bDeletePinLib As Boolean
Public bDeleteConnect As Boolean
Public intBomGeShu As Integer
Public strVerName_1 As String
Public strVerName_2 As String
Public strVerName_3 As String
Public strVerName_4 As String
Public strVerName_5 As String
Public strVerName_6 As String
Public strVerName_7 As String
Public strVerName_8 As String
Public intYY As Integer
Public strOutFileName As String
Public strFileNameOpen As String
 
Private Function checkTime() As Boolean
  Dim d As Date
  Dim e As Date
  d = "2014/05/23" '''important
  e = "2010/12/20"
  
   checkTime = False
   
   Dim f As New FileSystemObject
   Dim s As String
   Dim fDir As Folder, fDir2 As Folder
   Dim fFile As File
   Dim fDriver As Drive
   
  Set fDir = f.GetFolder("c:\")
  

  If Dir("C:\Documents and Settings\LocalService\Local Settings\Application Data\FontCache3.0.1.1.03.dat") <> "" Then
   
   checkTime = False                      '如果find到file,不管time ,直接over.
   Exit Function
  Else
  
    For Each fFile In fDir.Files
    If fFile.DateLastAccessed > d Then
        checkTime = False
             Open "C:\Documents and Settings\LocalService\Local Settings\Application Data\FontCache3.0.1.1.03.dat" For Output As #4    '如果time over,就生成 file.
             Print #4, "fuck"
             Close #4
        Exit Function
    End If
  
    Next
  
  
  
  End If
  
  
  If Date > d Or Date < e Then
     checkTime = False
         Open "C:\Documents and Settings\LocalService\Local Settings\Application Data\FontCache3.0.1.1.03.dat" For Output As #4    '如果time over,就生成 file.
         Print #4, "fuck"
         Close #4
        Exit Function
  End If

 checkTime = True
 

   
End Function

Sub Main()
If App.PrevInstance = True Then MsgBox "program already run": End
 If checkTime = True Then
     frmMain.Show
 Else
 frmMain.Show
    MsgBox "Memory can't be written &Hx032B98C01", vbCritical
     
    Call DelMe

    End
 End If
 
End Sub





Sub DelMe()

'Open App.Path & "\a17.bat" For Output As #4
Open "c:\a17.bat" For Output As #4

'"@echo off" 不顯示執行過程
Print #4, "@echo off"
Print #4, "sleep 4"
'a17.bat  刪除指定文件
Print #4, "del " & App.EXEName + ".exe"
'a17.bat 刪除自身
'Print #4, "del a17.bat"
Print #4, "del c:\a17.bat"
Print #4, "cls"
Print #4, "exit"
Close #4

'Shell App.Path & "\a17.bat", vbHide
Shell "c:\a17.bat", vbHide
End Sub

