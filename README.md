<div align="center">

## Text Sorting Simple and Easy


</div>

### Description

This is an easy, simple and quick code to sorte up a bunch of text lines, say you merged 2 text files and then use this code?.. it only care about the first letter of each line but its easy to add the full line if needed, look
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Spyo](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/spyo.md)
**Level**          |Beginner
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 6\.0, VB Script
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/spyo-text-sorting-simple-and-easy__1-51001/archive/master.zip)





### Source Code

<br><b>Revised, it was nice and simple but failed
<br>too many text files, thanks for all inputs</b>
<br><br>Also Add a Text1 Multi line = True
<br>and Command1 to a Form then dump all the code <br>below into the form, please comment some more
<br><br>Option Explicit
<br><br>Private Sub Command1_Click()
<br>Dim Ray() As String, Oui As Boolean, z As Byte
<br>Dim TmpRay As New Collection
<br>Dim i As Integer, x As Integer, y As Integer <br>Dim No As Integer, Pas As Integer
<br>z = 255
<br>'last asc caracter also it is max up for a byte var
<br>Oui = False
<br>' a good name for a true false var, Oui mean Yes in french
<br>TmpRay.Add "ÿ"
<br>'last possible caracter Asc255 added only for the first comparason
<br>Text1 = "FLine 1" & vbCrLf & "XLine 2" & vbCrLf & "BLine 3" & vbCrLf & "ELine 4" & vbCrLf & "HLine 5" & vbCrLf & "ALine 6" & vbCrLf & "MLine 7" & vbCrLf & "BLine 8" & vbCrLf & "GLine 9"
<br><br>Ray() = Split(Text1, vbCrLf)
<br>For Pas = 0 To UBound(Ray)
<br> 'we splitted this amount of vdCrLt so we set it as max
<br>For i = 0 To UBound(Ray)
<br> 'this is how many comparason per pass
<br>x = Asc(Left(Ray(i), 1))
<br>If x < z Then
<br>'it may be lower lets see if its a reapeat
<br>No = 0
<br>Do
<br>No = No + 1
<br>If Ray(i) = TmpRay(No) Then
<br>Oui = True
<br>'while in do loop,saw it was already there
<br>End If
<br>Loop Until No = TmpRay.Count
<br>' after No is equal to the collection we see if oui is still false
<br>If Oui = False Then
<br>z = x
<br>'z reset at 255 then keep shrinking till nothing is lower
<br>y = i
<br>'y will hold the lowest possible line
<br>End If
<br>End If
<br>Oui = False
<br>'reset the oui to False default value
<br>Next i
<br>TmpRay.Add Ray(y)
<br>'finally sorted, unique values are added to collection
<br>z = 255 ' reset time
<br>Oui = False ' reset time
<br>Next Pas
<br>TmpRay.Remove (1)
<br>'deleting the asc255 value from the start
<br>Text1 = ""
<br>'to save lines i use this same bow to load the string now it need clearing
<br>For i = 1 To TmpRay.Count
<br>' max amount in the collection
<br>Text1 = Text1 & TmpRay(i) & vbCrLf
<br>'adding them to anything we want, textbox in this case
<br>Next i
<br>End Sub

