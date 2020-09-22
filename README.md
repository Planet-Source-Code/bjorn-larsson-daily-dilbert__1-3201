<div align="center">

## Daily Dilbert


</div>

### Description

This program Downloads the archive page from

' http://www.unitedmedia.com/comics/dilbert

' extracts the file name of the picture. Download it. Convert it to bitmap.

' Changes the desktop wallpaper to the new dilbert strip.
 
### More Info
 
You must have an internet connection..

2 files on c:\ dilbert.gif and dilbert.bmp

' A changed wallpaper.

Sometimes the desktop wallpaper is not visible

' unless you refresh the desktop yourself

' or select it manually. I don't know why this happends...


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bjorn Larsson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bjorn-larsson.md)
**Level**          |Unknown
**User Rating**    |5.9 (643 globes from 109 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Jokes/ Humor](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/jokes-humor__1-40.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bjorn-larsson-daily-dilbert__1-3201/archive/master.zip)

### API Declarations

```
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
```


### Source Code

```
' 1. Create a new form.
' 2. Add a Textbox,a pictureBox and an Inet control.
' 3. Let them all have their default name.
' 4. Put all the code expect the global decleration in the "form load procedure"
 Dim Pos As Integer
 Dim Pos2 As Integer
 Dim Bilden() As Byte
 Dim NrString As String
 Text1.Text = Inet1.OpenURL ("http://www.unitedmedia.com/comics/dilbert/archive/") 'Download the page.
 Pos = InStr(1, Text1.Text, "/comics/dilbert/archive/images/dilbert")
 Pos2 = InStr(Pos, Text1.Text, ".gif")
 NrString = Mid(Text1.Text, Pos, Pos2 - Pos)
 Text1.Text = "http://www.unitedmedia.com" + NrString + ".gif" ' Debug filename
 Bilden() = Inet1.OpenURL("http://www.unitedmedia.com" + NrString + ".gif", icByteArray) ' Download picture.
 Open "C:\dilbert.gif" For Binary Access Write As #1 ' Save the file.
 Put #1, , Bilden()
 Close #1
 Picture1.Picture = LoadPicture("c:\dilbert.gif") 'Reload it to PictureBox
 SavePicture Picture1.Picture, "c:\dilbert.bmp"  'Converted to bmp..
 Call SystemParametersInfo(20, 0, "c:\dilbert.bmp", 1) 'Change the wallpaper.
 Unload Me ' Exit program
```

