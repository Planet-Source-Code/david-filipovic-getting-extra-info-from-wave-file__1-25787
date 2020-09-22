<div align="center">

## Getting extra info from wave file


</div>

### Description

Finding out if wave file is mono or stereo, 8bit or 16bit, and how many KHz's.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[David Filipovic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/david-filipovic.md)
**Level**          |Beginner
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Sound/MP3](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/sound-mp3__1-45.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/david-filipovic-getting-extra-info-from-wave-file__1-25787/archive/master.zip)





### Source Code

<font face=Verdana size=2 color=black>
 In order for this to work you would have to have these on your form:<br>
</font>
<font face=Verdana size=2 color=red>
   - Text Boxes named:<br>
       txtFile<br>
       Freq<br>
       Bit<br>
       Channel<br>
   - Button named:<br>
       Load<br><br>
 Note: I didn't want to use Common dialog or any other control for
 simplicity, you just type the location of your wave file in a textbox.<br>
 <br>
 Code:
</font>
 <hr>
<font face=Verdana size=2 color=black>
Dim Buf As String * 58<br>Dim beg As Byte<br><br>
Private Sub Load_Click()<br>
  Open txtFile.Text For Binary As #1<br>
    Get #1, 1, Buf<br>
  Close #1<br>
  beg = InStr(1, Buf, "WAVE")<br>
  If beg = 0 Then<br>
    MsgBox "Sorry not a wave file...", vbCritical, "Error..."<br>
  Else<br>
<font color=green>
    'The 23rd byte in a wave file determines if file is Mono(1 Ascii) or Stereo(2 Ascii)<br>
</font>
    If Mid(Buf, 23, 1) = Chr$(1) Then<br>
      Channel = "Mono"<br>
    Else<br>
      Channel = "Stereo"<br>
    End If<br>
<font color=green>
    'The 25th byte in a wave file determines how many KHz the file has<br>
    '44KHz(Ascii 68{44 hexadecimal})<br>
    '22KHz(Ascii 34{22 hexadecimal})<br>
    '11KHz(Ascii 17{11 hexadecimal})<br>
</font>
    Freq = Sredi(Mid(Buf, 25, 1)) / 17 * 11 & " KHz"<br>
<font color=green>
    'The 35 byte in a wave file determines if the file is 16bit(16 ascii) or 8bit(8 ascii)<br>
</font>
    Bit = Sredi(Mid(Buf, 35, 1)) & " Bit"<br>
  End If<br>
End Sub<br><br>
Private Function Sredi(ByVal accStr As String) As String<br>
  Sredi = Trim(Str(Asc(accStr)))<br>
End Function<br>
</font>
<hr>
</font>

