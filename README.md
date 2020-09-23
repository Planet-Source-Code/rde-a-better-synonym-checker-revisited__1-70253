<div align="center">

## A Better Synonym Checker Revisited


</div>

### Description

This is an extension of 'A Better Synonym Checker' submission by Zaphod that can be found at www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=32832&amp;lngWId=1 ...

His SynonymCheck method keeps M$ Word hidden, and does NOT require the M$ Word 9 Object Library ...

This 4k class contains the following functions: GrammarCheck, SpellCheck, SynonymCheck, SpellCheckResults, SynonymCheckResults ...

The last two do NOT display an M$ dialog box, but return an array of possible results ...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2008-03-13 08:40:18
**By**             |[Rde](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rde.md)
**Level**          |Beginner
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Microsoft Office Apps/VBA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/microsoft-office-apps-vba__1-42.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[A\_Better\_S2105843132008\.zip](https://github.com/Planet-Source-Code/rde-a-better-synonym-checker-revisited__1-70253/archive/master.zip)

### API Declarations

```
Private cMsWord As cMsSpell
'
Set cMsWord = New cMsSpell
'
s = "check tis words" &amp; vbCrLf &amp; _
  "some more textt"
'
Debug.Print cMsWord.SpellCheck(s)
```





