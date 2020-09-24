<div align="center">

## UDMX® AIntervals


</div>

### Description

<< You will need to see to understand >>This a function that I have made which gets the Years,Days,Hours,Minutes,and Seconds by two dates. For example... Let say you want to see how long it is between this ( From 2/12/2001 2:30:10 PM To 2/12/2002 3:32:05 PM ) Well it is exactly 1 year, 0 days, 1 hour, 1 minute, and 55 seconds. See that??? It pretty simple to the eye of a Mathimatician but it is hard for some ppl. This is really a big help when it comes to a chat software because let say you want to know how long a user has been online since they log on... see ^_^ very useful... This code < function > is completely free.. The preview of the code in PSC is messed up so please copy and paste it in the text form.. thx bye ***This is extremely useful. Don't think so??? then don't vote for me but if you know < or think> that this code is very useful... Please Please Vote for me...^_^ I alway try to create the stuff that hasn't been avaiable for free. This is one of the two that I have made. CHECK OUT THE UDMX® HL RTB. Trust me you need to see that one. http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=29415&lngWId=1 that is the link for code and this is the screen shot: http://www.planetsourcecode.com/upload/ScreenShots/PIC2002211322258976.jpg  ( Please take a minute and vote ^_^ )
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[uDmx IoCp©](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/udmx-iocp.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/udmx-iocp-udmx-aintervals__1-31751/archive/master.zip)





### Source Code

```
'This is the correct comments format'
Public Function AIntervals(FDate, SDate, URformat) As String
 '^_^ all my softwares and API codes are
 ' usually copyrighted © but with the funct
 ' ions
 'I give out it freely ^_^
 'But please (if you can) please put my n
 ' ame ( UDMX IOCP® ) on your software ( <=
 ' > you make 1 )
 Dim xY, xD, xH, xN, xS
 xY = DateDiff("yyyy", FDate, SDate)
 xD = DateDiff("D", FDate, SDate)
 xH = DateDiff("H", FDate, SDate)
 xN = DateDiff("N", FDate, SDate)
 xS = DateDiff("S", FDate, SDate)
 Dim VarD, VarH, VarN, VarS
 S1:
 VarD = xD - (xY * 365)
 If VarD < 0 Then		'If the date is negative then
 VarD = -VarD + 365	'It takes the 365 from the years then add it to the opposite of it self
 xY = xY - 1	'Since it took the 365 from the years, then the year must be subtract by 1
 End If
 S2:
 VarH = xH - (xD * 24)
 If VarH < 0 Then		'If the hour is negative then
 VarH = -VarH + 24	'It takes the 24 from the days then add it to the opposite of it self
 xD = xD - 1	'Since it took the 24 from the days, then the date must be subtract by 1
 GoTo S1 	'After this you will need to go back and recalculate *note that the date does not need to do this*
 End If
 S3:
 VarN = xN - (xH * 60)
 If VarN < 0 Then		'If the hour is negative then
 VarN = -VarN + 60	'It takes takes the 60 from the days then ad it to the opposite of it self
 xH = xH - 1	'Since it took the 60 from the days, then the days must be subtract by 1
 GoTo S2		'After this it will need to go back and recalculate
 End If
 VarS = xS - (xN * 60)
 If VarS < 0 Then	'If the minute is negative then
 VarS = -VarS + 60	'It takes takes the 60 from the hours then ad it to the opposite of it self
 xN = xN - 1	'Since it took the 60 from the hours, then the hours must be subtract by 1
 GoTo S3		'After this it will need to go back and recalculate
 End If
 AIntervals = Replace(Replace(Replace(Replace(Replace(LCase(URformat), "yyyy", xY), "d", VarD), "h", VarH), "n", VarN), "s", VarS)		'This is a multi replace function that I used.
'<replace continued...> All it does is replace the variable that you have inserted to this function
End Function
```

