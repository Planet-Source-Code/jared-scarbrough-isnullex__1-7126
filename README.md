<div align="center">

## IsNullEx


</div>

### Description

This code provides an IsNull function like that used in SQL Server. This allows you to check if a variable contains null, and if so return a specific value. This can prevent you from having to write things like  iif(not isnull(rst!FieldName), rst!FieldName, ""). For new programmers, especially those working with databases, you will quickly become aware that trying to access a database field that contains a null value often results in the dreaded error #94 - Invalid Use of Null, which means you constantly have to write code to trap for that possibility.
 
### More Info
 
ValueToCheck = a variant representing he value to check for null

varWhatToReturnIfNull = the value to return if ValueToCheck is null

If ValueToCheck is null, then varWhatToReturnIfNull is returned, otherwise, ValueToCheck is returned


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jared Scarbrough](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jared-scarbrough.md)
**Level**          |Intermediate
**User Rating**    |4.3 (47 globes from 11 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jared-scarbrough-isnullex__1-7126/archive/master.zip)





### Source Code

```
Public Function IsNullEx(ValueToCheck As Variant, varWhatToReturnIfNull) As Variant
  If IsNull(ValueToCheck) Then
    IsNullEx = varWhatToReturnIfNull
  Else
    IsNullEx = ValueToCheck
  End If
End Function
Usage example:
txtClientName = IsNullEx(rst!ClientName, "unknown")
```

