Attribute VB_Name = "UnitTest"
Option Explicit

Function Array2Str(Args) As String
    Dim i As Integer
    For i = 0 To UBound(Args)
        Array2Str = Array2Str & IIf(i > 0, ", ", "") & IIf(VarType(Args(i)) = vbString, Chr(34), "") & Args(i) & IIf(VarType(Args(i)) = vbString, Chr(34), "")
    Next i
End Function

Sub UnitTestSprintf()
    
    Dim s As New ResourceSprintf
    Dim FormatStr As String
    Dim ShouldPrint As String
    Dim Args As Variant
    
    ' ----------------------------------------------
    Debug.Print "SIMPLE FORMAT CONVERSIONS"
    
    FormatStr = "%s"
    Args = Array("World")
    ShouldPrint = "World"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "%10s"
    Args = Array("World")
    ShouldPrint = "     World"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "%-10s"
    Args = Array("World")
    ShouldPrint = "World     "
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "%-10.3s"
    Args = Array("World")
    ShouldPrint = "Wor       "
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "%5s"
    Args = Array("VeryLongWord")
    ShouldPrint = "VeryLongWord"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "%.5s"
    Args = Array("VeryLongWord")
    ShouldPrint = "VeryL"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "%.2f"
    Args = Array(3.1415)
    ShouldPrint = "3.14"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "%.2e"
    Args = Array(3.1415)
    ShouldPrint = "3.14e+00"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "%f"
    Args = Array(3.1415)
    ShouldPrint = "3.1415"  ' 6 digits is default
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "%d"
    Args = Array(3.1415)
    ShouldPrint = "3"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "%.0f"
    Args = Array(3.1415)
    ShouldPrint = "3"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "%#.0f"
    Args = Array(3.1415)
    ShouldPrint = "3."
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "%04d"
    Args = Array(23)
    ShouldPrint = "0023"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "%-04d"
    Args = Array(23)
    ShouldPrint = "2300"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)" & " Note: Left-align + pad-w-zeros assumes want trailing zeros (otherwise use %04d)"
    FormatStr = "%-+04d"
    Args = Array(23)
    ShouldPrint = "+230"  ' TODO: Fix this
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "%i"
    Args = Array(-23)
    ShouldPrint = "-23"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "%u"
    Args = Array(-23)
    ShouldPrint = "23"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "%+d"
    Args = Array(23)
    ShouldPrint = "+23"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "% d"
    Args = Array(23)
    ShouldPrint = " 23"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "% d"
    Args = Array(-23)
    ShouldPrint = "-23"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "%o"
    Args = Array(9)
    ShouldPrint = "11"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "%x"
    Args = Array(111)
    ShouldPrint = "6f"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "%X"
    Args = Array(111)
    ShouldPrint = "6F"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "%g"
    Args = Array(0.000001)
    ShouldPrint = "1.0e-06"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "%g"
    Args = Array(0.01)
    ShouldPrint = "0.01"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "%6.2f"
    Args = Array("Hello")
    ShouldPrint = " Hello"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    FormatStr = "%6.2f"
    Args = Array("3.1415")
    ShouldPrint = "  3.14"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ") = " & Chr(34) & s.sprintf(FormatStr, Args) & Chr(34) & " (Should be """ & ShouldPrint & """)"
    
    Debug.Print ""
    
    
    FormatStr = "File%05d_%04d-%02d-%02d.%s"
    Args = Array(3, 2019, 2, 10, "dat")
    ShouldPrint = "File00003_2019-02-10.dat"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ")"
    Debug.Print Chr(34) & s.sprintf(FormatStr, Args) & Chr(34)
    Debug.Print Chr(34) & ShouldPrint & Chr(34) & "(Should be)"
    Debug.Print ""
    
    
    ' ----------------------------------------------
    Debug.Print "IDENTIFIERS"
    
    FormatStr = "Word1=%3$s, Word2=%1$s, Word3=%2$s"
    Args = Array("Arg1", "Arg2", "Arg3")
    ShouldPrint = "Word1=Arg3, Word2=Arg1, Word3=Arg2"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ")"
    Debug.Print Chr(34) & s.sprintf(FormatStr, Args) & Chr(34)
    Debug.Print Chr(34) & ShouldPrint & Chr(34) & "(Should be)"
    Debug.Print ""
    
    FormatStr = "Word1=%3$s, Word2=%s, Word3=%s, Word4=%s"
    Args = Array("Arg1", "Arg2", "Arg3")
    ShouldPrint = "Word1=Arg3, Word2=Arg1, Word3=Arg2, Word4=Arg3"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ")"
    Debug.Print Chr(34) & s.sprintf(FormatStr, Args) & Chr(34)
    Debug.Print Chr(34) & ShouldPrint & Chr(34) & "(Should be)"
    Debug.Print ""
    
    
    Debug.Print "BREAK LINES and SPECIAL CHARACTERS"
    
    FormatStr = "LINE1\nLINE2"
    Args = Array()
    ShouldPrint = "LINE1" & vbNewLine & "LINE2"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ")"
    Debug.Print Chr(34) & s.sprintf(FormatStr, Args) & Chr(34)
    Debug.Print Chr(34) & ShouldPrint & Chr(34) & "(Should be)"
    Debug.Print ""
    
    FormatStr = "The sprintf format should be %%03d followed by a \\n"
    Args = Array()
    ShouldPrint = "The sprintf format should be %03d followed by a \n"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ")"
    Debug.Print Chr(34) & s.sprintf(FormatStr, Args) & Chr(34)
    Debug.Print Chr(34) & ShouldPrint & Chr(34) & "(Should be)"
    Debug.Print ""
    
    FormatStr = "Tab test\tHere, Backspace Here\b"
    Args = Array()
    ShouldPrint = "Tab test" & vbTab & "here, Backspace here" & vbBack
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ")"
    Debug.Print Chr(34) & s.sprintf(FormatStr, Args) & Chr(34)
    Debug.Print Chr(34) & ShouldPrint & Chr(34) & "(Should be)"
    Debug.Print ""
    
    ' ----------------------------------------------
    Debug.Print "TABLE ALIGNMENT - RIGHT"
    
    Debug.Print ">>sprintf(""Index  Value1   Value2"")"
    Debug.Print "  sprintf(""%5d  %6.2f   %6.2f"", 1, 0.2, 5.7)"
    Debug.Print "  sprintf(""%5d  %6.2f   %6.2f"", 2, 10.2, -15)"
    Debug.Print Chr(34) & s.sprintf("Index  Value1   Value2") & Chr(34)
    Debug.Print Chr(34) & s.sprintf("%5d  %6.2f   %6.2f", 1, 0.2, 5.7) & Chr(34)
    Debug.Print Chr(34) & s.sprintf("%5d  %6.2f   % 6.2f", 2, 10.2, -15) & Chr(34)
    Debug.Print ""
    
    Debug.Print "TABLE ALIGNMENT - LEFT"
    
    Debug.Print ">>sprintf(""Index  Value1   Value2"")"
    Debug.Print "  sprintf(""%-5d  %-6.2f   %-6.2f"", 1, 0.2, 5.7)"
    Debug.Print "  sprintf(""%-5d  %-6.2f   %-6.2f"", 2, 10.2, -5.7)"
    Debug.Print Chr(34) & s.sprintf("Index  Value1   Value2") & Chr(34)
    Debug.Print Chr(34) & s.sprintf("%-5d  %-6.2f   %-6.2f", 1, 0.2, 5.7) & Chr(34)
    Debug.Print Chr(34) & s.sprintf("%-5d  %-6.2f   %-6.2f", 2, 10.2, -5.7) & Chr(34)
    Debug.Print ""
    
    ' ----------------------------------------------
    Debug.Print "INPUT ERRORS"
    
    FormatStr = "%td %d"
    Args = Array(5, 34)
    ShouldPrint = "%td 5"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ")"
    Debug.Print Chr(34) & s.sprintf(FormatStr, Args) & Chr(34)
    Debug.Print Chr(34) & ShouldPrint & Chr(34) & "(Should be)" & " Note: error doesn't consume input"
    Debug.Print ""
    
    FormatStr = "The bank rate is 15% and rising"
    Args = Array()
    ShouldPrint = "The bank rate is 15% and rising"
    Debug.Print ">>sprintf(""" & FormatStr & """, " & Array2Str(Args) & ")"
    Debug.Print Chr(34) & s.sprintf(FormatStr, Args) & Chr(34)
    Debug.Print Chr(34) & ShouldPrint & Chr(34) & "(Should be)" & " Note: if error string returned it would be ""The bank rate is 15<ERROR> rising"""
    Debug.Print ""
    
    
    ' ----------------------------------------------
    Debug.Print "INPUT ARGS ON DIFFERENT FORMS"
    
    FormatStr = "Hello %d the %s"
    Args = Array(2, "World")
    ShouldPrint = "Hello 2 the World"
    Debug.Print ">>Args = Array(2, ""World"")"
    Debug.Print ">>sprintf(""" & FormatStr & """, Args)"
    Debug.Print Chr(34) & s.sprintf(FormatStr, Args) & Chr(34)
    Debug.Print Chr(34) & ShouldPrint & Chr(34) & "(Should be)"
    Debug.Print ""
    
    FormatStr = "Hello %d the %s"
    Args = Array(2, "World")
    ShouldPrint = "Hello 2 the World"
    Debug.Print ">>sprintf(""" & FormatStr & """, 2, ""World"")"
    Debug.Print Chr(34) & s.sprintf(FormatStr, 2, "World") & Chr(34)
    Debug.Print Chr(34) & ShouldPrint & Chr(34) & "(Should be)"
    Debug.Print ""
    
End Sub





Sub UnitTestGetFormatSpecProperty()

    Dim s As New ResourceSprintf
    Dim FormatStr As String
    
    FormatStr = "%s"
    Debug.Print ""
    Debug.Print "Format string """ & FormatStr & Chr(34); " " & IIf(s.GetFormatSpecProperty(FormatStr, Invalid), "is NOT a valid formatSpec", "is a valid formatSpec with") & " the following properties:"
    Debug.Print "   " & ".Invalid=" & s.GetFormatSpecProperty(FormatStr, Invalid)
    Debug.Print "   " & ".ConversionChar=" & Chr(34) & s.GetFormatSpecProperty(FormatStr, ConversionChar) & Chr(34)
    Debug.Print "   " & ".Fieldwidth=" & s.GetFormatSpecProperty(FormatStr, Fieldwidth)
    Debug.Print "   " & ".Precision=" & s.GetFormatSpecProperty(FormatStr, Precision)
    Debug.Print "   " & ".FlagLeftAlign=" & s.GetFormatSpecProperty(FormatStr, FlagLeftAlign)
    Debug.Print "   " & ".FlagLeadingZeros=" & s.GetFormatSpecProperty(FormatStr, FlagLeadingZeros)
    Debug.Print "   " & ".FlagSign=" & s.GetFormatSpecProperty(FormatStr, FlagSign)
    Debug.Print "   " & ".FlagSpace=" & s.GetFormatSpecProperty(FormatStr, FlagSpace)
    Debug.Print "   " & ".FlagHash=" & s.GetFormatSpecProperty(FormatStr, FlagHash)
    Debug.Print "   " & ".Identifier=" & s.GetFormatSpecProperty(FormatStr, Identifier)
    Debug.Print "   " & ".PadChar=" & Chr(34) & s.GetFormatSpecProperty(FormatStr, PadChar) & Chr(34)
    
    
    FormatStr = "%d"
    Debug.Print ""
    Debug.Print "Format string """ & FormatStr & Chr(34); " " & IIf(s.GetFormatSpecProperty(FormatStr, Invalid), "is NOT a valid formatSpec", "is a valid formatSpec with") & " the following properties:"
    Debug.Print "   " & ".Invalid=" & s.GetFormatSpecProperty(FormatStr, Invalid)
    Debug.Print "   " & ".ConversionChar=" & Chr(34) & s.GetFormatSpecProperty(FormatStr, ConversionChar) & Chr(34)
    Debug.Print "   " & ".Fieldwidth=" & s.GetFormatSpecProperty(FormatStr, Fieldwidth)
    Debug.Print "   " & ".Precision=" & s.GetFormatSpecProperty(FormatStr, Precision)
    Debug.Print "   " & ".FlagLeftAlign=" & s.GetFormatSpecProperty(FormatStr, FlagLeftAlign)
    Debug.Print "   " & ".FlagLeadingZeros=" & s.GetFormatSpecProperty(FormatStr, FlagLeadingZeros)
    Debug.Print "   " & ".FlagSign=" & s.GetFormatSpecProperty(FormatStr, FlagSign)
    Debug.Print "   " & ".FlagSpace=" & s.GetFormatSpecProperty(FormatStr, FlagSpace)
    Debug.Print "   " & ".FlagHash=" & s.GetFormatSpecProperty(FormatStr, FlagHash)
    Debug.Print "   " & ".Identifier=" & s.GetFormatSpecProperty(FormatStr, Identifier)
    Debug.Print "   " & ".PadChar=" & Chr(34) & s.GetFormatSpecProperty(FormatStr, PadChar) & Chr(34)
    
    FormatStr = "%f"
    Debug.Print ""
    Debug.Print "Format string """ & FormatStr & Chr(34); " " & IIf(s.GetFormatSpecProperty(FormatStr, Invalid), "is NOT a valid formatSpec", "is a valid formatSpec with") & " the following properties:"
    Debug.Print "   " & ".Invalid=" & s.GetFormatSpecProperty(FormatStr, Invalid)
    Debug.Print "   " & ".ConversionChar=" & Chr(34) & s.GetFormatSpecProperty(FormatStr, ConversionChar) & Chr(34)
    Debug.Print "   " & ".Fieldwidth=" & s.GetFormatSpecProperty(FormatStr, Fieldwidth)
    Debug.Print "   " & ".Precision=" & s.GetFormatSpecProperty(FormatStr, Precision)
    Debug.Print "   " & ".FlagLeftAlign=" & s.GetFormatSpecProperty(FormatStr, FlagLeftAlign)
    Debug.Print "   " & ".FlagLeadingZeros=" & s.GetFormatSpecProperty(FormatStr, FlagLeadingZeros)
    Debug.Print "   " & ".FlagSign=" & s.GetFormatSpecProperty(FormatStr, FlagSign)
    Debug.Print "   " & ".FlagSpace=" & s.GetFormatSpecProperty(FormatStr, FlagSpace)
    Debug.Print "   " & ".FlagHash=" & s.GetFormatSpecProperty(FormatStr, FlagHash)
    Debug.Print "   " & ".Identifier=" & s.GetFormatSpecProperty(FormatStr, Identifier)
    Debug.Print "   " & ".PadChar=" & Chr(34) & s.GetFormatSpecProperty(FormatStr, PadChar) & Chr(34)

    FormatStr = "%.4f"
    Debug.Print ""
    Debug.Print "Format string """ & FormatStr & Chr(34); " " & IIf(s.GetFormatSpecProperty(FormatStr, Invalid), "is NOT a valid formatSpec", "is a valid formatSpec with") & " the following properties:"
    Debug.Print "   " & ".Invalid=" & s.GetFormatSpecProperty(FormatStr, Invalid)
    Debug.Print "   " & ".ConversionChar=" & Chr(34) & s.GetFormatSpecProperty(FormatStr, ConversionChar) & Chr(34)
    Debug.Print "   " & ".Fieldwidth=" & s.GetFormatSpecProperty(FormatStr, Fieldwidth)
    Debug.Print "   " & ".Precision=" & s.GetFormatSpecProperty(FormatStr, Precision)
    Debug.Print "   " & ".FlagLeftAlign=" & s.GetFormatSpecProperty(FormatStr, FlagLeftAlign)
    Debug.Print "   " & ".FlagLeadingZeros=" & s.GetFormatSpecProperty(FormatStr, FlagLeadingZeros)
    Debug.Print "   " & ".FlagSign=" & s.GetFormatSpecProperty(FormatStr, FlagSign)
    Debug.Print "   " & ".FlagSpace=" & s.GetFormatSpecProperty(FormatStr, FlagSpace)
    Debug.Print "   " & ".FlagHash=" & s.GetFormatSpecProperty(FormatStr, FlagHash)
    Debug.Print "   " & ".Identifier=" & s.GetFormatSpecProperty(FormatStr, Identifier)
    Debug.Print "   " & ".PadChar=" & Chr(34) & s.GetFormatSpecProperty(FormatStr, PadChar) & Chr(34)
    
    FormatStr = "%5.4f"
    Debug.Print ""
    Debug.Print "Format string """ & FormatStr & Chr(34); " " & IIf(s.GetFormatSpecProperty(FormatStr, Invalid), "is NOT a valid formatSpec", "is a valid formatSpec with") & " the following properties:"
    Debug.Print "   " & ".Invalid=" & s.GetFormatSpecProperty(FormatStr, Invalid)
    Debug.Print "   " & ".ConversionChar=" & Chr(34) & s.GetFormatSpecProperty(FormatStr, ConversionChar) & Chr(34)
    Debug.Print "   " & ".Fieldwidth=" & s.GetFormatSpecProperty(FormatStr, Fieldwidth)
    Debug.Print "   " & ".Precision=" & s.GetFormatSpecProperty(FormatStr, Precision)
    Debug.Print "   " & ".FlagLeftAlign=" & s.GetFormatSpecProperty(FormatStr, FlagLeftAlign)
    Debug.Print "   " & ".FlagLeadingZeros=" & s.GetFormatSpecProperty(FormatStr, FlagLeadingZeros)
    Debug.Print "   " & ".FlagSign=" & s.GetFormatSpecProperty(FormatStr, FlagSign)
    Debug.Print "   " & ".FlagSpace=" & s.GetFormatSpecProperty(FormatStr, FlagSpace)
    Debug.Print "   " & ".FlagHash=" & s.GetFormatSpecProperty(FormatStr, FlagHash)
    Debug.Print "   " & ".Identifier=" & s.GetFormatSpecProperty(FormatStr, Identifier)
    Debug.Print "   " & ".PadChar=" & Chr(34) & s.GetFormatSpecProperty(FormatStr, PadChar) & Chr(34)

    FormatStr = "%-+05.4f"
    Debug.Print ""
    Debug.Print "Format string """ & FormatStr & Chr(34); " " & IIf(s.GetFormatSpecProperty(FormatStr, Invalid), "is NOT a valid formatSpec", "is a valid formatSpec with") & " the following properties:"
    Debug.Print "   " & ".Invalid=" & s.GetFormatSpecProperty(FormatStr, Invalid)
    Debug.Print "   " & ".ConversionChar=" & Chr(34) & s.GetFormatSpecProperty(FormatStr, ConversionChar) & Chr(34)
    Debug.Print "   " & ".Fieldwidth=" & s.GetFormatSpecProperty(FormatStr, Fieldwidth)
    Debug.Print "   " & ".Precision=" & s.GetFormatSpecProperty(FormatStr, Precision)
    Debug.Print "   " & ".FlagLeftAlign=" & s.GetFormatSpecProperty(FormatStr, FlagLeftAlign)
    Debug.Print "   " & ".FlagLeadingZeros=" & s.GetFormatSpecProperty(FormatStr, FlagLeadingZeros)
    Debug.Print "   " & ".FlagSign=" & s.GetFormatSpecProperty(FormatStr, FlagSign)
    Debug.Print "   " & ".FlagSpace=" & s.GetFormatSpecProperty(FormatStr, FlagSpace)
    Debug.Print "   " & ".FlagHash=" & s.GetFormatSpecProperty(FormatStr, FlagHash)
    Debug.Print "   " & ".Identifier=" & s.GetFormatSpecProperty(FormatStr, Identifier)
    Debug.Print "   " & ".PadChar=" & Chr(34) & s.GetFormatSpecProperty(FormatStr, PadChar) & Chr(34)

    FormatStr = "%-0+5.4f"
    Debug.Print ""
    Debug.Print "Format string """ & FormatStr & Chr(34); " " & IIf(s.GetFormatSpecProperty(FormatStr, Invalid), "is NOT a valid formatSpec", "is a valid formatSpec with") & " the following properties:"
    Debug.Print "   " & ".Invalid=" & s.GetFormatSpecProperty(FormatStr, Invalid)
    Debug.Print "   " & ".ConversionChar=" & Chr(34) & s.GetFormatSpecProperty(FormatStr, ConversionChar) & Chr(34)
    Debug.Print "   " & ".Fieldwidth=" & s.GetFormatSpecProperty(FormatStr, Fieldwidth)
    Debug.Print "   " & ".Precision=" & s.GetFormatSpecProperty(FormatStr, Precision)
    Debug.Print "   " & ".FlagLeftAlign=" & s.GetFormatSpecProperty(FormatStr, FlagLeftAlign)
    Debug.Print "   " & ".FlagLeadingZeros=" & s.GetFormatSpecProperty(FormatStr, FlagLeadingZeros)
    Debug.Print "   " & ".FlagSign=" & s.GetFormatSpecProperty(FormatStr, FlagSign)
    Debug.Print "   " & ".FlagSpace=" & s.GetFormatSpecProperty(FormatStr, FlagSpace)
    Debug.Print "   " & ".FlagHash=" & s.GetFormatSpecProperty(FormatStr, FlagHash)
    Debug.Print "   " & ".Identifier=" & s.GetFormatSpecProperty(FormatStr, Identifier)
    Debug.Print "   " & ".PadChar=" & Chr(34) & s.GetFormatSpecProperty(FormatStr, PadChar) & Chr(34)

    FormatStr = "%-0+.4f"
    Debug.Print ""
    Debug.Print "Format string """ & FormatStr & Chr(34); " " & IIf(s.GetFormatSpecProperty(FormatStr, Invalid), "is NOT a valid formatSpec", "is a valid formatSpec with") & " the following properties:"
    Debug.Print "   " & ".Invalid=" & s.GetFormatSpecProperty(FormatStr, Invalid)
    Debug.Print "   " & ".ConversionChar=" & Chr(34) & s.GetFormatSpecProperty(FormatStr, ConversionChar) & Chr(34)
    Debug.Print "   " & ".Fieldwidth=" & s.GetFormatSpecProperty(FormatStr, Fieldwidth)
    Debug.Print "   " & ".Precision=" & s.GetFormatSpecProperty(FormatStr, Precision)
    Debug.Print "   " & ".FlagLeftAlign=" & s.GetFormatSpecProperty(FormatStr, FlagLeftAlign)
    Debug.Print "   " & ".FlagLeadingZeros=" & s.GetFormatSpecProperty(FormatStr, FlagLeadingZeros)
    Debug.Print "   " & ".FlagSign=" & s.GetFormatSpecProperty(FormatStr, FlagSign)
    Debug.Print "   " & ".FlagSpace=" & s.GetFormatSpecProperty(FormatStr, FlagSpace)
    Debug.Print "   " & ".FlagHash=" & s.GetFormatSpecProperty(FormatStr, FlagHash)
    Debug.Print "   " & ".Identifier=" & s.GetFormatSpecProperty(FormatStr, Identifier)
    Debug.Print "   " & ".PadChar=" & Chr(34) & s.GetFormatSpecProperty(FormatStr, PadChar) & Chr(34)

    FormatStr = "%34$#.4f"
    Debug.Print ""
    Debug.Print "Format string """ & FormatStr & Chr(34); " " & IIf(s.GetFormatSpecProperty(FormatStr, Invalid), "is NOT a valid formatSpec", "is a valid formatSpec with") & " the following properties:"
    Debug.Print "   " & ".Invalid=" & s.GetFormatSpecProperty(FormatStr, Invalid)
    Debug.Print "   " & ".ConversionChar=" & Chr(34) & s.GetFormatSpecProperty(FormatStr, ConversionChar) & Chr(34)
    Debug.Print "   " & ".Fieldwidth=" & s.GetFormatSpecProperty(FormatStr, Fieldwidth)
    Debug.Print "   " & ".Precision=" & s.GetFormatSpecProperty(FormatStr, Precision)
    Debug.Print "   " & ".FlagLeftAlign=" & s.GetFormatSpecProperty(FormatStr, FlagLeftAlign)
    Debug.Print "   " & ".FlagLeadingZeros=" & s.GetFormatSpecProperty(FormatStr, FlagLeadingZeros)
    Debug.Print "   " & ".FlagSign=" & s.GetFormatSpecProperty(FormatStr, FlagSign)
    Debug.Print "   " & ".FlagSpace=" & s.GetFormatSpecProperty(FormatStr, FlagSpace)
    Debug.Print "   " & ".FlagHash=" & s.GetFormatSpecProperty(FormatStr, FlagHash)
    Debug.Print "   " & ".Identifier=" & s.GetFormatSpecProperty(FormatStr, Identifier)
    Debug.Print "   " & ".PadChar=" & Chr(34) & s.GetFormatSpecProperty(FormatStr, PadChar) & Chr(34)

End Sub
