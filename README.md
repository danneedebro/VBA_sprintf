# VBA_sprintf
The C/C++/Matlab-function **sprintf** implemented in VBA

# Motivation
Ever wanted to format a number in excel VBA? It's a hazzle. Say that you would like to print out two floats with a fieldwidth of 6 and a precision of 2 as text. This is how you could do that in VBA:

``Debug.Print format(format(1.23,"0.00"),"@@@@@@") & " " & format(format(9.5,"0.00"),"@@@@@@")``

Using sprintf the command would be

``Debug.Print sprintf("%6.2f %6.2f", 1.23, 9.5)``

The second option is clearly much easier and it amazes me that every language doesn't implement the sprintf-function as a standard function

There probably a tonne of these implementations out there but I needed something that could probe the format spec (the `%03d` code) and return the fieldwidth etc (`GetFormatSpecProperty`)

# Usage
Import this .cls file as a new Class module and instanciate it from your module using
```
Dim s as New ResourceSprintf
Debug.Print s.sprintf("%6.2f %6.2f", 1.23, 9.5)
```

# Unit test
There are some nuances in how the sprintf work on different langugages. Here is a unit test how my implementation works on a set of strings

```
SIMPLE FORMAT CONVERSIONS
>>sprintf("%s", "World") = "World" (Should be "World")
>>sprintf("%10s", "World") = "     World" (Should be "     World")
>>sprintf("%-10s", "World") = "World     " (Should be "World     ")
>>sprintf("%-10.3s", "World") = "Wor       " (Should be "Wor       ")
>>sprintf("%5s", "VeryLongWord") = "VeryLongWord" (Should be "VeryLongWord")
>>sprintf("%.5s", "VeryLongWord") = "VeryL" (Should be "VeryL")
>>sprintf("%.2f", 3.1415) = "3.14" (Should be "3.14")
>>sprintf("%.2e", 3.1415) = "3.14e+00" (Should be "3.14e+00")
>>sprintf("%f", 3.1415) = "3.1415" (Should be "3.1415")
>>sprintf("%d", 3.1415) = "3" (Should be "3")
>>sprintf("%.0f", 3.1415) = "3" (Should be "3")
>>sprintf("%#.0f", 3.1415) = "3." (Should be "3.")
>>sprintf("%04d", 23) = "0023" (Should be "0023")
>>sprintf("%-04d", 23) = "2300" (Should be "2300") Note: Left-align + pad-w-zeros assumes want trailing zeros (otherwise use %04d)
>>sprintf("%-+04d", 23) = "+023" (Should be "+230")
>>sprintf("%i", -23) = "-23" (Should be "-23")
>>sprintf("%u", -23) = "23" (Should be "23")
>>sprintf("%+d", 23) = "+23" (Should be "+23")
>>sprintf("% d", 23) = " 23" (Should be " 23")
>>sprintf("% d", -23) = "-23" (Should be "-23")
>>sprintf("%o", 9) = "11" (Should be "11")
>>sprintf("%x", 111) = "6f" (Should be "6f")
>>sprintf("%X", 111) = "6F" (Should be "6F")
>>sprintf("%g", 0.000001) = "1.0e-06" (Should be "1.0e-06")
>>sprintf("%g", 0.01) = "0.01" (Should be "0.01")

>>sprintf("File%05d_%04d-%02d-%02d.%s", 3, 2019, 2, 10, "dat")
"File00003_2019-02-10.dat"
"File00003_2019-02-10.dat"(Should be)

IDENTIFIERS
>>sprintf("Word1=%3$s, Word2=%1$s, Word3=%2$s", "Arg1", "Arg2", "Arg3")
"Word1=Arg3, Word2=Arg1, Word3=Arg2"
"Word1=Arg3, Word2=Arg1, Word3=Arg2"(Should be)

>>sprintf("Word1=%3$s, Word2=%s, Word3=%s, Word4=%s", "Arg1", "Arg2", "Arg3")
"Word1=Arg3, Word2=Arg1, Word3=Arg2, Word4=Arg3"
"Word1=Arg3, Word2=Arg1, Word3=Arg2, Word4=Arg3"(Should be)

BREAK LINES and SPECIAL CHARACTERS
>>sprintf("LINE1\nLINE2", )
"LINE1
LINE2"
"LINE1
LINE2"(Should be)

>>sprintf("The sprintf format should be %%03d followed by a \\n", )
"The sprintf format should be %03d followed by a \n"
"The sprintf format should be %03d followed by a \n"(Should be)

>>sprintf("Tab test\tHere, Backspace Here\b", )
"Tab test   Here, Backspace Here"
"Tab test   here, Backspace here"(Should be)

TABLE ALIGNMENT - RIGHT
>>sprintf("Index  Value1   Value2")
  sprintf("%5d  %6.2f   %6.2f", 1, 0.2, 5.7)
  sprintf("%5d  %6.2f   %6.2f", 2, 10.2, -15)
"Index  Value1   Value2"
"    1    0.20     5.70"
"    2   10.20   -15.00"

TABLE ALIGNMENT - LEFT
>>sprintf("Index  Value1   Value2")
  sprintf("%-5d  %-6.2f   %-6.2f", 1, 0.2, 5.7)
  sprintf("%-5d  %-6.2f   %-6.2f", 2, 10.2, -5.7)
"Index  Value1   Value2"
"1      0.20     5.70  "
"2      10.20    -5.70 "

INPUT ERRORS
>>sprintf("%td %d", 5, 34)
"%td 5"
"%td 5"(Should be) Note: error doesn't consume input

>>sprintf("The bank rate is 15% and rising", )
"The bank rate is 15% and rising"
"The bank rate is 15% and rising"(Should be) Note: if error string returned it would be "The bank rate is 15<ERROR> rising"

INPUT ARGS ON DIFFERENT FORMS
>>Args = Array(2, "World")
>>sprintf("Hello %d the %s", Args)
"Hello 2 the World"
"Hello 2 the World"(Should be)

>>sprintf("Hello %d the %s", 2, "World")
"Hello 2 the World"
"Hello 2 the World"(Should be)
```
