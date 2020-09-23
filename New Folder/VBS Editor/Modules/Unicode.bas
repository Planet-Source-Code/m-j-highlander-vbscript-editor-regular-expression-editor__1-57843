Attribute VB_Name = "Unicode_Check"
'Info:
' BOM stands for Byte-Order Mark
' FF FE :           Unicode BOM (little-endian as in Windows)
' FE FF :           Unicode BOM (big-endian as in Mac)
' EF BB BF :        UTF-8 encoding  (UTF is identical to ASCII for alphanumeric chars, and uses 2 or more bytes to encode symbols...
' 2B 2F 76 38 2D :  UTF-7 BOM (not supported on Windows)

Option Explicit

'BOOL IsTextUnicode(
'  CONST VOID* pBuffer, // input buffer to be examined
'  int cb,             // size of input buffer
'  LPINT lpi );        // options

'lpi:[in/out] On input, specifies the tests to be applied to the input buffer text.
'On output, receives the results of the specified tests:
'1 if the contents of the buffer pass a test, zero for failure.
'Only flags that are set upon input to the function are significant upon output.
'If lpi is NULL, the function uses all available tests to determine whether the data in the buffer
'is likely to be Unicode text

'Returns a nonzero value if the data in the buffer passes the specified tests.
'Returns zero if the data in the buffer does not pass the specified tests

Public Declare Function IsTextUnicode Lib "advapi32" (lpBuffer As Any, ByVal cb As Long, lpi As Long) As Long

'The text is Unicode, and contains only zero-extended ASCII values/characters
Public Const IS_TEXT_UNICODE_ASCII16 = &H1

'Same as the preceding, except that the Unicode text is byte-reversed
Public Const IS_TEXT_UNICODE_REVERSE_ASCII16 = &H10

'The text is probably Unicode, with the determination made by applying statistical analysis.
'Absolute certainty is not guaranteed
Public Const IS_TEXT_UNICODE_STATISTICS = &H2

'Same as the preceding, except that the probably-Unicode text is byte-reversed
Public Const IS_TEXT_UNICODE_REVERSE_STATISTICS = &H20

'The text contains Unicode representations of one or more of these nonprinting characters:
'RETURN, LINEFEED, SPACE, CJK_SPACE, TAB
Public Const IS_TEXT_UNICODE_CONTROLS = &H4

'Same as the preceding, except that the Unicode characters are byte-reversed
Public Const IS_TEXT_UNICODE_REVERSE_CONTROLS = &H40

'The text contains the Unicode byte-order mark (BOM) 0xFEFF as its first character
Public Const IS_TEXT_UNICODE_SIGNATURE = &H8

'The text contains the Unicode byte-reversed byte-order mark (Reverse BOM) 0xFFFE as its first character
Public Const IS_TEXT_UNICODE_REVERSE_SIGNATURE = &H80

'The text contains one of these Unicode-illegal characters:
'embedded Reverse BOM, UNICODE_NUL, CRLF (packed into one WORD), or 0xFFFF
Public Const IS_TEXT_UNICODE_ILLEGAL_CHARS = &H100

'The number of characters in the string is odd.
'A string of odd length cannot (by definition) be Unicode text
Public Const IS_TEXT_UNICODE_ODD_LENGTH = &H200


Public Const IS_TEXT_UNICODE_DBCS_LEADBYTE = &H400

'The text contains null bytes, which indicate non-ASCII text
Public Const IS_TEXT_UNICODE_NULL_BYTES = &H1000

'This flag constant is a combination of IS_TEXT_UNICODE_ASCII16, IS_TEXT_UNICODE_STATISTICS, IS_TEXT_UNICODE_CONTROLS, IS_TEXT_UNICODE_SIGNATURE
Public Const IS_TEXT_UNICODE_UNICODE_MASK = &HF

'This flag constant is a combination of IS_TEXT_UNICODE_REVERSE_ASCII16, IS_TEXT_UNICODE_REVERSE_STATISTICS, IS_TEXT_UNICODE_REVERSE_CONTROLS, IS_TEXT_UNICODE_REVERSE_SIGNATURE
Public Const IS_TEXT_UNICODE_REVERSE_MASK = &HF0

'This flag constant is a combination of IS_TEXT_UNICODE_ILLEGAL_CHARS, IS_TEXT_UNICODE_ODD_LENGTH, and two currently unused bit flags
Public Const IS_TEXT_UNICODE_NOT_UNICODE_MASK = &HF00

'This flag constant is a combination of IS_TEXT_UNICODE_NULL_BYTES and three currently unused bit flags
Public Const IS_TEXT_UNICODE_NOT_ASCII_MASK = &HF000
Public Function IsUnicodeStr(ByVal sBuffer As String) As Boolean
'Returns True if sBuffer evaluates to a Unicode string
Dim dwRtnFlags As Long

dwRtnFlags = IS_TEXT_UNICODE_UNICODE_MASK
IsUnicodeStr = IsTextUnicode(ByVal sBuffer, Len(sBuffer), dwRtnFlags)

End Function
