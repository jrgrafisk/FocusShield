Attribute VB_Name = "modI18n"
Option Explicit
' Danish letters in every user-facing string in this project are written
' as plain-ASCII markers and decoded here at runtime with ChrW$, instead
' of being typed directly as accented characters in the source.
'
' Why: the VBA editor's "Import File" is not reliably UTF-8-safe. A
' literal ae/o-slash/a-ring typed straight into a .bas file can silently
' turn into mojibake depending on the Windows code page in effect when the
' file is imported. ChrW$(230) always produces the correct character at
' run time, regardless of how the source file itself was read in - so it
' is the only way to guarantee correct Danish text without requiring the
' user to fight file encodings during setup.
'
' Usage: DK("L~aa~n og afdrag") -> "Laan og afdrag" with proper letters.
'   ~ae~ -> ae (U+00E6)   ~AE~ -> AE (U+00C6)
'   ~oe~ -> o-slash (U+00F8)   ~OE~ -> O-slash (U+00D8)
'   ~aa~ -> a-ring (U+00E5)   ~AA~ -> A-ring (U+00C5)
'   ~ee~ -> e-acute (U+00E9)  (only lower-case needed anywhere in this project)

Public Function DK(ByVal s As String) As String
    s = Replace(s, "~ae~", ChrW$(230))
    s = Replace(s, "~oe~", ChrW$(248))
    s = Replace(s, "~aa~", ChrW$(229))
    s = Replace(s, "~ee~", ChrW$(233))
    s = Replace(s, "~AE~", ChrW$(198))
    s = Replace(s, "~OE~", ChrW$(216))
    s = Replace(s, "~AA~", ChrW$(197))
    DK = s
End Function
