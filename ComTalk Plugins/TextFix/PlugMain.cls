VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextFixMain"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Property Get FriendlyName() As String
FriendlyName = "TextFix" ' Friendly Name (Shows in Plugin Menu)
End Property

Public Function DoAction(Parametres)
' Actions Executed
Static TMPtexttofilter As String

TMPtexttofilter = Parametres
DoAction = FilterText(TMPtexttofilter)
End Function

Public Function KillMe()
On Error Resume Next
Unload TextFixIcon
End
End Function

Public Property Get ShowInMenu() As Boolean
ShowInMenu = False ' Show Plugin In Menu
End Property

Public Property Get PassBeforeSay() As Boolean
PassBeforeSay = True ' Pass Text to be spoken through plugin "DoAction" first
End Property

Function FilterText(texttofilter As String)
WTemp = ""
STemp = ""

While Right(texttofilter, 1) = " "
texttofilter = Mid(texttofilter, 0, Len(texttofilter) - 1)
Wend

WordList = Split(texttofilter, " ")
For Each Word In WordList
STemp = ""
LSTemp = ""
Word = LCase(Word)

If Right(Word, 1) = "!" Or Right(Word, 1) = "?" Or Right(Word, 1) = "#" Or Right(Word, 1) = "&" Or Right(Word, 1) = "*" Or Right(Word, 1) = "," Or Right(Word, 1) = "(" Or Right(Word, 1) = ")" Or Right(Word, 1) = "$" Or Right(Word, 1) = "~" Or Right(Word, 1) = "-" Or Right(Word, 1) = "_" Or Right(Word, 1) = "+" Or Right(Word, 1) = "\" Or Right(Word, 1) = "/" Or Right(Word, 1) = "." Or Right(Word, 1) = Chr(34) Then
STemp = Right(Word, 1)
Word = Mid(Word, 1, Len(Word) - 1)
End If

If Mid(Word, 1, 1) = Chr(34) Then
LSTemp = Mid(Word, 1, 1)
End If

  If Word = "c" Then Word = "see"
  If Word = "ic" Then Word = "i see"
  If Word = "oic" Then Word = "oh i see"
  If Word = "u" Then Word = "you"
  If Word = "ya" Then Word = "you"
  If Word = "ur" Then Word = "you are"
  If Word = "cu" Then Word = "see you"
  If Word = "im" Then Word = "i am"
  If Word = "lol" Then Word = "laughs out loud"
  If Word = "rotf" Then Word = "rolling on the floor"
  If Word = "rotflmao" Then Word = "rolling on the floor laughing my ass off"
  If Word = "brb" Then Word = "be right back"
  If Word = "ctn" Then Word = "can't talk now"
  If Word = "dunno" Then Word = "don't know"
  If Word = "dont" Then Word = "don't"
  If Word = "cant" Then Word = "can't"
  If Word = "wtf" Then Word = "what the fuck"
  If Word = "r" Then Word = "are"
  If Word = "b" Then Word = "be"
  If Word = "whos" Then Word = "who is"
  If Word = "goin" And Left(Word, 5) <> "going" Then Word = "going"
  If Word = "who's" Then Word = "who is"
  If Word = "whats" Then Word = "what is"
  If Word = "what's" Then Word = "what is"
  If Word = "whens" Then Word = "when is"
  If Word = "when's" Then Word = "when is"
  If Word = "wheres" Then Word = "where is"
  If Word = "where's" Then Word = "where is"
  If Word = "whys" Then Word = "why is"
  If Word = "why's" Then Word = "why is"
  If Word = "hows" Then Word = "how is"
  If Word = "how's" Then Word = "how is"
  If Word = "whatcha" Then Word = "what are you"
  If Word = "whatya" Then Word = "what are you"
  If Word = "whacha" Then Word = "what are you"
  If Word = "wazzap" Then Word = "what is up"
  If Word = "wussup" Then Word = "what is up"
  If Word = "sup" Then Word = "what is up"
  If Word = "wuzzup" Then Word = "what is up"
  If Word = "wassup" Then Word = "what is up"
  If Word = "doin" Then Word = "doing"
  If Word = "(r)" Then Word = "[rainbow symbol]"
  If Word = ":-)" Then Word = "[happy face]"
  If Word = ":-d" Then Word = "[smiling face]"
  If Word = ":-(" Then Word = "[sad face]"
  If Word = ":'(" Then Word = "[crying face]"
  If Word = ":-o" Then Word = "[mouth-open face]"
  If Word = ":)" Then Word = "[happy face]"
  If Word = ":d" Then Word = "[smiling face]"
  If Word = ":(" Then Word = "[sad face]"
  If Word = ":o" Then Word = "[mouth-open face]"
  If Word = "=)" Then Word = "[happy face]"
  If Word = "=d" Then Word = "[smiling face]"
  If Word = "=(" Then Word = "[sad face]"
  If Word = "=|" Then Word = "[normal face]"
  If Word = "(6)" Then Word = "[devil face]"
  If Word = "(8)" Then Word = "[music note]"
  If Word = "lve" Then Word = "love"
  If Word = "skool" Then Word = "school"
  If Word = "skl" Then Word = "school"
  If Word = "hme" Then Word = "home"
  If Word = "prk" Then Word = "park"
  If Word = "wnt" Then Word = "went"
  If Word = "nqr" Then Word = "not quite right"
  If Word = "abbr" Then Word = "abbreviation"
  If Word = "cnt" Then Word = "can't"
  If Word = "hve" Then Word = "have"
  If Word = "brcs" Then Word = "braces"
  If Word = "blve" Then Word = "believe"
  If Word = "wld" Then Word = "would"
  If Word = "tht" Then Word = "that"
  If Word = "bh" Then Word = "bloody hell"
  If Word = "msc" Then Word = "miscellaneous"
  If Word = "yslf" Then Word = "yourself"
  If Word = "bboat" Then Word = "bannana boat"
  If Word = "hll" Then Word = "hell"
  If Word = "hvn" Then Word = "heaven"
  If Word = "hevn" Then Word = "heaven"
  If Word = "bk" Then Word = "book"
  If Word = "#" Then Word = "number"
  If Word = "hlp" Then Word = "help"
  If Word = "cm" Then Word = "come"
  If Word = "cmptr" Then Word = "computer"
  If Word = "mntr" Then Word = "monitor"
  If Word = "kbrd" Then Word = "keyboard"
  If Word = "pncl" Then Word = "pencil"
  If Word = "prnt" Then Word = "print"
  If Word = "prntr" Then Word = "printer"
  If Word = "scnnr" Then Word = "scanner"
  If Word = "cpy" Then Word = "copy"
  If Word = "y" Then Word = "why"
  If Word = "g2g" Then Word = "got to go"
  If Word = "hve" Then Word = "have"
  If Word = "bt" Then Word = "but"
  If Word = "plz" Then Word = "please"
  If Word = "stpd" Then Word = "stupid"
  If Word = "hi" Then Word = "hello"
  If Word = "cya" Then Word = "see you"
  If Word = "1/2" Then Word = "half"
  If Word = "1/4" Then Word = "quarter"
  
If WTemp <> "" Then
WTemp = WTemp & " " & LSTemp & Word & STemp
Else
WTemp = LSTemp & Word & STemp
End If
Next

FilterText = StrConv(WTemp, vbProperCase)
End Function

Public Property Get PClassName() As String
PClassName = PACN
End Property

Public Property Get PMadeBy() As String
PMadeBy = MEMADEBY
End Property

Public Property Get PRequirements() As String
PRequirements = MEREQIREMENTS
End Property

