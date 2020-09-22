Attribute VB_Name = "DiskSubs"
Option Explicit
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" _
Alias "GetDiskFreeSpaceExA" _
(ByVal lpRootPathName As String, _
lpFreeBytesAvailableToCaller As Currency, _
lpTotalNumberOfBytes As Currency, _
lpTotalNumberOfFreeBytes As Currency) As Long

Dim r As Long, BytesFreeToCalller As Currency, TotalBytes As Currency
Dim TotalFreeBytes As Currency, TotalBytesUsed As Currency
Dim TNB As Double
Dim TFB As Double
Dim FreeBytes As Long
Dim DriveLetter As String
Dim DLetter As String
Dim spaceInt As Integer

Global IveWarned As Boolean ' Already notified the user of low disk space

Type DISKSPACEINFO
    RootPath As String * 3
    FreeBytes As Long
    TotalBytes As Long
    FreePcnt As Single
    UsedPcnt As Single
End Type


Public Function GetDiskInfo(RootS As String)
If NoDiskSpaceCheck = True Then
Debug.Print "DiskSpaceCheck Canceled."
Else
DriveLetter = RootS

spaceInt = InStr(DriveLetter, " ")
If spaceInt > 0 Then DriveLetter = Left$(DriveLetter, spaceInt - 1)

If Right$(DriveLetter, 1) <> "\" Then DriveLetter = DriveLetter & "\"
DLetter = Left(UCase(DriveLetter), 1)

    Call GetDiskFreeSpaceEx(DriveLetter, BytesFreeToCalller, TotalBytes, TotalFreeBytes)
TNB = TotalBytes * 10000
    TFB = (TotalBytes - TotalFreeBytes) * 10000
GetDiskInfo = Int(Int(Format$(BytesFreeToCalller * 10000, "###########0")) / 1024000)
Debug.Print "Drive C Free Disk Space: " & GetDiskInfo & " MB"
End If
End Function
