'#Global Declaration
Const ForReading = 1
Const ForWriting = 2

'=========== << Main Function >> =========
Function Main()
BrowseConfFile() 			'Browse for the Configuration File 
End Function 
'%#$%#%#======%##% >> END OF MAIN FUNCTION CODE << %%%%%#%=====#%#%#%#%#%

'************ ### Function Definitions #### *****************

'###### FN-SUB :: Browse for the Config File ] ###############

Sub BrowseConfFile()
'###### Get the Configuration File Path ###############
Set wShell=CreateObject("WScript.Shell")
Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
ConfFilePath = Cstr(oExec.StdOut.ReadLine)
'wscript.echo ConfFilePath

if ConfFilePath <> "" then
ReadConfFile(ConfFilePath)  ' Sub-Fn.Call
Else
wscript.echo "Operation cancelled.!"
End if

'CLEAR MEM
Set wShell= Nothing
Set oExec =Nothing
End Sub

'###### FN-SUB :: Read the Configuration File[ Reads the FilePath,Text needs to be replaced] ###############
Sub ReadConfFile(ConfFilePath)
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set openFile = objFSO.OpenTextFile(ConfFilePath,ForReading)

txtRe = openFile.ReadAll
Dim arr : arr = Split(txtRe,vbCrlf)
Dim i : i=UBound(arr)-1

do while i > 0 		'Looping for getting all entries
WrFilePath = arr(i-2)
ActTxt = arr(i-1)
RepTxt =arr(i)	
i=i-3

'wscript.echo WrFilePath&"--"&ActTxt&"--"&RepTxt
Call WriteFile(WrFilePath,ActTxt,RepTxt)
loop

openFile.Close
'CLEAR MEM
Set objFSO = Nothing
Set openFile = Nothing
End Sub

'###### FN-SUB :: Replace the text / Get the no. of replacements ############### 

Sub WriteFile(WrFilePath,ActTxt,RepTxt)
'wscript.echo WrFilePath
Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FileExists(WrFilePath) then	'----------------- IF File found check-----

Set openFile = objFSO.OpenTextFile(WrFilePath,ForReading)
Set getFile = objFSO.GetFile(WrFilePath) ' Only for Filesize check

'=== Text Check =======
If getFile.Size = 0 Then
wscript.echo("Replace Failed-Empty Source File : "&WrFilePath)
openFile.Close
else
'--- Take BackUp before Replace
CpFilePath = "Old_"&WrFilePath&".bak"
objFSO.CopyFile WrFilePath,CpFilePath,True

'---- Replace fn. Starts
txtWr = openFile.ReadAll
openFile.Close
NewFile = Replace(txtWr,ActTxt,RepTxt)

'== Write the new text ======
Set wrFile = objFSO.OpenTextFile(WrFilePath,ForWriting)
wrFile.WriteLine NewFile
wrFile.Close

'==== Get the Replaced Places ========
Set openFile = objFSO.OpenTextFile(WrFilePath,ForReading)
Dim RepCnt 
'Read Line by Line
do until openFile.AtEndOfStream
txtL=openFile.ReadLine
tmpstr=Cstr(txtL)
'Check only for non-null line
if tmpstr <> "" then
'wscript.echo tmpstr
'Get Count of Replacement/presence
Dim arr : arr = split(tmpStr,RepTxt)
RepCnt=RepCnt+UBound(arr)
End If
loop
Msgbox "TEXT=["&RepTxt&"] is replaced/available now for : "&RepCnt&"  times"

Set openFile = Nothing
Set objFSO = Nothing
End if

'----------------- IF File found check-----

Else
wscript.echo "No write done,File not found : "&WrFilePath
End if ' File Exist check

End Sub

'************************* END OF FN.DEFn****************************************

'=======$$$$$$$ ## Main Function ## $$$$$$$========
Call Main()