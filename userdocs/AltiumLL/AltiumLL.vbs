Dim InstalledDir
Dim SchLib
Dim SchComponent
Dim cbFilename
Dim prtName
Dim username
Dim password
Dim CurrentSCH
Dim SchDocFile
Dim stpFileName
Dim AppVersion
Dim decChar

AppVersion = "2.3"
decChar = Mid(FormatNumber(0.1,1,true,false,-2), 2, 1)

Sub ProcessExternalCB()
	Dim fso
	Dim f
	Dim cbPath
	Dim WshShell
	Dim InstallFolder
	Dim CurrentSheet

	Set WshShell = CreateObject("WScript.Shell")
	InstallFolder = WshShell.SpecialFolders("MyDocuments")
	InstalledDir = InstallFolder & "\AltiumLL"

	Set fso = CreateObject("Scripting.FileSystemObject")

	txt_PcbLib.Text = InstalledDir & "\SamacSys.PcbLib"
	txt_SchLib.Text = InstalledDir & "\SamacSys.SchLib"

	If Not fso.FileExists(txt_PcbLib.Text) Then
		MsgBox "PcbLib not found: " & txt_PcbLib.Text, vbSystemModal, "Altium Library Loader"
		Exit Sub
	End If

	If Not fso.FileExists(txt_SchLib.Text) Then
		MsgBox "SchLib not found: " & txt_SchLib.Text, vbSystemModal, "Altium Library Loader"
		Exit Sub
	End If

	Client.StartServer("SCH")
	Set CurrentSheet = SchServer.GetCurrentSchDocument

	If CurrentSheet Is Nothing Then
		MsgBox "A schematic document (*.SchDoc) must be opened or selected before ECAD Model can be added to design.", vbSystemModal, "Altium Library Loader"
		Exit Sub
	End If

	If UCase(Right(CurrentSheet.GetState_DocumentName, 7)) <> ".SCHDOC" Then
		MsgBox "A schematic document (*.SchDoc) must be opened or selected before ECAD Model can be added to design.", vbSystemModal, "Altium Library Loader"
		Exit Sub
	End If

	SchDocFile = CurrentSheet.GetState_DocumentName

	If Not fso.FileExists(InstalledDir & "\Temp\current_cb.txt") Then
		MsgBox "CB pointer file not found: " & InstalledDir & "\Temp\current_cb.txt", vbSystemModal, "Altium Library Loader"
		Exit Sub
	End If

	Set f = fso.OpenTextFile(InstalledDir & "\Temp\current_cb.txt")
	cbPath = f.ReadLine
	f.Close

	If Not fso.FileExists(cbPath) Then
		MsgBox "CB file not found: " & cbPath, vbSystemModal, "Altium Library Loader"
		Exit Sub
	End If

	ProcessCB cbPath
End Sub

Sub EnsureSettingsFile()
	Dim fso
	Dim WshShell
	Dim InstallFolder
	Dim TXTFileName
	Dim DownloadsFolder
	Dim TXTFile

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set WshShell = CreateObject("WScript.Shell")

	InstallFolder = WshShell.SpecialFolders("MyDocuments")
	InstalledDir = InstallFolder & "\AltiumLL"
	DownloadsFolder = GetDownloadsFolder()
	TXTFileName = InstalledDir & "\AltiumLL.txt"

	If Not fso.FolderExists(InstalledDir) Then
		fso.CreateFolder(InstalledDir)
	End If

	If Not fso.FolderExists(InstalledDir & "\Temp") Then
		fso.CreateFolder(InstalledDir & "\Temp")
	End If

	If Not fso.FolderExists(InstalledDir & "\SamacSys_Images") Then
		fso.CreateFolder(InstalledDir & "\SamacSys_Images")
	End If

	If Not fso.FileExists(TXTFileName) Then
		Set TXTFile = fso.CreateTextFile(TXTFileName, True)

		TXTFile.WriteLine ""
		TXTFile.WriteLine ""
		TXTFile.WriteLine DownloadsFolder
		TXTFile.WriteLine InstalledDir & "\SamacSys.SchLib"
		TXTFile.WriteLine InstalledDir & "\SamacSys.PcbLib"
		TXTFile.WriteLine "False"
		TXTFile.WriteLine ""
		TXTFile.WriteLine ""
		TXTFile.WriteLine "True"
		TXTFile.WriteLine "True"

		TXTFile.Close
	End If
End Sub

Function GetDownloadsFolder()
	Dim sh
	Dim folder
	Dim WshShell

	GetDownloadsFolder = vbNullString

	On Error Resume Next

	Set sh = CreateObject("Shell.Application")

	Set folder = sh.NameSpace("shell:Downloads")
	If Not folder Is Nothing Then
		GetDownloadsFolder = folder.Self.Path
	End If

	If GetDownloadsFolder = vbNullString Then
		Set folder = sh.NameSpace(&H374DE290)
		If Not folder Is Nothing Then
			GetDownloadsFolder = folder.Self.Path
		End If
	End If

	If GetDownloadsFolder = vbNullString Then
		Set WshShell = CreateObject("WScript.Shell")
		GetDownloadsFolder = WshShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Downloads"
	End If

	On Error GoTo 0
End Function

Sub CreateCourtyardLine(x1,y1,x2,y2,width)
	Dim atrack
	Dim PCBLib
	Dim footprint

	Set PCBLib = PCBServer.GetCurrentPCBLibrary
	Set footprint = PCBLib.CurrentComponent

	x1 = replace(x1,".",decChar)
	y1 = replace(y1,".",decChar)
	x2 = replace(x2,".",decChar)
	y2 = replace(y2,".",decChar)
	width = replace(width,".",decChar)

	atrack = PCBServer.PCBObjectFactory(eTrackObject,eNoDimension,eCreate_Default)
	atrack.width = mmstocoord(width)
	atrack.layer = eMechanical15
	atrack.x1 = mmstocoord(x1)+footprint.x
	atrack.x2 = mmstocoord(x2)+footprint.x
	atrack.y1 = mmstocoord(y1)+footprint.y
	atrack.y2 = mmstocoord(y2)+footprint.y
	if footprint Is Nothing Then Exit Sub
	PCBServer.GetCurrentPCBBoard.AddPCBObject(atrack)
	'PCBServer.SendMessageToRobots footprint.I_ObjectAddress,c_Broadcast,PCBM_BoardRegisteration,atrack.I_ObjectAddress
	PCBServer.PostProcess
End Sub

Sub CreateCourtyardCircle(x,y,rad,width)
	Dim acircle
	Dim PCBLib
	Dim footprint

	Set PCBLib = PCBServer.GetCurrentPCBLibrary
	Set footprint = PCBLib.CurrentComponent

	x = replace(x,".",decChar)
	y = replace(y,".",decChar)
	rad = replace(rad,".",decChar)
	width = replace(width,".",decChar)

	acircle = PCBServer.PCBObjectFactory(eArcObject,eNoDimension,eCreate_Default)
	acircle.XCenter = mmstocoord(x)+footprint.x
	acircle.YCenter = mmstocoord(y)+footprint.y
	acircle.Radius = mmstocoord(rad)
	acircle.LineWidth = mmstocoord(width)
	acircle.StartAngle = 0
	acircle.EndAngle = 360
	acircle.layer = eMechanical15

	if footprint Is Nothing Then Exit Sub
	PCBServer.GetCurrentPCBBoard.AddPCBObject(acircle)
	'PCBServer.SendMessageToRobots footprint.I_ObjectAddress,c_Broadcast,PCBM_BoardRegisteration,acircle.I_ObjectAddress
	PCBServer.PostProcess
End Sub

Sub CreateSilkscreenLine(x1,y1,x2,y2,width)
	Dim atrack
	Dim PCBLib
	Dim footprint

	Set PCBLib = PCBServer.GetCurrentPCBLibrary
	Set footprint = PCBLib.CurrentComponent

	x1 = replace(x1,".",decChar)
	y1 = replace(y1,".",decChar)
	x2 = replace(x2,".",decChar)
	y2 = replace(y2,".",decChar)
	width = replace(width,".",decChar)

	atrack = PCBServer.PCBObjectFactory(eTrackObject,eNoDimension,eCreate_Default)
	atrack.width = mmstocoord(width)
	atrack.layer = eTopoverlay
	atrack.x1 = mmstocoord(x1)+footprint.x
	atrack.x2 = mmstocoord(x2)+footprint.x
	atrack.y1 = mmstocoord(y1)+footprint.y
	atrack.y2 = mmstocoord(y2)+footprint.y
	if footprint Is Nothing Then Exit Sub
	PCBServer.GetCurrentPCBBoard.AddPCBObject(atrack)
	'PCBServer.SendMessageToRobots footprint.I_ObjectAddress,c_Broadcast,PCBM_BoardRegisteration,atrack.I_ObjectAddress
	PCBServer.PostProcess
End Sub

Sub CreateSilkscreenCircle(x,y,rad,width)
	Dim acircle
	Dim PCBLib
	Dim footprint

	Set PCBLib = PCBServer.GetCurrentPCBLibrary
	Set footprint = PCBLib.CurrentComponent

	x = replace(x,".",decChar)
	y = replace(y,".",decChar)
	rad = replace(rad,".",decChar)
	width = replace(width,".",decChar)

	acircle = PCBServer.PCBObjectFactory(eArcObject,eNoDimension,eCreate_Default)
	acircle.XCenter = mmstocoord(x)+footprint.x
	acircle.YCenter = mmstocoord(y)+footprint.y
	acircle.Radius = mmstocoord(rad)
	acircle.LineWidth = mmstocoord(width)
	acircle.StartAngle = 0
	acircle.EndAngle = 360
	acircle.layer = eTopoverlay

	if footprint Is Nothing Then Exit Sub
	PCBServer.GetCurrentPCBBoard.AddPCBObject(acircle)
	'PCBServer.SendMessageToRobots footprint.I_ObjectAddress,c_Broadcast,PCBM_BoardRegisteration,acircle.I_ObjectAddress
	PCBServer.PostProcess
End Sub

Sub CreateSilkscreenArc(x,y,rad,startAngle,endAngle,width)
	Dim aarc
	Dim footprint

	Set PCBLib = PCBServer.GetCurrentPCBLibrary
	Set footprint = PCBLib.CurrentComponent

	x = replace(x,".",decChar)
	y = replace(y,".",decChar)
	rad = replace(rad,".",decChar)
	startAngle = replace(startAngle,".",decChar)
	endAngle = replace(endAngle,".",decChar)
	width = replace(width,".",decChar)

	aarc = PCBServer.PCBObjectFactory(eArcObject,eNoDimension,eCreate_Default)
	aarc.XCenter = mmstocoord(x)+footprint.x
	aarc.YCenter = mmstocoord(y)+footprint.y
	aarc.Radius = mmstocoord(rad)
	aarc.LineWidth = mmstocoord(width)
	aarc.StartAngle = startAngle
	aarc.EndAngle = endAngle
	aarc.layer = eTopoverlay

	if footprint Is Nothing Then Exit Sub
	PCBServer.GetCurrentPCBBoard.AddPCBObject(aarc)
	'PCBServer.SendMessageToRobots footprint.I_ObjectAddress,c_Broadcast,PCBM_BoardRegisteration,aarc.I_ObjectAddress
	PCBServer.PostProcess
End Sub

Sub CreateAssemblyLine(x1,y1,x2,y2,width)
	Dim atrack
	Dim PCBLib
	Dim footprint

	Set PCBLib = PCBServer.GetCurrentPCBLibrary
	Set footprint = PCBLib.CurrentComponent

	x1 = replace(x1,".",decChar)
	y1 = replace(y1,".",decChar)
	x2 = replace(x2,".",decChar)
	y2 = replace(y2,".",decChar)
	width = replace(width,".",decChar)

	atrack = PCBServer.PCBObjectFactory(eTrackObject,eNoDimension,eCreate_Default)
	atrack.width = mmstocoord(width)
	atrack.layer = eMechanical13
	atrack.x1 = mmstocoord(x1)+footprint.x
	atrack.x2 = mmstocoord(x2)+footprint.x
	atrack.y1 = mmstocoord(y1)+footprint.y
	atrack.y2 = mmstocoord(y2)+footprint.y
	if footprint Is Nothing Then Exit Sub
	PCBServer.GetCurrentPCBBoard.AddPCBObject(atrack)
	'PCBServer.SendMessageToRobots footprint.I_ObjectAddress,c_Broadcast,PCBM_BoardRegisteration,atrack.I_ObjectAddress
	PCBServer.PostProcess
End Sub

Sub CreateAssemblyCircle(x,y,rad,width)
	Dim acircle
	Dim PCBLib
	Dim footprint

	Set PCBLib = PCBServer.GetCurrentPCBLibrary
	Set footprint = PCBLib.CurrentComponent

	x = replace(x,".",decChar)
	y = replace(y,".",decChar)
	rad = replace(rad,".",decChar)
	width = replace(width,".",decChar)

	acircle = PCBServer.PCBObjectFactory(eArcObject,eNoDimension,eCreate_Default)
	acircle.XCenter = mmstocoord(x)+footprint.x
	acircle.YCenter = mmstocoord(y)+footprint.y
	acircle.Radius = mmstocoord(rad)
	acircle.LineWidth = mmstocoord(width)
	acircle.StartAngle = 0
	acircle.EndAngle = 360
	acircle.layer = eMechanical13

	if footprint Is Nothing Then Exit Sub
	PCBServer.GetCurrentPCBBoard.AddPCBObject(acircle)
	'PCBServer.SendMessageToRobots footprint.I_ObjectAddress,c_Broadcast,PCBM_BoardRegisteration,acircle.I_ObjectAddress
	PCBServer.PostProcess
End Sub

Sub CreateAssemblyArc(x,y,rad,startAngle,endAngle,width)
	Dim aarc
	Dim footprint

	Set PCBLib = PCBServer.GetCurrentPCBLibrary
	Set footprint = PCBLib.CurrentComponent

	x = replace(x,".",decChar)
	y = replace(y,".",decChar)
	rad = replace(rad,".",decChar)
	startAngle = replace(startAngle,".",decChar)
	endAngle = replace(endAngle,".",decChar)
	width = replace(width,".",decChar)

	aarc = PCBServer.PCBObjectFactory(eArcObject,eNoDimension,eCreate_Default)
	aarc.XCenter = mmstocoord(x)+footprint.x
	aarc.YCenter = mmstocoord(y)+footprint.y
	aarc.Radius = mmstocoord(rad)
	aarc.LineWidth = mmstocoord(width)
	aarc.StartAngle = startAngle
	aarc.EndAngle = endAngle
	aarc.layer = eMechanical13

	if footprint Is Nothing Then Exit Sub
	PCBServer.GetCurrentPCBBoard.AddPCBObject(aarc)
	'PCBServer.SendMessageToRobots footprint.I_ObjectAddress,c_Broadcast,PCBM_BoardRegisteration,aarc.I_ObjectAddress
	PCBServer.PostProcess
End Sub

Sub CreateRoundedPTH(x,y,length,width,holesize,name,plated,rotation)
	Dim apad
	Dim PCBLib
	Dim footprint

	Set PCBLib = PCBServer.GetCurrentPCBLibrary
	Set footprint = PCBLib.CurrentComponent

	x = replace(x,".",decChar)
	y = replace(y,".",decChar)
	length = replace(length,".",decChar)
	width = replace(width,".",decChar)
	holesize = replace(holesize,".",decChar)
	rotation = replace(rotation,".",decChar)

	apad = pcbserver.PCBObjectFactory(ePadObject,eNoDimension,eCreate_Default)
	apad.Mode = ePadMode_Simple
	apad.name = name
	apad.HoleType = eRoundHole
	apad.HoleSize = mmstocoord(holesize)
	apad.Rotation = rotation
	apad.Plated = CInt(plated)
	apad.TopShape = eRounded
	apad.TopXSize = mmstocoord(length)
	apad.TopYSize = mmstocoord(width)
	apad.layer = eMultiLayer
	apad.x = mmstocoord(x)+footprint.x
	apad.y = mmstocoord(y)+footprint.y
	if footprint Is Nothing Then Exit Sub
	footprint.board.addpcbobject(apad)
	'PCBServer.SendMessageToRobots footprint.I_ObjectAddress,c_Broadcast,PCBM_BoardRegisteration,apad.I_ObjectAddress
	PCBServer.PostProcess
End Sub

Sub CreateRectangularPTH(x,y,length,width,holesize,name,plated,rotation)
	Dim apad
	Dim PCBLib
	Dim footprint

	Set PCBLib = PCBServer.GetCurrentPCBLibrary
	Set footprint = PCBLib.CurrentComponent

	x = replace(x,".",decChar)
	y = replace(y,".",decChar)
	length = replace(length,".",decChar)
	width = replace(width,".",decChar)
	holesize = replace(holesize,".",decChar)
	rotation = replace(rotation,".",decChar)

	apad = pcbserver.PCBObjectFactory(ePadObject,eNoDimension,eCreate_Default)
	apad.Mode = ePadMode_Simple
	apad.name = name
	apad.HoleType = eRoundHole
	apad.HoleSize = mmstocoord(holesize)
	apad.Rotation = rotation
	apad.Plated = CInt(plated)
	apad.TopShape = eRectangular
	apad.TopXSize = mmstocoord(length)
	apad.TopYSize = mmstocoord(width)
	apad.layer = eMultiLayer
	apad.x = mmstocoord(x)+footprint.x
	apad.y = mmstocoord(y)+footprint.y
	if footprint Is Nothing Then Exit Sub
	footprint.board.addpcbobject(apad)
	'PCBServer.SendMessageToRobots footprint.I_ObjectAddress,c_Broadcast,PCBM_BoardRegisteration,apad.I_ObjectAddress
	PCBServer.PostProcess
End Sub

Sub CreateRectangularSMD(x,y,length,width,holesize,name,plated,rotation)
	Dim apad
	Dim PCBLib
	Dim footprint

	Set PCBLib = PCBServer.GetCurrentPCBLibrary

	If PCBLib Is Nothing Then
		MsgBox "PCB library is not active", vbSystemModal, "Altium Library Loader"
		Exit Sub
	End If

	Set footprint = PCBLib.CurrentComponent

	If footprint Is Nothing Then
		MsgBox "PCB footprint is not active", vbSystemModal, "Altium Library Loader"
		Exit Sub
	End If

	x = Replace(x, ".", decChar)
	y = Replace(y, ".", decChar)
	length = Replace(length, ".", decChar)
	width = Replace(width, ".", decChar)
	holesize = Replace(holesize, ".", decChar)
	rotation = Replace(rotation, ".", decChar)

	Set apad = PCBServer.PCBObjectFactory(ePadObject, eNoDimension, eCreate_Default)

	If apad Is Nothing Then
		MsgBox "Failed to create pad object", vbSystemModal, "Altium Library Loader"
		Exit Sub
	End If

	apad.Mode = ePadMode_Simple
	apad.Name = name
	apad.HoleType = eRoundHole
	apad.HoleSize = MMsToCoord(holesize)
	apad.Rotation = rotation
	apad.Plated = CInt(plated)

	apad.TopShape = eRectangular
	apad.TopXSize = MMsToCoord(length)
	apad.TopYSize = MMsToCoord(width)

	apad.Layer = eTopLayer
	apad.X = MMsToCoord(x) + footprint.X
	apad.Y = MMsToCoord(y) + footprint.Y

	PCBServer.GetCurrentPCBBoard.AddPCBObject(apad)

	'PCBServer.SendMessageToRobots _
	'footprint.I_ObjectAddress, _
	'c_Broadcast, _
	'PCBM_BoardRegisteration, _
	'apad.I_ObjectAddress
End Sub

Sub CreateRoundedRectangularSMD(x,y,length,width,holesize,name,plated,rotation)
	Dim apad
	Dim PCBLib
	Dim footprint

	Set PCBLib = PCBServer.GetCurrentPCBLibrary
	Set footprint = PCBLib.CurrentComponent

	x = replace(x,".",decChar)
	y = replace(y,".",decChar)
	length = replace(length,".",decChar)
	width = replace(width,".",decChar)
	holesize = replace(holesize,".",decChar)
	rotation = replace(rotation,".",decChar)

	apad = pcbserver.PCBObjectFactory(ePadObject,eNoDimension,eCreate_Default)
	apad.Mode = ePadMode_Simple
	apad.name = name
	apad.HoleType = eRoundHole
	apad.HoleSize = mmstocoord(holesize)
	apad.Rotation = rotation
	apad.Plated = CInt(plated)
	apad.TopShape = eRoundedRectangular
	apad.StackCRPctOnLayer(eToplayer) = 100
	apad.TopXSize = mmstocoord(length)
	apad.TopYSize = mmstocoord(width)
	apad.layer = eToplayer
	apad.x = mmstocoord(x)+footprint.x
	apad.y = mmstocoord(y)+footprint.y
	if footprint Is Nothing Then Exit Sub
	footprint.board.addpcbobject(apad)
	'PCBServer.SendMessageToRobots footprint.I_ObjectAddress,c_Broadcast,PCBM_BoardRegisteration,apad.I_ObjectAddress
	PCBServer.PostProcess
End Sub

Sub CreateRoundedSMD(x,y,length,width,holesize,name,plated,rotation)
	Dim apad
	Dim PCBLib
	Dim footprint

	Set PCBLib = PCBServer.GetCurrentPCBLibrary
	Set footprint = PCBLib.CurrentComponent

	x = replace(x,".",decChar)
	y = replace(y,".",decChar)
	length = replace(length,".",decChar)
	width = replace(width,".",decChar)
	holesize = replace(holesize,".",decChar)
	rotation = replace(rotation,".",decChar)

	apad = pcbserver.PCBObjectFactory(ePadObject,eNoDimension,eCreate_Default)
	apad.Mode = ePadMode_Simple
	apad.name = name
	apad.HoleType = eRoundHole
	apad.HoleSize = mmstocoord(holesize)
	apad.Rotation = rotation
	apad.Plated = CInt(plated)
	apad.TopShape = eRounded
	apad.TopXSize = mmstocoord(length)
	apad.TopYSize = mmstocoord(width)
	apad.layer = eToplayer
	apad.x = mmstocoord(x)+footprint.x
	apad.y = mmstocoord(y)+footprint.y
	if footprint Is Nothing Then Exit Sub
	footprint.board.addpcbobject(apad)
	'PCBServer.SendMessageToRobots footprint.I_ObjectAddress,c_Broadcast,PCBM_BoardRegisteration,apad.I_ObjectAddress
	PCBServer.PostProcess
End Sub

Sub AssignSTEPmodel(STEPFileName, RotX, RotY, RotZ, X, Y, Z)

	Dim STEPmodel
	Dim Model
	Dim PCBLib
	Dim footprint
	Dim fso
	Dim temp_fp

	Set fso = CreateObject("Scripting.FileSystemObject")

	If Not fso.FileExists(STEPFileName) Then
		Exit Sub
	End If

	Set PCBLib = PCBServer.GetCurrentPCBLibrary
	If PCBLib Is Nothing Then Exit Sub

	Set footprint = PCBLib.CurrentComponent
	If footprint Is Nothing Then Exit Sub

	Set temp_fp = PCBLib.CreateNewComponent
	PCBLib.CurrentComponent = temp_fp

	Set STEPmodel = PCBServer.PCBObjectFactory(eComponentBodyObject, eNoDimension, eCreate_Default)
	If STEPmodel Is Nothing Then
		PCBLib.CurrentComponent = footprint
		Exit Sub
	End If

	Set Model = STEPmodel.ModelFactory_FromFilename(STEPFileName, False)
	If Model Is Nothing Then
		PCBLib.CurrentComponent = footprint
		Exit Sub
	End If

	X = Replace(X, ".", decChar)
	Y = Replace(Y, ".", decChar)
	Z = Replace(Z, ".", decChar)
	RotX = Replace(RotX, ".", decChar)
	RotY = Replace(RotY, ".", decChar)
	RotZ = Replace(RotZ, ".", decChar)

	Model.SetState RotX, RotY, RotZ, MMsToCoord(Z)
	STEPmodel.Model = Model

	'PCBLib.CurrentComponent = footprint

	footprint.AddPCBObject(STEPmodel)
	STEPmodel.MoveByXY MMsToCoord(X), MMsToCoord(Y)

	' PCBLib.RemoveComponent(temp_fp)

End Sub

Sub CreateComponentInLib(Name,Description,RefDes)

	Set SCHLib = SchServer.GetCurrentSchDocument

	If SCHLib is Nothing Then
		MsgBox "This is not a SCH library document", vbSystemModal, "Altium Library Loader"
		Exit Sub
	End If

	If SCHLib.ObjectID <> eSchLib Then
		MsgBox "Please open schematic library.", vbSystemModal, "Altium Library Loader"
		Exit Sub
	End If

	'Create a library component (a page of the library is created).
	Set SchComponent = SchServer.SchObjectFactory(eSchComponent, eCreate_Default)
	If SchComponent Is Nothing Then Exit Sub

	'SCHLib.AddSchComponent(SchComponent)
	'SchServer.RobotManager.SendMessage nil, c_BroadCast, SCHM_PrimitiveRegistration, SchComponent.I_ObjectAddress

	'Set up parameters for the library component.
	SchComponent.CurrentPartID = 1
	SchComponent.DisplayMode   = 0

	'Define the LibReference and add the component to the library.
	SchComponent.LibReference = Name

	SchComponent.Designator.Text      = RefDes
	SchComponent.ComponentDescription = Description

End Sub

Sub CreateFootprintInLib(Name,Description)
	Dim PCBLib

	'PCBServer.PreProcess;
	Set PCBLib = PCBServer.GetCurrentPCBLibrary
	If PCBLib is Nothing Then
		MsgBox "This is not a PCB library document", vbSystemModal, "Altium Library Loader"
		Exit Sub
	End If
	Set footprint = pcblib.CreateNewComponent
	PcbLib.CurrentComponent = footprint
	footprint.Name = Name
	footprint.Description = Description

	PCBServer.PostProcess
End Sub

Sub AddParameter(name, value)
	Set Param = SchServer.SchObjectFactory(eParameter, eCreate_Default)
	Param.Name = name
	Param.Text = value
	Param.ShowName = false
	Param.IsHidden = true
	SchComponent.AddSchObject(Param)
End Sub

Sub CreatePin(x, y, designator, name, orientation, length, pintype, pinnames)
	Dim apin

	Set apin = SchServer.SchObjectFactory(ePin, eCreate_Default)
	apin.PinLength = MilsToCoord(length)
	apin.Location = Point(MilsToCoord(x), MilsToCoord(y))
	apin.Color = 0 'BLACK
	apin.Orientation = orientation
	apin.Designator = designator 'in single quotes
	apin.Name = name 'in single quotes
	apin.Electrical = pintype
	apin.ShowName = pinnames
	apin.OwnerPartId = SchLib.CurrentSchComponent.CurrentPartID
	apin.OwnerPartDisplayMode = SchLib.CurrentSchComponent.DisplayMode
	SchComponent.AddSchObject(apin)

End Sub

Sub CreateLeftPin(x, y, designator, name, length, pintype, pinnames)
	Dim apin

	Set apin = SchServer.SchObjectFactory(ePin, eCreate_Default)
	apin.PinLength = MilsToCoord(length)
	apin.Location = Point(MilsToCoord(x), MilsToCoord(y))
	apin.Color = 0 'BLACK
	apin.Orientation = eRotate180
	apin.Designator = designator 'in single quotes
	apin.Name = name 'in single quotes
	apin.Electrical = pintype
	apin.ShowName = pinnames
	apin.OwnerPartId = SchLib.CurrentSchComponent.CurrentPartID
	apin.OwnerPartDisplayMode = SchLib.CurrentSchComponent.DisplayMode
	SchComponent.AddSchObject(apin)

End Sub

Sub CreateRightPin(x, y, designator, name, length, pintype, pinnames)
	Dim apin

	Set apin = SchServer.SchObjectFactory(ePin, eCreate_Default)
	apin.PinLength = MilsToCoord(length)
	apin.Location = Point(MilsToCoord(x), MilsToCoord(y))
	apin.Color = 0 'BLACK
	apin.Orientation = eRotate0
	apin.Designator = designator 'in single quotes
	apin.Name = name 'in single quotes
	apin.Electrical = pintype
	apin.ShowName = pinnames
	apin.OwnerPartId = SchLib.CurrentSchComponent.CurrentPartID
	apin.OwnerPartDisplayMode = SchLib.CurrentSchComponent.DisplayMode
	SchComponent.AddSchObject(apin)

End Sub

Sub CreateTopPin(x, y, designator, name, length, pintype, pinnames)
	Dim apin

	Set apin = SchServer.SchObjectFactory(ePin, eCreate_Default)
	apin.PinLength = MilsToCoord(length)
	apin.Location = Point(MilsToCoord(x), MilsToCoord(y))
	apin.Color = 0 'BLACK
	apin.Orientation = eRotate90
	apin.Designator = designator 'in single quotes
	apin.Name = name 'in single quotes
	apin.Electrical = pintype
	apin.ShowName = pinnames
	apin.OwnerPartId = SchLib.CurrentSchComponent.CurrentPartID
	apin.OwnerPartDisplayMode = SchLib.CurrentSchComponent.DisplayMode
	SchComponent.AddSchObject(apin)

End Sub

Sub CreateBottomPin(x, y, designator, name, length, pintype, pinnames)
	Dim apin

	Set apin = SchServer.SchObjectFactory(ePin, eCreate_Default)
	apin.PinLength = MilsToCoord(length)
	apin.Location = Point(MilsToCoord(x), MilsToCoord(y))
	apin.Color = 0 'BLACK
	apin.Orientation = eRotate270
	apin.Designator = designator 'in single quotes
	apin.Name = name 'in single quotes
	apin.Electrical = pintype
	apin.ShowName = pinnames
	apin.OwnerPartId = SchLib.CurrentSchComponent.CurrentPartID
	apin.OwnerPartDisplayMode = SchLib.CurrentSchComponent.DisplayMode
	SchComponent.AddSchObject(apin)

End Sub

Sub DrawLine(x1, y1, x2, y2, width)
	Dim aline

	'Create a line object for the new library component.
	Set aline = SchServer.SchObjectFactory(eLine,eCreate_Default)
	If aline Is Nothing Then Exit Sub

	'Define the line parameters.
	aline.LineWidth = eSmall
	aline.Location = Point(MilsToCoord(x1), MilsToCoord(y1))
	aline.Corner = Point(MilsToCoord(x2), MilsToCoord(y2))
	aline.Color = 0 'Black
	aline.OwnerPartId = SCHLib.CurrentSchComponent.CurrentPartID
	aline.OwnerPartDisplayMode = SCHLib.CurrentSchComponent.DisplayMode
	SchComponent.AddSchObject(aline)
End Sub

Sub DrawCircle(x, y, rad, width)
	Dim acircle

	'Create a circle object for the new library component.
	Set acircle = SchServer.SchObjectFactory(eArc,eCreate_Default)
	If acircle Is Nothing Then Exit Sub

	'Define the circle parameters.
	acircle.LineWidth = eSmall
	acircle.StartAngle = 0
	acircle.EndAngle =360
	acircle.Location = Point(MilsToCoord(x), MilsToCoord(y))
	acircle.Radius = MilsToCoord(rad)
	acircle.Color = 0 'Black
	acircle.OwnerPartId = SCHLib.CurrentSchComponent.CurrentPartID
	acircle.OwnerPartDisplayMode = SCHLib.CurrentSchComponent.DisplayMode
	SchComponent.AddSchObject(acircle)
End Sub

Sub DrawArc(x1, y1, startAngle, endAngle, rad, width)
	Dim aarc

	'Create a arc object for the new library component.
	Set aarc = SchServer.SchObjectFactory(eArc,eCreate_Default)
	If aarc Is Nothing Then Exit Sub

	'Define the arc parameters.
	aarc.LineWidth = eSmall
	aarc.StartAngle = startAngle
	aarc.EndAngle = endAngle
	aarc.Location = Point(MilsToCoord(x1), MilsToCoord(y1))
	aarc.Radius = MilsToCoord(rad)
	aarc.Color = 0 'Black
	aarc.OwnerPartId = SCHLib.CurrentSchComponent.CurrentPartID
	aarc.OwnerPartDisplayMode = SCHLib.CurrentSchComponent.DisplayMode
	SchComponent.AddSchObject(aarc)
End Sub

Sub DrawRectangle(x1, y1, x2, y2, width)
	Dim arectangle

	'Create a line object for the new library component.

	'R := SchServer.SchObjectFactory(eRectangle, eCreate_Default);
	'If R = Nil Then Exit;

	Set arectangle = SchServer.SchObjectFactory(eRectangle,eCreate_Default)
	If arectangle Is Nothing Then Exit Sub

	'Define the line parameters.
	arectangle.LineWidth = eSmall
	arectangle.Location = Point(MilsToCoord(x1), MilsToCoord(y1))
	arectangle.Corner = Point(MilsToCoord(x2), MilsToCoord(y2))
	arectangle.Color = RGB(128,0,0)
	arectangle.AreaColor = RGB(255,255,176)
	arectangle.IsSolid = true
	arectangle.OwnerPartId = SCHLib.CurrentSchComponent.CurrentPartID
	arectangle.OwnerPartDisplayMode = SCHLib.CurrentSchComponent.DisplayMode
	SchComponent.AddSchObject(arectangle)
End Sub

Sub DrawLine2(x1, y1, x2, y2, width)
	Dim aline

	'Create a line object for the new library component.
	Set aline = SchServer.SchObjectFactory(eLine,eCreate_Default)
	If aline Is Nothing Then Exit Sub

	'Define the line parameters.
	aline.LineWidth = eSmall
	aline.Location = Point(MilsToCoord(x1), MilsToCoord(y1))
	aline.Corner = Point(MilsToCoord(x2), MilsToCoord(y2))
	aline.Color = RGB(0,0,255) 'Blue
	aline.OwnerPartId = SCHLib.CurrentSchComponent.CurrentPartID
	aline.OwnerPartDisplayMode = SCHLib.CurrentSchComponent.DisplayMode
	SchComponent.AddSchObject(aline)
End Sub

Sub DrawCircle2(x, y, rad, width)
	Dim acircle

	'Create a circle object for the new library component.
	Set acircle = SchServer.SchObjectFactory(eArc,eCreate_Default)
	If acircle Is Nothing Then Exit Sub

	'Define the circle parameters.
	acircle.LineWidth = eSmall
	acircle.StartAngle = 0
	acircle.EndAngle =360
	acircle.Location = Point(MilsToCoord(x), MilsToCoord(y))
	acircle.Radius = MilsToCoord(rad)
	acircle.Color = RGB(0,0,255) 'Blue
	acircle.OwnerPartId = SCHLib.CurrentSchComponent.CurrentPartID
	acircle.OwnerPartDisplayMode = SCHLib.CurrentSchComponent.DisplayMode
	SchComponent.AddSchObject(acircle)
End Sub

Sub DrawArc2(x1, y1, startAngle, endAngle, rad, width)
	Dim aarc

	'Create a arc object for the new library component.
	Set aarc = SchServer.SchObjectFactory(eArc,eCreate_Default)
	If aarc Is Nothing Then Exit Sub

	'Define the arc parameters.
	aarc.LineWidth = eSmall
	aarc.StartAngle = startAngle
	aarc.EndAngle = endAngle
	aarc.Location = Point(MilsToCoord(x1), MilsToCoord(y1))
	aarc.Radius = MilsToCoord(rad)
	aarc.Color = RGB(0,0,255) 'Blue
	aarc.OwnerPartId = SCHLib.CurrentSchComponent.CurrentPartID
	aarc.OwnerPartDisplayMode = SCHLib.CurrentSchComponent.DisplayMode
	SchComponent.AddSchObject(aarc)
End Sub

Sub FinaliseComponentInLib()

	Set SCHLib = SchServer.GetCurrentSchDocument

	SCHLib.AddSchComponent(SchComponent)

	'Send a system notification that a new component has been added to the library.
	SchServer.RobotManager.SendMessage nil, c_BroadCast, SCHM_PrimitiveRegistration, SchLib.CurrentSchComponent.I_ObjectAddress
	'SCHLib.CurrentSchComponent = SchComponent

	'Refresh library.
	SCHLib.GraphicallyInvalidate

End Sub

Sub AssignFootprint(LibraryPath, ModelName, ModelMapping)

	Dim Model
	Dim ModelType

	ModelType = "PCBLIB"
	Set Model = SchComponent.AddSchImplementation
	Model.ClearAllDatafileLinks
	Model.MapAsString = ModelMapping
	Model.ModelName = ModelName
	Model.ModelType = ModelType

	Model.AddDataFileLink ModelName, LibraryPath, ModelType
	Model.IsCurrent = True
End Sub

Sub btn_SchLibClick(Sender)
	OpenDialog1.InitialDir = ""
	OpenDialog1.Filter = "Schematic Library (*.SchLib)|*.SchLib"
	If OpenDialog1.Execute Then
		txt_SchLib.Text = OpenDialog1.Filename
		UpdateTXT
	End If
End Sub

Sub btn_PcbLibClick(Sender)
	OpenDialog2.InitialDir = ""
	OpenDialog2.Filter = "PCB Library (*.PcbLib)|*.PcbLib"
	If OpenDialog2.Execute Then
		txt_PcbLib.Text = OpenDialog2.Filename
		UpdateTXT
	End If
End Sub

Sub Form1Show(Sender)

	StringGrid1.Cols(0)(0) = "SYM/FP"
	StringGrid1.Cols(1)(0) = "3D"
	StringGrid1.Cols(2)(0) = "Manuf."
	StringGrid1.Cols(3)(0) = "MPN"
	StringGrid1.Cols(4)(0) = "Desc."
	StringGrid1.Cols(5)(0) = "Partner"
	StringGrid1.Cols(6)(0) = "PartID"

	StringGrid1.ColWidths(5) = 0
	StringGrid1.ColWidths(6) = 0
	StringGrid1.ColWidths(7) = 0

	Set WshShell = CreateObject("WScript.Shell")
	Set fso = CreateObject("Scripting.FileSystemObject")

	InstallFolder = WshShell.SpecialFolders("MyDocuments")
	InstalledDir = InstallFolder & "\AltiumLL"

	UsernameTXT = ReadTXT(1)
	PasswordTXT = ReadTXT(2)
	DnldsFldrTXT = ReadTXT(3)
	SchLibTXT = ReadTXT(4)
	PcbLibTXT = ReadTXT(5)
	UseProxyTXT = ReadTXT(6)
	ProxyAddressTXT = ReadTXT(7)
	ProxyPortTXT = ReadTXT(8)
	ShowInstructionTXT = ReadTXT(9)
	AltiumSymbolsTXT = ReadTXT(10)

	txt_Username.Text = UsernameTXT
	txt_Password.Text = PasswordTXT
	If Trim(txt_Username.Text) = vbNullString or Trim(txt_Password.Text) = vbNullString Then
		ShowSettings()
		lbl_Message.Caption = "Please Login with Email and Password before continuing..."
	Else
		ShowECADModels()
	End If
	txt_DnldsFldr.Text = DnldsFldrTXT
	txt_SchLib.Text = SchLibTXT
	txt_PcbLib.Text = PcbLibTXT
	chk_Proxy.Checked = UseProxyTXT
	txt_Address.Text = ProxyAddressTXT
	txt_Port.Text = ProxyPortTXT
	chk_ShowInstruction.Checked = ShowInstructionTXT
	chk_AltiumSymbols.Checked = AltiumSymbolsTXT

	Set WshShell = Nothing
	Set fso = Nothing

End Sub

Sub ProcessCB(filename)
	Dim stepPath
	Dim stepPathAlt
	Dim PCBLibObj
	Dim comp
	Dim compName
	Dim i
	PcbLib = txt_PcbLib.Text

	PcbLibDoc = Client.OpenDocument("PcbLib",PcbLib)
	Client.ShowDocument(PcbLibDoc)

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.OpenTextFile(filename)

	AddPCB = true
	AddSCH = true

	Do Until f.AtEndOfStream
		lineArray = Split(f.ReadLine, ", ")
		'MsgBox "CMD: " & lineArray(0)
		If lineArray(0) = "CreateFootprintInLib" Then
			AddPCB = AddPcbLib(lineArray(1))
			If Not AddPCB Then
				MsgBox "Footprint " & lineArray(1) & " already exits! Skipping Library Load.", vbSystemModal, "Altium Library Loader"
				If stpFileName <> vbNullString Then fso.DeleteFile(InstalledDir & "\Temp\" & stpFileName)
			Else
				CreateFootprintInLib lineArray(1), lineArray(2)
			End If
		ElseIf lineArray(0) = "CreateRectangularSMD" And AddPCB Then
			CreateRectangularSMD lineArray(1), lineArray(2), lineArray(3), lineArray(4), lineArray(5), lineArray(6), lineArray(7), lineArray(8)
		ElseIf lineArray(0) = "CreateRoundedRectangularSMD" And AddPCB Then
			CreateRoundedRectangularSMD lineArray(1), lineArray(2), lineArray(3), lineArray(4), lineArray(5), lineArray(6), lineArray(7), lineArray(8)
		ElseIf lineArray(0) = "CreateRoundedSMD" And AddPCB Then
			CreateRoundedSMD lineArray(1), lineArray(2), lineArray(3), lineArray(4), lineArray(5), lineArray(6), lineArray(7), lineArray(8)
		ElseIf lineArray(0) = "CreateRoundedPTH" And AddPCB Then
			CreateRoundedPTH lineArray(1), lineArray(2), lineArray(3), lineArray(4), lineArray(5), lineArray(6), lineArray(7), lineArray(8)
		ElseIf lineArray(0) = "CreateRectangularPTH" And AddPCB Then
			CreateRectangularPTH lineArray(1), lineArray(2), lineArray(3), lineArray(4), lineArray(5), lineArray(6), lineArray(7), lineArray(8)
		ElseIf lineArray(0) = "CreateCourtyardLine" And AddPCB Then
			CreateCourtyardLine lineArray(1), lineArray(2), lineArray(3), lineArray(4), lineArray(5)
		ElseIf lineArray(0) = "CreateAssemblyLine" And AddPCB Then
			CreateAssemblyLine lineArray(1), lineArray(2), lineArray(3), lineArray(4), lineArray(5)
		ElseIf lineArray(0) = "CreateSilkscreenLine" And AddPCB Then
			CreateSilkscreenLine lineArray(1), lineArray(2), lineArray(3), lineArray(4), lineArray(5)
		ElseIf lineArray(0) = "CreateCourtyardCircle" And AddPCB Then
			CreateCourtyardCircle lineArray(1), lineArray(2), lineArray(3), lineArray(4)
		ElseIf lineArray(0) = "CreateAssemblyCircle" And AddPCB Then
			CreateAssemblyCircle lineArray(1), lineArray(2), lineArray(3), lineArray(4)
		ElseIf lineArray(0) = "CreateSilkscreenCircle" And AddPCB Then
			CreateSilkscreenCircle lineArray(1), lineArray(2), lineArray(3), lineArray(4)
		ElseIf lineArray(0) = "CreateAssemblyArc" And AddPCB Then
			CreateAssemblyArc lineArray(1), lineArray(2), lineArray(3), lineArray(4), lineArray(5), lineArray(6)
		ElseIf lineArray(0) = "CreateSilkscreenArc" And AddPCB Then
			CreateSilkscreenArc lineArray(1), lineArray(2), lineArray(3), lineArray(4), lineArray(5), lineArray(6)
		ElseIf lineArray(0) = "AssignSTEPmodel" And AddPCB Then
			stepPath = InstalledDir & "\Temp\" & lineArray(1)

			If Not fso.FileExists(stepPath) Then
				If LCase(Right(stepPath, 4)) = ".stp" Then
					stepPathAlt = Left(stepPath, Len(stepPath) - 4) & ".step"
				ElseIf LCase(Right(stepPath, 5)) = ".step" Then
					stepPathAlt = Left(stepPath, Len(stepPath) - 5) & ".stp"
				Else
					stepPathAlt = stepPath & ".stp"
				End If

				If fso.FileExists(stepPathAlt) Then
					stepPath = stepPathAlt
				End If
			End If

			If fso.FileExists(stepPath) Then
				AssignSTEPmodel stepPath, lineArray(2), lineArray(3), lineArray(4), lineArray(5), lineArray(6), lineArray(7)
				fso.DeleteFile(stepPath)
			Else
				' STEP is optional; continue without 3D model
			End If

		ElseIf lineArray(0) = "CreateComponentInLib" Then
			PCBServer.PostProcess
			PcbLibDoc.DoFileSave("PcbLib")

			Client.CloseDocument(PcbLibDoc)
			Set PcbLibDoc = Client.OpenDocument("PcbLib", PcbLib)
			Client.ShowDocument(PcbLibDoc)

			Set PCBLibObj = PCBServer.GetCurrentPCBLibrary

			If Not PCBLibObj Is Nothing Then
				For i = PCBLibObj.ComponentCount - 1 To 0 Step -1
					Set comp = PCBLibObj.GetComponent(i)

					If Not comp Is Nothing Then
						compName = comp.Name

						If Left(compName, 13) = "PCBComponent_" Then
							PCBLibObj.RemoveComponent(comp)
						End If
					End If
				Next
			End If

			SchLib = txt_SchLib.Text
			SchLibDoc = Client.OpenDocument("SchLib", SchLib)
			Client.ShowDocument(SchLibDoc)
			prtName = lineArray(1)
			AddSCH = AddSchLib(prtName)
			Set SCHLib = SchServer.GetCurrentSchDocument
			Set SchComponent = SchServer.SchObjectFactory(eSchComponent, eCreate_Default)

			If AddSCH Then
				If SCHLib Is Nothing Then
					MsgBox "This is not a SCH library document", vbSystemModal, "Altium Library Loader"
					Exit Sub
				End If

				If SCHLib.ObjectID <> eSchLib Then
					MsgBox "Please open schematic library.", vbSystemModal, "Altium Library Loader"
					Exit Sub
				End If

				If SchComponent Is Nothing Then Exit Sub

				SchComponent.CurrentPartID = 1
				SchComponent.DisplayMode = 0
				SchComponent.LibReference = lineArray(1)
				SchComponent.Designator.Text = lineArray(3)
				SchComponent.ComponentDescription = lineArray(2)
			End If
		ElseIf lineArray(0) = "CreateLeftPin" And AddSCH Then
			CreateLeftPin lineArray(1), lineArray(2), lineArray(3), lineArray(4), lineArray(5), CInt(lineArray(6)), CInt(lineArray(7))
		ElseIf lineArray(0) = "CreateRightPin" And AddSCH Then
			CreateRightPin lineArray(1), lineArray(2), lineArray(3), lineArray(4), lineArray(5), CInt(lineArray(6)), CInt(lineArray(7))
		ElseIf lineArray(0) = "CreateTopPin" And AddSCH Then
			CreateTopPin lineArray(1), lineArray(2), lineArray(3), lineArray(4), lineArray(5), CInt(lineArray(6)), CInt(lineArray(7))
		ElseIf lineArray(0) = "CreateBottomPin" And AddSCH Then
			CreateBottomPin lineArray(1), lineArray(2), lineArray(3), lineArray(4), lineArray(5), CInt(lineArray(6)), CInt(lineArray(7))
		ElseIf lineArray(0) = "DrawLine" And AddSCH Then
			DrawLine lineArray(1), lineArray(2), lineArray(3), lineArray(4), lineArray(5)
		ElseIf lineArray(0) = "DrawCircle" And AddSCH Then
			DrawCircle lineArray(1), lineArray(2), lineArray(3), lineArray(4)
		ElseIf lineArray(0) = "DrawArc" And AddSCH Then
			DrawArc lineArray(1), lineArray(2), lineArray(3), lineArray(4), lineArray(5), lineArray(6)
		ElseIf lineArray(0) = "DrawRectangle" And AddSCH Then
			DrawRectangle lineArray(1), lineArray(2), lineArray(3), lineArray(4), lineArray(5)
		ElseIf lineArray(0) = "DrawLine2" And AddSCH Then
			DrawLine2 lineArray(1), lineArray(2), lineArray(3), lineArray(4), lineArray(5)
		ElseIf lineArray(0) = "DrawCircle2" And AddSCH Then
			DrawCircle2 lineArray(1), lineArray(2), lineArray(3), lineArray(4)
		ElseIf lineArray(0) = "DrawArc2" And AddSCH Then
			DrawArc2 lineArray(1), lineArray(2), lineArray(3), lineArray(4), lineArray(5), lineArray(6)
		ElseIf lineArray(0) = "AddParameter" And AddSCH Then
			AddParameter lineArray(1), lineArray(2)
		ElseIf lineArray(0) = "AssignFootprint" And AddSCH Then
			AssignFootprint PcbLib, lineArray(1), lineArray(2)
		End If
	Loop
	f.Close

	'If AddSCH Then FinaliseComponentInLib()
	'FinaliseComponentInLib()

	'FinaliseComponentInLib
	'Set SCHLib = SchServer.GetCurrentSchDocument
	If AddSCH Then
		SCHLib.AddSchComponent(SchComponent)
		'Send a system notification that a new component has been added to the library.
		SchServer.RobotManager.SendMessage nil, c_BroadCast, SCHM_PrimitiveRegistration, SchComponent.I_ObjectAddress
		'SCHLib.CurrentSchComponent = SchComponent
		'Refresh library.
		SCHLib.GraphicallyInvalidate
		SchLibDoc.DoFileSave("SchLib")
	End If

	fso.DeleteFile(filename)
	Set fso = Nothing

	CurrentSCH = Client.OpenDocument("SCH", SchDocFile)
	Client.ShowDocument(CurrentSCH)

	If AddSCH Then
		AddToSch
	Else
		LastSlashIndex = InStrRev(txt_SchLib.Text,"\")
		LibName = Mid(txt_SchLib.Text,LastSlashIndex+1)
		MsgBox "Component " & prtName & " already exits!" & vbCrLf & vbCrLf & "Please place from " & LibName, vbSystemModal, "Altium Library Loader"
	End if
End Sub

Sub ExtractEPW(fldr, dst)
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set sa = CreateObject("Shell.Application")
	For Each f In fldr.Items
		If f.Type = "File folder" Then
			ExtractEPW f.GetFolder, dst
		ElseIf f.Type = "EPW File" Or f.Type = "File EPW" Then
			If fso.GetExtensionName(f.Name) = vbNullString Then
				cbFilename = f.Name & ".epw"
			Else
				cbFilename = f.Name
			End If
			sa.NameSpace(dst).CopyHere f.Path
		End If
	Next
	Set fso = Nothing
	Set sa = Nothing
End Sub

Sub zCleanUp(file, count)
	'Clean up
	Dim i, fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	For i = 1 To count
		If fso.FolderExists(fso.GetSpecialFolder(2) & "\Temporary Directory " & i & " for " & file) = True Then
			text = fso.DeleteFolder(fso.GetSpecialFolder(2) & "\Temporary Directory " & i & " for " & file, True)
		Else
			Exit For
		End If
	Next
	Set fso = Nothing
End Sub

Sub btn_DnldsFldrClick(Sender)
	Dim DownloadsFolder
	Dim objShell
	Dim objFolder

	DownloadsFolder = GetDownloadsFolder()

	Set objShell = CreateObject("Shell.Application")
	Set objFolder = objShell.BrowseForFolder(0, "Select a folder:", 0, DownloadsFolder)

	If Not objFolder Is Nothing Then
		txt_DnldsFldr.Text = objFolder.Items.Item.Path
		UpdateTXT
	End If

	Set objFolder = Nothing
	Set objShell = Nothing
End Sub

Sub Prechecks()
	EnsureSettingsFile()
	If IsLibraryLoaderRunning Then
		MsgBox "Please uninstall Library Loader before running Altium Library Loader.", vbSystemModal, "Altium Library Loader"
	Else
		Form1.Show()
	End If
End Sub

Sub AddToSch()
	Dim ObjectHandle
	Set ObjectHandle = SchServer.LoadComponentFromLibrary(prtName, txt_SchLib.Text)
	Set CurrentSheet = SchServer.GetCurrentSchDocument
	CurrentSheet.AddSchObject(ObjectHandle)
	ObjectHandle.MoveToXY 0, 0
	ObjectHandle.SetState_Orientation 0
	CurrentSheet.GraphicallyInvalidate
	Set ObjectHandle = Nothing
	If chk_ShowInstruction.Checked Then MsgBox "Please check the bottom left corner of your schematic to drag component " & prtName & " into the desired position.", vbSystemModal, "Altium Library Loader"
End Sub

Function PlaceComponent(TargetSheet, LibraryLocation, LibraryRef, X, Y, Orientation)
	Dim ObjectHandle
	Set ObjectHandle = SchServer.LoadComponentFromLibrary(LibraryRef, LibraryLocation)

	'TargetSheet.AddAndPositionSchObject(ObjectHandle)
	TargetSheet.AddSchObject(ObjectHandle)
	ObjectHandle.MoveToXY X, Y
	ObjectHandle.SetState_Orientation Orientation
	TargetSheet.GraphicallyInvalidate
	Set ObjectHandle = Nothing
End Function

Function ProcessSelectedPart(partID)

	If CheckLogin Then
		stpFileName = vbNullString
		selectedRow = StringGrid1.Row
		If StringGrid1.Cols(0)(selectedRow) = "Y" Or partID <> vbNullString Then
			Client.StartServer("SCH")
			Set CurrentSheet = SchServer.GetCurrentSchDocument
			If CurrentSheet Is Nothing Then
				MsgBox "A schematic document (*.SchDoc) must be opened or selected before ECAD Model can be added to design.", vbSystemModal, "Altium Library Loader"
				ProcessSelectedPart = False
				Exit Function
			ElseIf UCase(Right(CurrentSheet.GetState_DocumentName,7)) <> ".SCHDOC" Then
				MsgBox "A schematic document (*.SchDoc) must be opened or selected before ECAD Model can be added to design.", vbSystemModal, "Altium Library Loader"
				ProcessSelectedPart = False
				Exit Function
			End If
			SchDocFile = CurrentSheet.GetState_DocumentName
		End If
		Screen.Cursor = crHourglass
		ecadModelPath = InstalledDir & "\Temp\"

		If partID = vbNullString Then
			selectedRow = StringGrid1.Row
			epwMAN = StringGrid1.Cols(2)(selectedRow)
			epwMPN = StringGrid1.Cols(3)(selectedRow)
			epwPNA = StringGrid1.Cols(5)(selectedRow)
			If epwPNA <> "SamacSys" Then epwWSP = epwPNA
			partID = StringGrid1.Cols(6)(selectedRow)
		End If

		username = txt_Username.Text
		password = txt_Password.Text

		If StringGrid1.Cols(0)(selectedRow) = "N" And StringGrid1.Cols(1)(selectedRow) = "Y" Then '3D Only

			prtName = strClean(StringGrid1.Cols(3)(selectedRow))

			Set oReq = CreateHttpRequest()
			If epwWSP = vbNullString Then
				oReq.Open "GET", "https://ad.componentsearchengine.com/ga/model.php?partID=" & partID & "&pi=2&step=1&st=4&lt=4", False
			Else
				oReq.Open "GET", "https://" & epwWSP & ".componentsearchengine.com/ga/model.php?partID=" & partID & "&pi=2&step=1&st=4&lt=4", False
			End If
			oReq.setRequestHeader "Authorization", "Basic " + Replace(Mid(Base64Encode(username + ":" + password),5),vbLf,"")
			oReq.setRequestHeader "User-Agent", "AltiumLibraryLoaderV" & AppVersion
			If chk_Proxy.Checked Then
				oReq.setProxy 2, Trim(txt_Address.Text) & ":" & Trim(txt_Port.Text), ""
			End if
			Err.Clear
			oReq.send

			If Err.Number <> 0 Then
				MsgBox "HTTP send failed: " & Err.Description, vbSystemModal, "Altium Library Loader"
				Err.Clear
				Screen.Cursor = crDefault
				ProcessSelectedPart = False
				Exit Function
			End If
			Set bStrm = createobject("Adodb.Stream")
			with bStrm
				.Type = 1 '//binary
				.open
				.write oReq.responseBody
				stpFileName = prtName & ".stp"
				.savetofile ecadModelPath & stpFileName, 2 '//overwrite
				.Close
			end With
			Set oReq = Nothing
			vbAns = MsgBox("3D Model STEP file has been saved to " & ecadModelPath & stpFileName & vbCrLf & vbCrLf & "Would you like to build or request the Symbol and Footprint?", vbYesNo, "Altium Library Loader")
			If vbAns = vbYes Then
				CreateObject("WScript.Shell").Run("https://ad.componentsearchengine.com/partRequest.html?partID=" & partID)
			End If
			Screen.Cursor = crDefault
			lbl_Message.Caption = vbNullString
			ProcessSelectedPart = True
			Exit Function
		End If

		If epwWSP = vbNullString Then
			If chk_AltiumSymbols.Checked Then
				cb = httpGET("https://ad.componentsearchengine.com/ga/model.php?partID=" & partID & "&st=4&lt=4&pi=2&ver=2", username, password)
			Else
				cb = httpGET("https://ad.componentsearchengine.com/ga/model.php?partID=" & partID & "&st=4&lt=4&pi=2", username, password)
			End If
		Else
			If chk_AltiumSymbols.Checked Then
				cb = httpGET("https://" & epwWSP & ".componentsearchengine.com/ga/model.php?partID=" & partID & "&st=4&lt=4&pi=2&ver=2", username, password)
			Else
				cb = httpGET("https://" & epwWSP & ".componentsearchengine.com/ga/model.php?partID=" & partID & "&st=4&lt=4&pi=2", username, password)
			End If
		End If
		If Left(cb, 6) = "Error:" Then
			MsgBox cb, vbSystemModal, "Altium Library Loader"
			Screen.Cursor = crDefault
			ProcessSelectedPart = False
			Exit Function
		End If
		If Trim(cb) = "Error: This part ID is not released" Or cb = "Error: Data entry for this part ID is not yet complete." Then
			MsgBox "This part is almost ready. You have 3 options:" & vbCrLf & vbCrLf & "1. Find Alternate" & vbCrLf & "2. Online Build" & vbCrLf & "3. Request we build it for FREE!", vbSystemModal, "Find Alternate, Build or Request"
			If epwWSP = vbNullString Then
				CreateObject("WScript.Shell").Run("https://ad.componentsearchengine.com/entry_u.php?mna=" & URLEncode(epwMAN) & "&mpn=" & URLEncode(epwMPN) & "&pna=" & URLEncode(epwPNA))
			Else
				CreateObject("WScript.Shell").Run("https://" & epwWSP & ".componentsearchengine.com/entry_u.php?mna=" & URLEncode(epwMAN) & "&mpn=" & URLEncode(epwMPN) & "&pna=" & URLEncode(epwPNA))
			End If
			ProcessSelectedPart = True
			Screen.Cursor = crDefault
			lbl_Message.Caption = vbNullString
			Exit Function
		ElseIf InStr(cb, "Error") Then
			MsgBox cb, vbSystemModal, "Altium Library Loader"
			ProcessSelectedPart = False
			Screen.Cursor = crDefault
			lbl_Message.Caption = vbNullString
			Exit Function
		End If

		cbArray = Split(cb, vbCrLf)
		prtName = cbArray(0)

		If InStr(cb,"AssignSTEPmodel") <> 0 Then
			Set oReq = CreateHttpRequest()
			If epwWSP = vbNullString Then
				oReq.Open "GET", "https://ad.componentsearchengine.com/ga/model.php?partID=" & partID & "&pi=2&step=1&st=4&lt=4", False
			Else
				oReq.Open "GET", "https://" & epwWSP & ".componentsearchengine.com/ga/model.php?partID=" & partID & "&pi=2&step=1&st=4&lt=4", False
			End If
			oReq.setRequestHeader "Authorization", "Basic " + Replace(Mid(Base64Encode(username + ":" + password),5),vbLf,"")
			oReq.setRequestHeader "User-Agent", "AltiumLibraryLoaderV" & AppVersion
			If chk_Proxy.Checked Then
				oReq.setProxy 2, Trim(txt_Address.Text) & ":" & Trim(txt_Port.Text), ""
			End if

			Err.Clear
			oReq.send

			If Err.Number <> 0 Then
				MsgBox "HTTP send failed: " & Err.Description, vbSystemModal, "Altium Library Loader"
				Err.Clear
				Screen.Cursor = crDefault
				ProcessSelectedPart = False
				Exit Function
			End If
			Set bStrm = createobject("Adodb.Stream")
			with bStrm
				.Type = 1 '//binary
				.open
				.write oReq.responseBody
				stpFileName = prtName & ".stp"
				.savetofile ecadModelPath & stpFileName, 2 '//overwrite
				.Close
			end With
			Set oReq = Nothing
		End If

		outFile= ecadModelPath & prtName & ".cb"
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set cbFile = fso.CreateTextFile(outFile,True)
		cbFile.Write cb
		cbFile.Close
		Set fso = Nothing
		ProcessCB(ecadModelPath & prtName & ".cb")

		Screen.Cursor = crDefault
		lbl_Message.Caption = vbNullString
		ProcessSelectedPart = True
	Else
		ShowSettings()
		lbl_Message.Caption = "Please Login with Email and Password before continuing..."
		ProcessSelectedPart = False
	End If

End Function

Function Base64Encode(sText)

	Dim oXML, oNode
	Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
	Set oNode = oXML.CreateElement("base64")
	oNode.dataType = "bin.base64"
	oNode.nodeTypedValue =Stream_StringToBinary(sText)
	Base64Encode = oNode.text
	Set oNode = Nothing
	Set oXML = Nothing
End Function

Function Stream_StringToBinary(Text)
	Const adTypeText = 2
	Const adTypeBinary = 1

	'Create Stream object
	Dim BinaryStream 'As New Stream
	Set BinaryStream = CreateObject("ADODB.Stream")

	'Specify stream type - we want To save text/string data.
	BinaryStream.Type = adTypeText

	'Specify charset For the source text (unicode) data.
	BinaryStream.CharSet = "utf-8"

	'Open the stream And write text/string data To the object
	BinaryStream.Open
	BinaryStream.WriteText Text

	'Change stream type To binary
	BinaryStream.Position = 0
	BinaryStream.Type = adTypeBinary

	'Ignore first two bytes - sign of
	BinaryStream.Position = 0

	'Open the stream And get binary data from the object
	Stream_StringToBinary = BinaryStream.Read

	Set BinaryStream = Nothing
End Function

'Stream_BinaryToString Function
'Binary - VT_UI1 | VT_ARRAY data To convert To a string

Function Stream_BinaryToString(Binary)
	Const adTypeText = 2
	Const adTypeBinary = 1

	'Create Stream object
	Dim BinaryStream 'As New Stream
	Set BinaryStream = CreateObject("ADODB.Stream")

	'Specify stream type - we want To save text/string data.
	BinaryStream.Type = adTypeBinary

	'Open the stream And write text/string data To the object
	BinaryStream.Open
	BinaryStream.Write Binary

	'Change stream type To binary
	BinaryStream.Position = 0
	BinaryStream.Type = adTypeText

	'Specify charset For the source text (unicode) data.
	BinaryStream.CharSet = "utf-8"

	'Open the stream And get binary data from the object
	Stream_BinaryToString = BinaryStream.ReadText
	Set BinaryStream = Nothing
End Function

Function httpGET(url, username, password)
	Dim oReq
	Dim auth

	httpGET = vbNullString
	On Error Resume Next

	auth = "Basic " + Replace(Mid(Base64Encode(username + ":" + password), 5), vbLf, "")

	Set oReq = CreateHttpRequest()
	If oReq Is Nothing Then
		httpGET = "Error: failed to create HTTP object"
		Exit Function
	End If

	Err.Clear
	oReq.Open "GET", url, False
	oReq.setRequestHeader "Authorization", auth
	oReq.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
	oReq.setRequestHeader "User-Agent", "AltiumLibraryLoaderV" & AppVersion

	If chk_Proxy.Checked Then
		oReq.setProxy 2, Trim(txt_Address.Text) & ":" & Trim(txt_Port.Text), ""
	End If

	oReq.Send

	If Err.Number <> 0 Then
		httpGET = "Error: HTTP send failed: " & Err.Description & vbCrLf & url
		Err.Clear
		Set oReq = Nothing
		Exit Function
	End If

	If oReq.Status < 200 Or oReq.Status >= 300 Then
		httpGET = "Error: HTTP status " & CStr(oReq.Status) & " " & oReq.statusText & vbCrLf & url
		Set oReq = Nothing
		Exit Function
	End If

	httpGET = oReq.responseText
	Set oReq = Nothing
	On Error GoTo 0
End Function

Function CreateHttpRequest()
	On Error Resume Next

	Set CreateHttpRequest = CreateObject("Msxml2.ServerXMLHTTP.6.0")

	If Err.Number <> 0 Then
		Err.Clear
		Set CreateHttpRequest = CreateObject("Msxml2.ServerXMLHTTP")
	End If

	On Error GoTo 0
End Function

Public Function URLEncode(StringToEncode)
	Dim TempAns
	Dim CurChr
	CurChr = 1
	Do Until CurChr - 1 = Len(StringToEncode)
		Select Case Asc(Mid(StringToEncode, CurChr, 1))
		Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122
			TempAns = TempAns & Mid(StringToEncode, CurChr, 1)
		Case 32
			TempAns = TempAns & "%" & Hex(32)
		Case Else
			TempAns = TempAns & "%" & Right("00" & Hex(Asc(Mid(StringToEncode, CurChr, 1))), 2)
		End Select
		CurChr = CurChr + 1
	Loop

	URLEncode = TempAns
End Function

Function ReadTXT(LineNo)
	Dim TXTFileName
	Dim fso
	Dim TXTFile
	Dim txtData
	Dim arrData

	ReadTXT = vbNullString

	TXTFileName = InstalledDir & "\AltiumLL.txt"
	Set fso = CreateObject("Scripting.FileSystemObject")

	If Not fso.FileExists(TXTFileName) Then
		Exit Function
	End If

	Set TXTFile = fso.OpenTextFile(TXTFileName, 1)
	txtData = TXTFile.ReadAll
	TXTFile.Close

	arrData = Split(txtData, vbCrLf)

	If LineNo > 0 And LineNo <= UBound(arrData) + 1 Then
		ReadTXT = arrData(LineNo - 1)
	End If
End Function

Function UpdateTXT()
	TXTFileName = InstalledDir & "\AltiumLL.txt"
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.createTextfile(TXTFileName, True)
	With f
		.Writeline txt_Username.Text
		.Writeline txt_Password.Text
		.Writeline txt_DnldsFldr.Text
		.Writeline txt_SchLib.Text
		.Writeline txt_PcbLib.Text
		.Writeline chk_Proxy.Checked
		.Writeline txt_Address.Text
		.Writeline txt_Port.Text
		.Writeline chk_ShowInstruction.Checked
		.Writeline chk_AltiumSymbols.Checked
		.Close
	End With
	Set fso = Nothing
End Function

Sub btn_LoginClick(Sender)
	Login()
End Sub

Sub Login()
	responseStr = httpGET("https://ad.componentsearchengine.com/ga/auth.txt?",txt_Username.Text,txt_Password.Text)
	If responseStr = "OK" Then
		MsgBox "Login successful", vbSystemModal, "Altium Library Loader"
		UpdateTXT
		lbl_Message.Caption = vbNullString
		ShowECADModels()
	Else
		'Msgbox responseStr, vbSystemModal, "Altium Library Loader"
		Msgbox "Login failed: Please check your user name and password and try again.", vbSystemModal, "Altium Library Loader"
		lbl_Message.Caption = "Please Login with Email and Password before continuing..."
	End If
End Sub

Sub lbl_RegisterClick(Sender)
	CreateObject("WScript.Shell").Run("https://ad.componentsearchengine.com/register.php")
End Sub

Sub lbl_ForgotPasswordClick(Sender)
	CreateObject("WScript.Shell").Run("https://ad.componentsearchengine.com/resetPassword.php")
End Sub

Function AddPcbLib(footprint)
	PCBLib = PCBServer.GetCurrentPCBLibrary
	If PCBLib is Nothing Then
		MsgBox "This is not a PCB library document", vbSystemModal, "Altium Library Loader"
		Exit Function
	End If
	For j = 0 to PCBLib.ComponentCount - 1
		Component = PCBLib.GetComponent(j)
		If Component.Name = footprint Then
			AddPcbLib = false
			Exit Function
		End If
	Next
	AddPcbLib = true
End Function

Function AddSchLib(component)
	Set CurrentLib = SCHServer.GetCurrentSchDocument
	Set LibraryIterator = CurrentLib.SchLibIterator_Create
	LibraryIterator.AddFilter_ObjectSet(MkSet(eSchComponent))
	Set LibComp = LibraryIterator.FirstSchObject
	LibCompNameNext = LibComp.LibReference
	On Error Resume Next
	Do
		LibCompNamePrev = LibCompNameNext
		If LibCompNameNext = component Then
			AddSchLib = false
			Exit Function
		End If
		Set LibComp = LibraryIterator.NextSchObject
		LibCompNameNext = LibComp.LibReference
	Loop Until LibCompNameNext = LibCompNamePrev
	CurrentLib.SchIterator_Destroy(LibraryIterator)
	AddSchLib = true
End Function

Sub chk_ShowInstructionClick(Sender)
	UpdateTXT
End Sub

Function CheckLogin()
	If Trim(txt_Username.Text) = vbNullString Or Trim(txt_Password.Text) = vbNullString Then
		CheckLogin = False
	Else
		responseStr = httpGET("https://ad.componentsearchengine.com/ga/auth.txt?",txt_Username.Text,txt_Password.Text)
		If responseStr = "OK" Then
			CheckLogin = True
		Else
			Msgbox "Login failed: Please check your user name and password and try again.", vbSystemModal, "Altium Library Loader"
			CheckLogin = False
		End If
	End If
End Function

Function IsLibraryLoaderRunning()
	sComputerName = "."
	Set objWMIService = GetObject("winmgmts:\\" & sComputerName & "\root\cimv2")
	sQuery = "SELECT * FROM Win32_Process"
	Set objItems = objWMIService.ExecQuery(sQuery)

	For Each objItem In objItems
		If objItem.Name = "Library Loader.exe" Then
			IsLibraryLoaderRunning = true
			Exit Function
		End if
	Next
	IsLibraryLoaderRunning = false
End Function

Sub img_SettingsClick(Sender)
	ShowSettings()
End Sub

Sub img_SettingsDisabledClick(Sender)
	ShowSettings()
End Sub

Sub HideSharePage()
	img_Facebook.Visible = False
	img_Twitter.Visible = False
	img_LinkedIn.Visible = False
	img_Email.Visible = False
End Sub

Sub ShowSharePage()
	img_Facebook.Visible = True
	img_Twitter.Visible = True
	img_LinkedIn.Visible = True
	img_Email.Visible = True
End Sub

Sub ShowSettingsPage()
	grp_Login.Visible = True
	grp_Settings.Visible = True
	lbl_Username.Visible = True
	txt_Username.Visible = True
	btn_Login.Visible = True
	lbl_Register.Visible = True
	lbl_Password.Visible = True
	txt_Password.Visible = True
	lbl_ForgotPassword.Visible = True

	lbl_DnldsFldr.Visible = True
	txt_DnldsFldr.Visible = True
	btn_DnldsFldr.Visible = True
	lbl_SchLib.Visible = True
	txt_SchLib.Visible = True
	btn_SchLib.Visible = True
	lbl_PcbLib.Visible = True
	txt_PcbLib.Visible = True
	btn_PcbLib.Visible = True
	chk_ShowInstruction.Visible = True

	txt_Username.SetFocus

End Sub

Sub HideSettingsPage()
	grp_Login.Visible = False
	grp_Settings.Visible = False

	lbl_Username.Visible = False
	txt_Username.Visible = False
	btn_Login.Visible = False
	lbl_Register.Visible = False
	lbl_Password.Visible = False
	txt_Password.Visible = False
	lbl_ForgotPassword.Visible = False
	lbl_DnldsFldr.Visible = False
	txt_DnldsFldr.Visible = False
	btn_DnldsFldr.Visible = False
	lbl_SchLib.Visible = False
	txt_SchLib.Visible = False
	btn_SchLib.Visible = False
	lbl_PcbLib.Visible = False
	txt_PcbLib.Visible = False
	btn_PcbLib.Visible = False
	chk_ShowInstruction.Visible = False

	UpdateTXT

End Sub

Sub ShowSettings()

	img_ECADModelsDisabled.Visible = True
	img_ECADModels.Visible = False
	img_SettingsDisabled.Visible = False
	img_Settings.Visible = True
	img_ShareDisabled.Visible = True
	img_Share.Visible = False
	img_HelpDisabled.Visible = True
	img_Help.Visible = False

	HideECADModelsPage()
	HideSharePage()
	HideHelpPage()

	ShowSettingsPage()

End Sub

Sub ShowECADModelsPage()
	pnl_OpenModel.Visible = True
	txt_Keywords.Visible = True
	pnl_Search.Visible = True
	lbl_SearchResults.Visible = True
	StringGrid1.Visible = True
	pnl_MoreInfo.Visible = True
	pnl_Design.Visible = True
	pnl_Images.Visible = True
	txt_Keywords.setfocus
End Sub

Sub HideECADModelsPage()

	pnl_OpenModel.Visible = False
	txt_Keywords.Visible = False
	pnl_Search.Visible = False
	lbl_SearchResults.Visible = False
	StringGrid1.Visible = False
	pnl_MoreInfo.Visible = False
	pnl_Design.Visible = False
	pnl_Images.Visible = False
End Sub

Sub ShowHelpPage()
	img_GetStarted.Visible = True
	img_Instruction.Visible = True
	img_ContactUs.Visible = True
	img_Review.Visible = True
	img_Feedback.Visible = True
	img_AboutUs.Visible = True
End Sub

Sub HideHelpPage()
	img_GetStarted.Visible = False
	img_Instruction.Visible = False
	img_ContactUs.Visible = False
	img_Review.Visible = False
	img_Feedback.Visible = False
	img_AboutUs.Visible = False
End Sub

Sub ShowECADModels()

	img_ECADModelsDisabled.Visible = False
	img_ECADModels.Visible = True
	img_SettingsDisabled.Visible = True
	img_Settings.Visible = False
	img_ShareDisabled.Visible = True
	img_Share.Visible = False
	img_HelpDisabled.Visible = True
	img_Help.Visible = False

	HideSettingsPage()
	HideSharePage()
	HideHelpPage()

	ShowECADModelsPage()

End Sub

Sub ShowShare()

	img_ECADModelsDisabled.Visible = True
	img_ECADModels.Visible = False
	img_SettingsDisabled.Visible = True
	img_Settings.Visible = False
	img_ShareDisabled.Visible = False
	img_Share.Visible = True
	img_HelpDisabled.Visible = True
	img_Help.Visible = False

	HideECADModelsPage()
	HideSettingsPage
	HideHelpPage()

	ShowSharePage()

End Sub

Sub ShowHelp()

	img_ECADModelsDisabled.Visible = True
	img_ECADModels.Visible = False
	img_SettingsDisabled.Visible = True
	img_Settings.Visible = False
	img_ShareDisabled.Visible = True
	img_Share.Visible = False
	img_HelpDisabled.Visible = False
	img_Help.Visible = True

	HideECADModelsPage()
	HideSettingsPage()
	HideSharePage()
	ShowHelpPage()

End Sub

Sub ShowBlank()

	img_ECADModelsDisabled.Visible = True
	img_ECADModels.Visible = False
	img_SettingsDisabled.Visible = True
	img_Settings.Visible = False
	img_ShareDisabled.Visible = True
	img_Share.Visible = False
	img_HelpDisabled.Visible = True
	img_Help.Visible = False

	HideECADModelsPage()
	HideSettingsPage()
	HideSharePage()
	HideHelpPage()
	ShowBlankPage()

End Sub

Sub img_ECADModelsClick(Sender)
	ShowECADModels()
End Sub

Sub img_ECADModelsDisabledClick(Sender)
	ShowECADModels()
End Sub

Sub img_ShareClick(Sender)
	ShowShare()
End Sub

Sub img_ShareDisabledClick(Sender)
	ShowShare()
End Sub

Sub img_HelpClick(Sender)
	ShowHelp()
End Sub

Sub img_HelpDisabledClick(Sender)
	ShowHelp()
End Sub

Sub img_ECADModelsDisabledMouseMove(Sender, Shift, X, Y)
	img_ECADModelsDisabled.Cursor = crHandPoint
End Sub

Sub img_ECADModelsMouseMove(Sender, Shift, X, Y)
	img_ECADModels.Cursor = crHandPoint
End Sub

Sub img_SettingsMouseMove(Sender, Shift, X, Y)
	img_Settings.Cursor = crHandPoint
End Sub

Sub img_SettingsDisabledMouseMove(Sender, Shift, X, Y)
	img_SettingsDisabled.Cursor = crHandPoint
End Sub

Sub img_ShareDisabledMouseMove(Sender, Shift, X, Y)
	img_ShareDisabled.Cursor = crHandPoint
End Sub

Sub img_ShareMouseMove(Sender, Shift, X, Y)
	img_Share.Cursor = crHandPoint
End Sub

Sub img_HelpDisabledMouseMove(Sender, Shift, X, Y)
	img_HelpDisabled.Cursor = crHandPoint
End Sub

Sub img_HelpMouseMove(Sender, Shift, X, Y)
	img_Help.Cursor = crHandPoint
End Sub

Sub img_DesignDisabledMouseMove(Sender, Shift, X, Y)
	img_DesignDisabled.Cursor = crHandPoint
End Sub

Sub lbl_RegisterMouseMove(Sender, Shift, X, Y)
	lbl_Register.Cursor = crHandPoint
End Sub

Sub lbl_ForgotPasswordMouseMove(Sender, Shift, X, Y)
	lbl_ForgotPassword.Cursor = crHandPoint
End Sub

Sub btn_LoginMouseMove(Sender, Shift, X, Y)
	btn_Login.Cursor = crHandPoint
End Sub

Sub btn_DnldsFldrMouseMove(Sender, Shift, X, Y)
	btn_DnldsFldr.Cursor = crHandPoint
End Sub

Sub btn_SchLibMouseMove(Sender, Shift, X, Y)
	btn_SchLib.Cursor = crHandPoint
End Sub

Sub btn_PcbLibMouseMove(Sender, Shift, X, Y)
	btn_SchLib.Cursor = crHandPoint
End Sub

Sub img_EmailClick(Sender)
	CreateObject("WScript.Shell").Run("mailto:?subject=ALTIUM%20Libraries%20-%20FREE&body=Hi%20%0A%0AYou%20have%20to%20check%20out%20these%20Free%20PCB%20libraries%20for%20Altium,%20AMAZING!%0AGet%20the%20Altium%20Library%20Loader%20from%20https://componentsearchengine.com/tools")
End Sub

Sub img_GetStartedMouseMove(Sender, Shift, X, Y)
	img_GetStarted.Cursor = crHandPoint
End Sub

Sub img_InstructionMouseMove(Sender, Shift, X, Y)
	img_Instruction.Cursor = crHandPoint
End Sub

Sub img_ContactUsMouseMove(Sender, Shift, X, Y)
	img_ContactUs.Cursor = crHandPoint
End Sub

Sub img_ReviewMouseMove(Sender, Shift, X, Y)
	img_Review.Cursor = crHandPoint
End Sub

Sub img_FeedbackMouseMove(Sender, Shift, X, Y)
	img_Feedback.Cursor = crHandPoint
End Sub

Sub img_AboutUsMouseMove(Sender, Shift, X, Y)
	img_AboutUs.Cursor = crHandPoint
End Sub

Sub img_FacebookMouseMove(Sender, Shift, X, Y)
	img_Facebook.Cursor = crHandPoint
End Sub

Sub img_EmailMouseMove(Sender, Shift, X, Y)
	img_Email.Cursor = crHandPoint
End Sub

Sub img_TwitterMouseMove(Sender, Shift, X, Y)
	img_Twitter.Cursor = crHandPoint
End Sub

Sub img_LinkedInMouseMove(Sender, Shift, X, Y)
	img_LinkedIn.Cursor = crHandPoint
End Sub

Sub img_FacebookClick(Sender)
	CreateObject("WScript.Shell").Run("https://www.facebook.com/SamacSysLtd/")
End Sub

Sub img_TwitterClick(Sender)
	CreateObject("WScript.Shell").Run("https://twitter.com/SamacSys")
End Sub

Sub img_LinkedInClick(Sender)
	CreateObject("WScript.Shell").Run("https://www.linkedin.com/company/samacsys-ltd")
End Sub

Sub img_ContactUsClick(Sender)
	CreateObject("WScript.Shell").Run("https://www.samacsys.com/aboutus/contactus/")
End Sub

Sub img_AboutUsClick(Sender)
	CreateObject("WScript.Shell").Run("https://www.samacsys.com/aboutus/")
End Sub

Sub img_FeedbackClick(Sender)
	CreateObject("WScript.Shell").Run("https://www.samacsys.com/case-study-entry/")
End Sub

Sub pnl_SearchClick(Sender)
	PerformSearch txt_Keywords.Text, vbNullString, false, "ad"
End Sub

Sub PerformSearch(Keywords, partID, match, partner)
	'Perform Search using the entered Keywords and process returned JSON response
	If CheckLogin Then
		Screen.Cursor = crHourglass
		lbl_Message.Caption = vbNullString
		ClearGrid()
		HashIndex = InStr(Keywords,"#")
		If HashIndex <> 0 Then
			Keywords = Mid(Keywords,1,HashIndex-1)
			match = false
		End If
		jsonResponse = httpGET("https://" & partner & ".componentsearchengine.com/ga/search.php?kws=" & URLEncode(Keywords) & "&v=1", txt_Username.Text, txt_Password.Text)
		JSONToXML jsonResponse, Keywords, match
		txt_Keywords.Text = vbNullString
		Screen.Cursor = crDefault
		If lbl_Message.Caption <> "No parts found. Click here to request this part" Then
			AutoSizeCol StringGrid1,0
			AutoSizeCol StringGrid1,1
			AutoSizeCol StringGrid1,2
			AutoSizeCol StringGrid1,3
			AutoSizeCol StringGrid1,4

			StringGrid1.Row = 1
			ShowImages()
		Else
			HideImages()
		End If
	Else
		ShowSettings()
		lbl_Message.Caption = "Please Login with Email and Password before continuing..."
	End If

End Sub

Sub ClearGrid()
	For r = 1 To 10
		StringGrid1.Rows(r).Clear()
	Next
End Sub

Const stateRoot = 0
Const stateNameQuoted = 1
Const stateNameFinished = 2
Const stateValue = 3
Const stateValueQuoted = 4
Const stateValueQuotedEscaped = 5
Const stateValueQuotedEscapedHex = 6
Const stateValueUnquoted = 7
Const stateValueUnquotedEscaped = 8

Function JSONToXML(json, Keywords, match)
	Dim dom, xmlElem, i, ch, state, name, value, sHex
	Set dom = CreateObject("Microsoft.XMLDOM")
	state = stateRoot
	For i = 1 to Len(json)
		ch = Mid(json, i, 1)
		Select Case state
		Case stateRoot
			Select Case ch
			Case "["
				If dom.documentElement is Nothing Then
					Set xmlElem = dom.CreateElement("ARRAY")
					Set dom.documentElement = xmlElem
				Else
					Set xmlElem = XMLCreateChild(xmlElem, "ARRAY")
				End If
			Case "{"
				If dom.documentElement is Nothing Then
					Set xmlElem = dom.CreateElement("ROOT")
					Set dom.documentElement = xmlElem
				Else
					Set xmlElem = XMLCreateChild(xmlElem, "OBJECT")
				End If
			Case """"
				state = stateNameQuoted
				name = ""
			Case "}"
				Set xmlElem = xmlElem.parentNode
			Case "]"
				Set xmlElem = xmlElem.parentNode
			End Select
		Case stateNameQuoted
			Select Case ch
			Case """"
				state = stateNameFinished
			Case Else
				name = name + ch
			End Select
		Case stateNameFinished
			Select Case ch
			Case ":"
				value = ""
				State = stateValue
			Case Else                     '@@Enhancement#1: Handling Array values
				Set xmlitem = dom.createTextNode(name)
				xmlElem.appendChild(xmlitem)
				State = stateRoot
			End Select
		Case stateValue
			Select Case ch
			Case """"
				State = stateValueQuoted
			Case "{"
				Set xmlElem = XMLCreateChild(xmlElem, name)
				State = stateRoot
			Case "["
				Set xmlElem = XMLCreateChild(xmlElem, name)
				State = stateRoot
			Case " "
			Case Chr(9)
			Case vbCr
			Case vbLF
			Case Else
				value = ch
				State = stateValueUnquoted
			End Select
		Case stateValueQuoted
			Select Case ch
			Case """"
				xmlElem.setAttribute name, value
				state = stateRoot
			Case "\"
				state = stateValueQuotedEscaped
			Case Else
				value = value + ch
			End Select
		Case stateValueQuotedEscaped ' @@Enhancement#2: Handle escape sequences
			If ch = "u" Then  'Four digit hex. Ex: o = 00f8
				sHex = ""
				state = stateValueQuotedEscapedHex
			Else
				Select Case ch
				Case """"
					value = value + """"
				Case "\"
					value = value + "\"
				Case "/"
					value = value + "/"
				Case "b"    'Backspace
					value = value + chr(08)
				Case "f"    'Form-Feed
					value = value + chr(12)
				Case "n"    'New-line (LineFeed(10))
					value = value + vbLF
				Case "r"    'New-line (CarriageReturn/CRLF(13))
					value = value + vbCR
				Case "t"    'Horizontal-Tab (09)
					value = value + vbTab
				Case Else
					'do not accept any other escape sequence
				End Select
				state = stateValueQuoted
			End If
		Case stateValueQuotedEscapedHex
			sHex = sHex + ch
			If len(sHex) = 4 Then
				on error resume next
				value = value + Chr("&H" & sHex)    'Hex to String conversion
				on error goto 0
				state = stateValueQuoted
			End If
		Case stateValueUnquoted
			Select Case ch
			Case "}"
				xmlElem.setAttribute name, value
				Set xmlElem = xmlElem.parentNode
				state = stateRoot
			Case "]"
				xmlElem.setAttribute name, value
				Set xmlElem = xmlElem.parentNode
				state = stateRoot
			Case ","
				xmlElem.setAttribute Replace(name, " ", ""), value
				state = stateRoot
			Case "\"
				state = stateValueUnquotedEscaped
			Case Else
				value = value + ch
			End Select
		Case stateValueUnquotedEscaped ' @@TODO: Handle escape sequences
			value = value + ch
			state = stateValueUnquoted
		End Select
	Next

	Set objNodeList = dom.getElementsByTagName("ROOT")
	For each x in objNodeList
		PartCount = x.getAttribute("PartCount")
	Next

	If PartCount > 0 Then
		p = 1
		Set Parts = dom.selectNodes ("//ROOT/Parts/")
		For Each Part in Parts
			If Left(Part.nodeName,3) = "pid" Then
				partID = Mid(Part.nodeName,4)
				Manuf = Part.getAttribute("Manufacturer")
				PartNo = Part.getAttribute("PartNumber")
				Desc = Part.getAttribute("Description")
				Symbol = Part.getAttribute("Symbol")
				Image3D = Part.getAttribute("Image3D")
				If Symbol = "null" Then
					ECAD_M = "N"
				Else
					ECAD_M = "Y"
				End If
				Have3D = Part.getAttribute("Have3D")
				If Have3D = 1 Then
					Have3D = "Y"
				Else
					Have3D = "N"
				End If

				Include = False
				If match Then
					If Keywords = PartNo Then
						Include = True
					End If
				Else
					Include = True
				End If

				If ECAD_M = "Y" And Have3D = "Y" And Include Then
					StringGrid1.Cols(0)(p) = ECAD_M
					StringGrid1.Cols(1)(p) = Have3D
					StringGrid1.Cols(2)(p) = Manuf
					StringGrid1.Cols(3)(p) = PartNo
					StringGrid1.Cols(4)(p) = Desc
					StringGrid1.Cols(5)(p) = "SamacSys"
					StringGrid1.Cols(6)(p) = PartID
					StringGrid1.Cols(7)(p) = Image3D
					p= p+1
				End If
			End If
		Next

		Set Parts = dom.selectNodes ("//ROOT/Parts/")
		For Each Part in Parts
			If Left(Part.nodeName,3) = "pid" Then
				partID = Mid(Part.nodeName,4)
				Manuf = Part.getAttribute("Manufacturer")
				PartNo = Part.getAttribute("PartNumber")
				Desc = Part.getAttribute("Description")
				Symbol = Part.getAttribute("Symbol")
				Image3D = Part.getAttribute("Version")
				If Symbol = "null" Then
					ECAD_M = "N"
				Else
					ECAD_M = "Y"
				End If
				Have3D = Part.getAttribute("Have3D")
				If Have3D = 1 Then
					Have3D = "Y"
				Else
					Have3D = "N"
				End If

				Include = False
				If match = True Then
					If Keywords = PartNo Then
						Include = True
					End If
				Else
					Include = True
				End If

				If ECAD_M = "Y" And Have3D = "N" And Include Then
					StringGrid1.Cols(0)(p) = ECAD_M
					StringGrid1.Cols(1)(p) = Have3D
					StringGrid1.Cols(2)(p) = Manuf
					StringGrid1.Cols(3)(p) = PartNo
					StringGrid1.Cols(4)(p) = Desc
					StringGrid1.Cols(5)(p) = "SamacSys"
					StringGrid1.Cols(6)(p) = PartID
					StringGrid1.Cols(7)(p) = Image3D
					p= p+1
				End If
			End If
		Next

		Set Parts = dom.selectNodes ("//ROOT/Parts/")
		For Each Part in Parts
			If Left(Part.nodeName,3) = "pid" Then
				partID = Mid(Part.nodeName,4)
				Manuf = Part.getAttribute("Manufacturer")
				PartNo = Part.getAttribute("PartNumber")
				Desc = Part.getAttribute("Description")
				Symbol = Part.getAttribute("Symbol")
				Image3D = Part.getAttribute("Image3D")
				If Symbol = "null" Then
					ECAD_M = "N"
				Else
					ECAD_M = "Y"
				End If
				Have3D = Part.getAttribute("Have3D")
				If Have3D = 1 Then
					Have3D = "Y"
				Else
					Have3D = "N"
				End If

				Include = False
				If match = True Then
					If Keywords = PartNo Then
						Include = True
					End If
				Else
					Include = True
				End If

				If ECAD_M = "N" And Have3D = "Y" And Include Then
					StringGrid1.Cols(0)(p) = ECAD_M
					StringGrid1.Cols(1)(p) = Have3D
					StringGrid1.Cols(2)(p) = Manuf
					StringGrid1.Cols(3)(p) = PartNo
					StringGrid1.Cols(4)(p) = Desc
					StringGrid1.Cols(5)(p) = "SamacSys"
					StringGrid1.Cols(6)(p) = PartID
					StringGrid1.Cols(7)(p) = Image3D
					p= p+1
				End If
			End If
		Next

		Set Parts = dom.selectNodes ("//ROOT/Parts/")
		For Each Part in Parts
			If Left(Part.nodeName,3) = "pid" Then
				partID = Mid(Part.nodeName,4)
				Manuf = Part.getAttribute("Manufacturer")
				PartNo = Part.getAttribute("PartNumber")
				Desc = Part.getAttribute("Description")
				Symbol = Part.getAttribute("Symbol")
				Image3D = Part.getAttribute("Image3D")
				If Symbol = "null" Then
					ECAD_M = "N"
				Else
					ECAD_M = "Y"
				End If
				Have3D = Part.getAttribute("Have3D")
				If Have3D = 1 Then
					Have3D = "Y"
				Else
					Have3D = "N"
				End If

				Include = False
				If match Then
					If Keywords = PartNo Then
						Include = True
					End If
				Else
					Include = True
				End If

				If ECAD_M = "N" And Have3D = "N" And Include Then
					StringGrid1.Cols(0)(p) = ECAD_M
					StringGrid1.Cols(1)(p) = Have3D
					StringGrid1.Cols(2)(p) = Manuf
					StringGrid1.Cols(3)(p) = PartNo
					StringGrid1.Cols(4)(p) = Desc
					StringGrid1.Cols(5)(p) = "SamacSys"
					StringGrid1.Cols(6)(p) = PartID
					StringGrid1.Cols(7)(p) = Image3D
					p= p+1
				End If
			End If
		Next
	Else
		lbl_Message.Caption = "No parts found. Click here to request this part"
		lbl_Message.Font.Style = 4
	End If

	Set JSONToXML = dom
End Function

Function XMLCreateChild(xmlParent, tagName)
	Dim xmlChild
	If xmlParent is Nothing Then
		Set XMLCreateChild = Nothing
		Exit Function
	End If
	If xmlParent.ownerDocument is Nothing Then
		Set XMLCreateChild = Nothing
		Exit Function
	End If
	If IsNumeric(tagName) Then tagName = "pid" & tagName
	Set xmlChild = xmlParent.ownerDocument.createElement(tagName)
	xmlParent.appendChild xmlChild
	Set XMLCreateChild = xmlChild
End Function

Sub pnl_SearchMouseMove(Sender, Shift, X, Y)
	pnl_Search.Cursor = crHandPoint
End Sub

Sub StringGrid1DblClick(Sender)
	selectedRow = StringGrid1.Row
	MsgBox StringGrid1.Cols(6)(selectedRow)
End Sub

Sub pnl_MoreInfoClick(Sender)
	selectedRow = StringGrid1.Row
	partID = StringGrid1.Cols(6)(selectedRow)
	If IsNumeric(partID) Then CreateObject("WScript.Shell").Run("https://ad.componentsearchengine.com/part.php?partID=" & partID)
End Sub

Sub pnl_DesignClick(Sender)
	lbl_Message.Caption = "Processing ECAD Model, Please Wait..."
	Form1.Refresh()
	If Not ProcessSelectedPart(vbNullString) Then lbl_Message.Caption = vbNullString
End Sub

Sub pnl_MoreInfoMouseMove(Sender, Shift, X, Y)
	pnl_MoreInfo.Cursor = crHandPoint
End Sub

Sub pnl_DesignMouseMove(Sender, Shift, X, Y)
	pnl_Design.Cursor = crHandPoint
End Sub

Sub pnl_OpenModelMouseMove(Sender, Shift, X, Y)
	pnl_OpenModel.Cursor = crHandPoint
End Sub

Sub StringGrid1MouseMove(Sender, Shift, X, Y)
	StringGrid1.Cursor = crHandPoint
End Sub

Sub pnl_OpenModelClick(Sender)

	OpenedZIP = False
	OpenedEPW = False
	epwMPN = vbNullString

	partID = vbNullString
	ecadModelPath = txt_DnldsFldr.Text & "\"

	OpenDialog3.InitialDir = ecadModelPath
	OpenDialog3.Filter = "ECAD Model Files|*.epw;*-part*.zip;*-PCB-*.zip;LIB_*.zip"
	If OpenDialog3.Execute Then

		lbl_Message.Caption = "Processing ECAD Model, Please Wait..."
		ClearGrid()
		HideImages()

		ecadModelPath = InstalledDir & "\Temp\"

		ecadModelFile = OpenDialog3.Filename

		SlashIndex = InStrRev(ecadModelFile,"\")
		ecadModelFilename = Mid(ecadModelFile,SlashIndex + 1)

		If Right(ecadModelFile,4) = ".zip" Then
			OpenedZIP = True
			Set sa = CreateObject("Shell.Application")
			ExtractEPW sa.NameSpace(ecadModelFile), ecadModelPath
			Set sa = Nothing
			epwFile = ecadModelPath & cbFilename
		Else
			OpenedEPW = True
			epwFile = ecadModelFile
		End If
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.OpenTextFile(epwFile)
		l=1
		Do Until f.AtEndOfStream
			epwFileLine = f.ReadLine
			If l=1 And IsNumeric(epwFileLine) Then partID = epwFileLine
			If Mid(epwFileLine, 1, 4) = "mpn=" Then epwMPN = Mid(epwFileLine, 5)
			If Mid(epwFileLine, 1, 2) = "p=" Then epwMPN = Mid(epwFileLine, 3)
			If Mid(epwFileLine, 1, 4) = "pna=" Then epwPNA = Mid(epwFileLine, 5)
			If Mid(epwFileLine, 1, 2) = "w=" Then epwWSP = Mid(epwFileLine, 3)
			l = l+1
		Loop

		f.Close

		If partID <> vbNullString And epwMPN = vbNullString Then
			If ProcessSelectedPart(partID) Then
				If OpenedZIP Then fso.DeleteFile(ecadModelFile)
				fso.DeleteFile(epwFile)
			Else
				lbl_Message.Caption = vbNullString
				If OpenedZIP Then fso.DeleteFile(epwFile)
			End If
		Else
			If epwPNA = vbNullString Then epwPNA = epwWSP
			PerformSearch epwMPN, partID, true, epwPNA
			StringGrid1.Cols(5)(1) = epwPNA
			If OpenedZIP Then fso.DeleteFile(ecadModelFile)
			fso.DeleteFile(epwFile)
		End If

		Set fso = Nothing

	End If
End Sub

Sub AutoSizeCol(Grid, Column)
	Dim i, W, WMax
	WMax = 0
	for i = 0 to (Grid.RowCount - 1)
		W = Grid.Canvas.TextWidth(Grid.Cells(Column, i))
		If W > WMax then
			WMax = W
		End If
	Next
	Grid.ColWidths(Column) = WMax + 10
End Sub

Sub StringGrid1Click(Sender)
	ShowImages()
End Sub

Sub ShowImages()
	Dim fso
	Dim ImagePath
	Dim selectedRow
	Dim partID
	Dim SymPcb
	Dim ThreeD
	Dim Image3D
	Dim xHttp
	Dim symPath
	Dim pcbPath
	Dim small3DPath

	On Error Resume Next

	Set fso = CreateObject("Scripting.FileSystemObject")

	ImagePath = InstalledDir & "\SamacSys_Images"

	If Not fso.FolderExists(InstalledDir) Then
		fso.CreateFolder(InstalledDir)
	End If

	If Not fso.FolderExists(ImagePath) Then
		fso.CreateFolder(ImagePath)
	End If

	selectedRow = StringGrid1.Row
	partID = StringGrid1.Cols(6)(selectedRow)
	SymPcb = StringGrid1.Cols(0)(selectedRow)
	ThreeD = StringGrid1.Cols(1)(selectedRow)
	Image3D = StringGrid1.Cols(7)(selectedRow)

	If Not IsNumeric(partID) Then
		HideImages()
		Exit Sub
	End If

	symPath = ImagePath & "\SYM_" & partID & ".png"
	pcbPath = ImagePath & "\PCB_" & partID & ".png"
	small3DPath = ImagePath & "\3DM_" & partID & "_small.png"

	If SymPcb = "Y" Then
		img_SYM.Visible = True
		img_PCB.Visible = True

		If fso.FileExists(symPath) And fso.FileExists(pcbPath) Then
			img_SYM.Picture.LoadFromFile symPath
			img_PCB.Picture.LoadFromFile pcbPath
		Else
			Set xHttp = CreateHttpRequest()

			If Not xHttp Is Nothing Then
				If Not fso.FileExists(symPath) Then
					DownloadImageToFile xHttp, "https://ad.componentsearchengine.com/symbol.php?partID=" & partID, symPath
				End If

				If Not fso.FileExists(pcbPath) Then
					DownloadImageToFile xHttp, "https://ad.componentsearchengine.com/footprint.php?partID=" & partID, pcbPath
				End If
			End If

			If fso.FileExists(symPath) Then
				img_SYM.Picture.LoadFromFile symPath
			Else
				img_SYM.Visible = False
			End If

			If fso.FileExists(pcbPath) Then
				img_PCB.Picture.LoadFromFile pcbPath
			Else
				img_PCB.Visible = False
			End If
		End If
	Else
		img_SYM.Visible = False
		img_PCB.Visible = False
	End If

	If ThreeD = "Y" And Image3D <> vbNullString Then
		img_3DM.Visible = True

		If fso.FileExists(small3DPath) Then
			img_3DM.Picture.LoadFromFile small3DPath
		Else
			CreatePlaceholderImage ImagePath & "\LOAD_3D.png", "LOAD 3D"
			img_3DM.Picture.LoadFromFile ImagePath & "\LOAD_3D.png"
		End If
	Else
		img_3DM.Visible = True
		CreatePlaceholderImage ImagePath & "\NO_3D.png", "NO 3D"
		img_3DM.Picture.LoadFromFile ImagePath & "\NO_3D.png"
	End If

	Set xHttp = Nothing
	Set fso = Nothing

	On Error GoTo 0
End Sub

Sub CreatePlaceholderImage(filePath, text)
	Dim fso

	Set fso = CreateObject("Scripting.FileSystemObject")

	If fso.FileExists(filePath) Then
		Exit Sub
	End If

	If text = "LOAD 3D" Then
		fso.CopyFile InstalledDir & "\LOAD_3D.png", filePath, True
	Else
		fso.CopyFile InstalledDir & "\NO_3D.png", filePath, True
	End If

	Set fso = Nothing
End Sub

Sub ResizeImageForPreview(srcPath, dstPath, maxW, maxH)
	Dim img
	Dim proc
	Dim outImg
	Dim ratioW
	Dim ratioH
	Dim scale
	Dim newW
	Dim newH
	Dim fso

	Set fso = CreateObject("Scripting.FileSystemObject")

	If fso.FileExists(dstPath) Then
		Exit Sub
	End If

	On Error Resume Next

	Set img = CreateObject("WIA.ImageFile")
	img.LoadFile srcPath

	If Err.Number <> 0 Then
		Err.Clear
		On Error GoTo 0
		Exit Sub
	End If

	ratioW = CDbl(maxW) / CDbl(img.Width)
	ratioH = CDbl(maxH) / CDbl(img.Height)

	If ratioW < ratioH Then
		scale = ratioW
	Else
		scale = ratioH
	End If

	If scale > 1 Then
		scale = 1
	End If

	newW = CInt(CDbl(img.Width) * scale)
	newH = CInt(CDbl(img.Height) * scale)

	Set proc = CreateObject("WIA.ImageProcess")
	proc.Filters.Add proc.FilterInfos("Scale").FilterID
	proc.Filters(1).Properties("MaximumWidth") = newW
	proc.Filters(1).Properties("MaximumHeight") = newH

	Set outImg = proc.Apply(img)
	outImg.SaveFile dstPath

	Err.Clear
	On Error GoTo 0
End Sub

Function DownloadImageToFile(xHttp, url, filePath)

	Dim bStrm

	DownloadImageToFile = False

	If CreateObject("Scripting.FileSystemObject").FileExists(filePath) Then
		DownloadImageToFile = True
		Exit Function
	End If

	On Error Resume Next

	Err.Clear

	xHttp.Open "GET", url, False

	If chk_Proxy.Checked Then
		xHttp.setProxy 2, Trim(txt_Address.Text) & ":" & Trim(txt_Port.Text), ""
	End If

	xHttp.Send

	If Err.Number <> 0 Then
		Err.Clear
		Exit Function
	End If

	If xHttp.Status < 200 Or xHttp.Status >= 300 Then
		Exit Function
	End If

	Set bStrm = CreateObject("Adodb.Stream")

	With bStrm
		.Type = 1
		.Open
		.Write xHttp.responseBody
		.SaveToFile filePath, 2
		.Close
	End With

	Set bStrm = Nothing

	If Err.Number <> 0 Then
		Err.Clear
		Exit Function
	End If

	DownloadImageToFile = True

	On Error GoTo 0

End Function

Sub HideImages()
	img_SYM.Visible = False
	img_PCB.Visible = False
	img_3DM.Visible = False
End Sub

Sub img_SYMMouseMove(Sender, Shift, X, Y)
	img_SYM.Cursor = crHandPoint
End Sub

Sub img_PCBMouseMove(Sender, Shift, X, Y)
	img_PCB.Cursor = crHandPoint
End Sub

Sub img_3DMMouseMove(Sender, Shift, X, Y)
	img_3DM.Cursor = crHandPoint
End Sub

Sub img_SYMClick(Sender)
	selectedRow = StringGrid1.Row
	partID = StringGrid1.Cols(6)(selectedRow)
	CreateObject("WScript.Shell").Run("https://ad.componentsearchengine.com/footprintPreview.php?partID=" & partID & "&u=0&target=symbol")
End Sub

Sub img_PCBClick(Sender)
	selectedRow = StringGrid1.Row
	partID = StringGrid1.Cols(6)(selectedRow)
	CreateObject("WScript.Shell").Run("https://ad.componentsearchengine.com/footprintPreview.php?partID=" & partID & "&u=0")
End Sub

Sub img_3DMClick(Sender)
	Dim fso
	Dim ImagePath
	Dim selectedRow
	Dim partID
	Dim ThreeD
	Dim Image3D
	Dim xHttp
	Dim filePath
	Dim smallPath

	Set fso = CreateObject("Scripting.FileSystemObject")

	ImagePath = InstalledDir & "\SamacSys_Images"

	If Not fso.FolderExists(ImagePath) Then
		fso.CreateFolder(ImagePath)
	End If

	selectedRow = StringGrid1.Row
	partID = StringGrid1.Cols(6)(selectedRow)
	ThreeD = StringGrid1.Cols(1)(selectedRow)
	Image3D = StringGrid1.Cols(7)(selectedRow)

	If ThreeD <> "Y" Or Image3D = vbNullString Then
		Exit Sub
	End If

	filePath = ImagePath & "\3DM_" & partID & ".png"
	smallPath = ImagePath & "\3DM_" & partID & "_small.png"

	If fso.FileExists(smallPath) Then
		img_3DM.Picture.LoadFromFile smallPath
		Exit Sub
	End If

	Set xHttp = CreateHttpRequest()

	If xHttp Is Nothing Then
		Exit Sub
	End If

	If DownloadImageToFile(xHttp, Image3D, filePath) Then
		ResizeImageForPreview filePath, smallPath, 220, 220

		If fso.FileExists(smallPath) Then
			img_3DM.Picture.LoadFromFile smallPath
		Else
			CreatePlaceholderImage ImagePath & "\LOAD_3D.png", "LOAD 3D"
			img_3DM.Picture.LoadFromFile ImagePath & "\LOAD_3D.png"
		End If
	End If

	Set xHttp = Nothing
	Set fso = Nothing
End Sub

Sub txt_KeywordsKeyDown(Sender, Key, Shift)
	If Key = 13 Then
		PerformSearch txt_Keywords.Text, vbNullString, false, "ad"
	End If
End Sub

Sub lbl_MessageMouseMove(Sender, Shift, X, Y)
	If lbl_Message.Caption = "No parts found. Click here to request this part" Then lbl_Message.Cursor = crHandPoint
End Sub

Sub lbl_MessageClick(Sender)
	If lbl_Message.Caption = "No parts found. Click here to request this part" Then
		CreateObject("WScript.Shell").Run("https://ad.componentsearchengine.com/newPart.php")
		lbl_Message.Caption = vbNullString
		lbl_Message.Cursor = crDefault
	End If
End Sub

Sub txt_UsernameKeyDown(Sender, Key, Shift)
	If Key = 13 Then
		txt_Password.setfocus
	End If
End Sub

Sub txt_PasswordKeyDown(Sender, Key, Shift)
	If Key = 13 Then
		Login()
	End If
End Sub

Sub img_InstructionClick(Sender)
	CreateObject("WScript.Shell").Run("https://www.samacsys.com/altium-designer-library-instructions")
End Sub

Sub img_GetStartedClick(Sender)
	CreateObject("WScript.Shell").Run("https://www.samacsys.com/altiumll-get-started/")
End Sub

Sub img_ReviewClick(Sender)
	CreateObject("WScript.Shell").Run("https://www.samacsys.com/altiumll-review/")
End Sub

Sub chk_AltiumSymbolsClick(Sender)
	UpdateTXT
End Sub

Sub chk_ProxyClick(Sender)
	If chk_Proxy.Checked Then
		txt_Address.Enabled = True
		txt_Port.Enabled = True
	Else
		txt_Address.Enabled = False
		txt_Port.Enabled = False
	End If
	UpdateTXT
End Sub

Function strClean(strtoclean)
	Dim objRegExp, outputStr
	Set objRegExp = New Regexp

	objRegExp.IgnoreCase = True
	objRegExp.Global = True

	objRegExp.Pattern = "[\\/:\*\?""<>\|]"
	outputStr = objRegExp.Replace(strtoclean,"_")

	strClean = outputStr
End Function
