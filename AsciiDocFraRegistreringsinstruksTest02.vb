Option Explicit

!INC Local Scripts.EAConstants-VBScript

' Script Name: AsciiDocFraRegistreringsinstruksTest02
' Original Author: Tore Johnsen, Åsmund Tjora
' Purpose: Generate documentation in AsciiDoc syntax
' Date: 08.04.2021
'
' Version: 0.2
' Date: 2021-04-16 Kent Jonsrud:
' Use Case: Guidance for photogrammetric data collection
' Actor is a person digitizing from aeriel photos
' Design Goal: compact summary of all guiding requirements
'
' TBD: feature type: table of tags from supertypes as FKB-standard A/B/C/D
' - feature type: additional definition (Minstestørrelse etc.)
' TBD: feature type: registreringsmetode
' TBD: feature type: tilleggsbeskrivelse
' TBD: feature type: forhold til andre objekttyper
' TBD: feature type: ?
' - feature attributes: table of all inherited attributes - name, SOSI_navn, type, length(?), FKB-standard A/B/C/D
' - one picture (and related picture text?) from each supertype
' - several pictures and related picture text from subtype

' TBD: common tables with attributes ??? really?

' TBD: common table of quality requirements for all feature types
' TBD: common table of quality requirements on each feature type
'	

' TBD: common table 
' TBD: opprydding !!!
'
Dim imgfolder, imgparent
Dim diagCounter

'
' Project Browser Script main function
Sub OnProjectBrowserScript()

    Dim treeSelectedType
    treeSelectedType = Repository.GetTreeSelectedItemType()

    Select Case treeSelectedType

        Case otPackage
			Repository.EnsureOutputVisible "Script"
			Repository.ClearOutput "Script"
            ' Code for when a package is selected
			imgfolder ="img"
			diagCounter = 0
			Dim thePackage As EA.Package
			set thePackage = Repository.GetTreeSelectedObject()
			Call ListAsciiDoc(thePackage)

        Case Else
            ' Error message
            Session.Prompt "This script does not support items of this type.", promptOK

    End Select

End Sub


Sub ListAsciiDoc(thePackage)
Dim imgFSO
Dim element As EA.Element
dim tag as EA.TaggedValue
Dim diag As EA.Diagram
Dim projectclass As EA.Project
set projectclass = Repository.GetProjectInterface()


	
Session.Output("=== "&thePackage.Name&"")
Session.Output("Definisjon: "&thePackage.Notes&"")

if thePackage.element.TaggedValues.Count > 0 then
	Session.Output(" ")	
	Session.Output("===== Tagged Values")
	Session.Output("[cols=""20,80""]")
	Session.Output("|===")
	for each tag in thePackage.element.TaggedValues
		if tag.Value <> "" then	
			if tag.Name <> "persistence" and tag.Name <> "SOSI_melding" then
			'	Session.Output("|Tag: "&tag.Name&"")
			'	Session.Output("|Verdi: "&tag.Value&"")
				Session.Output("|"&tag.Name&"")
				Session.Output("|"&tag.Value&"")
				Session.Output(" ")			
			end if
		end if
		if tag.Name = "SOSI_registreringsinstruks" then	
			imgfolder = tag.Value
		end if
	next

	Session.Output("|===")
end if
'-----------------Diagram-----------------
	Set imgFSO=CreateObject("Scripting.FileSystemObject")
	imgparent = imgFSO.GetParentFolderName(Repository.ConnectionString())  & "\" & imgfolder
	if false then				
		Session.Output(" DEBUG.")
		Session.Output(" imgfolder: " & imgfolder  & "...")
		Session.Output(" imgFSO.GetAbsolutePathName: " & imgFSO.GetAbsolutePathName("./")  & "...")
		Session.Output(" Repository.ConnectionString: " & Repository.ConnectionString() & "...")
		Session.Output(" imgFSO.GetBaseName: " & imgFSO.GetBaseName(Repository.ConnectionString())  & "...")
		Session.Output(" imgFSO.GetParentFolderName: " & imgFSO.GetParentFolderName(Repository.ConnectionString())  & "...")
		Session.Output(" imgparent: " & imgparent  & "...")
	end if
	if not imgFSO.FolderExists(imgparent) then
		imgFSO.CreateFolder imgparent
	end if

	for each tag in thePackage.element.TaggedValues
		if tag.Name = "SOSI_bildeAvModellelement" and tag.Value <> "" then
			diagCounter = diagCounter + 1
			Session.Output("[caption=""Figur "&diagCounter&": "",title="&tag.Name&"]")
		'	Session.Output("image::"&tag.Value&".png["&ThePackage.Name"."&tag.Name&"]")
			Session.Output("image::"&tag.Value&".png["&tag.Value&"]")
		end if
	next

For Each diag In thePackage.Diagrams
	diagCounter = diagCounter + 1
	Call projectclass.PutDiagramImageToFile(diag.DiagramGUID, imgparent & "\" & diag.Name & ".png", 1)
'	Call projectclass.PutDiagramImageToFile(diag.DiagramGUID, "" & diag.Name&".png", 1)
	Repository.CloseDiagram(diag.DiagramID)
	Session.Output("[caption=""Figur "&diagCounter&": "",title="&diag.Name&"]")
	Session.Output("image::"&diag.Name&".png["&diag.Name&"]")
Next

For each element in thePackage.Elements
'	If Ucase(element.Stereotype) = "FEATURETYPE" Then
	If Ucase(element.Stereotype) = "" Then
		Call ObjektOgDatatyper(element)
	End if
Next
	
For each element in thePackage.Elements
	If Ucase(element.Stereotype) = "DATATYPE" Then
		Call ObjektOgDatatyper(element)
	End if
Next

For each element in thePackage.Elements
	If Ucase(element.Stereotype) = "UNION" Then
		Call ObjektOgDatatyper(element)
	End if
Next

For each element in thePackage.Elements
	If Ucase(element.Stereotype) = "CODELIST" Then
		Call Kodelister(element)
	End if
	If Ucase(element.Stereotype) = "ENUMERATION" Then
		Call Kodelister(element)
	End if
	If element.Type = "Enumeration" Then
		Call Kodelister(element)
	End if
Next
	
dim pack as EA.Package
for each pack in thePackage.Packages
	Call ListAsciiDoc(pack)
next

Set imgFSO = Nothing
end sub

'-----------------ObjektOgDatatyper-----------------
Sub ObjektOgDatatyper(element)
Dim att As EA.Attribute
dim tag as EA.TaggedValue
Dim con As EA.Connector
Dim supplier As EA.Element
Dim client As EA.Element
Dim association
Dim aggregation
association = False
Dim generalizations
Dim numberSpecializations ' tar også med antall realiseringer her
Dim textVar
dim externalPackage

Session.Output(" ")
Session.Output("==== «"&element.Stereotype&"» "&element.Name&"")
For Each con In element.Connectors
	set supplier = Repository.GetElementByID(con.SupplierID)
	If con.Type = "Generalization" And supplier.ElementID <> element.ElementID Then
		Session.Output("Inherited definition: "&supplier.Notes&"")
		Session.Output(" ")
		numberSpecializations = numberSpecializations + 1
	End If
Next
Session.Output("Definisjon: "&element.Notes&"")
Session.Output(" ")
numberSpecializations = 0
For Each con In element.Connectors
	set supplier = Repository.GetElementByID(con.SupplierID)
	If con.Type = "Generalization" And supplier.ElementID <> element.ElementID Then
		Session.Output("*Supertype:* «" & supplier.Stereotype&"» "&supplier.Name&"")
		Session.Output(" ")
		numberSpecializations = numberSpecializations + 1
	End If
Next
For Each con In element.Connectors  
'realiseringer.  
'Må forbedres i framtidige versjoner dersom denne skal med 
'- full sti (opp til applicationSchema eller øverste pakke under "Model") til pakke som inneholder klassen som realiseres
	set supplier = Repository.GetElementByID(con.SupplierID)
	If con.Type = "Realisation" And supplier.ElementID <> element.ElementID Then
		set externalPackage = Repository.GetPackageByID(supplier.PackageID)
		textVar=getPath(externalPackage)
		Session.Output("*Realisering av:* " & textVar &"::«" & supplier.Stereotype&"» "&supplier.Name)
		Session.Output(" ")
		numberSpecializations = numberSpecializations + 1
	end if
next

if element.TaggedValues.Count > 0 then
	Session.Output("===== Tagged Values")
	Session.Output("[cols=""20,80""]")
	Session.Output("|===")
	for each tag in element.TaggedValues								
		if tag.Value <> "" then	
			if tag.Name <> "persistence" and tag.Name <> "SOSI_melding" then
			'	Session.Output("|Tag: "&tag.Name&"")
			'	Session.Output("|Verdi: "&tag.Value&"")
				Session.Output("|"&tag.Name&"")
				Session.Output("|"&tag.Value&"")
				Session.Output(" ")			
			end if
		end if
	next
	Session.Output("|===")
	
	for each tag in element.TaggedValues								
		if tag.Name = "SOSI_bildeAvModellelement" and tag.Value <> "" then
			diagCounter = diagCounter + 1
			Session.Output("[caption=""Figur "&diagCounter&": "",title="&tag.Name&"]")
		'	Session.Output("image::"&tag.Value&".png["&ThePackage.Name"."&tag.Name&"]")
			Session.Output("image::"&tag.Value&"["&tag.Value&"]")
		end if
	next
	if getTaggedValue(element,"SOSI_bildetekstTilModellelement") <> "" then
		Session.Output("Bildebeskrivelse: "& getTaggedValue(element,"SOSI_bildetekstTilModellelement") & "")
		Session.Output(" ")
	end if
end if

if element.AttributesEx.Count > 0 then
	Session.Output("===== Egenskaper")
	for each att in element.AttributesEx
		Session.Output("[cols=""20,80""]")
		Session.Output("|===")
		Session.Output("|*Navn:* ")
		Session.Output("|*"&att.name&"*")
		Session.Output(" ")
		Session.Output("|Definisjon: ")
		Session.Output("|"&getCleanDefinition(att.Notes)&"")
		Session.Output(" ")
		Session.Output("|Multiplisitet: ")
		Session.Output("|["&att.LowerBound&".."&att.UpperBound&"]")
		Session.Output(" ")
		if not att.Default = "" then
			Session.Output("|Initialverdi: ")
			Session.Output("|"&att.Default&"")
			Session.Output(" ")
		end if
		Session.Output("|Type: ")
		if att.ClassifierID <> 0 then
			Session.Output("|«" & Repository.GetElementByID(att.ClassifierID).Stereotype & "» "&att.Type&"")		
		else
			Session.Output("|"&att.Type&"")
		end if

		if att.TaggedValues.Count > 0 then
			Session.Output("|Tagged Values: ")
			Session.Output("|")
			for each tag in att.TaggedValues
				Session.Output(""&tag.Name& ": "&tag.Value&" + ")
			next
		end if
		Session.Output("|===")
	next
	Session.Output("===== Egenskapstabell") 'type, length(?), FKB-standard A/B/C/
	Session.Output("[cols=""15,15,15,7,7,7,7,7""]")
	Session.Output("|===")
	Session.Output("|*Navn:* ")
	Session.Output("|*SOSI_navn:* ")
	Session.Output("|*Type:* ")
	Session.Output("|*Len:* ")
	Session.Output("|*FKB-A:* ")
	Session.Output("|*FKB-B:* ")
	Session.Output("|*FKB-C:* ")
	Session.Output("|*FKB-D:* ")
	Session.Output(" ")
	for each att in element.AttributesEx
		Session.Output("|"&att.name&"")
		Session.Output("|"&getTaggedValue(att,"SOSI_navn")&"")
		Session.Output("|"&att.Type&"")
		Session.Output("|"&getTaggedValue(att,"SOSI_lengde")&"")
		' new tags in the PS?
		Session.Output("|"&getTaggedValue(att,"FKB-A")&"")
		Session.Output("|"&getTaggedValue(att,"FKB-B")&"")
		Session.Output("|"&getTaggedValue(att,"FKB-C")&"")
		Session.Output("|"&getTaggedValue(att,"FKB-D")&"")
		Session.Output(" ")	
	next
	Session.Output("|===")
end if

If element.Connectors.Count > numberSpecializations Then
	Relasjoner(element)
End If
End sub
'-----------------ObjektOgDatatyper End-----------------


'-----------------CodeList-----------------
Sub Kodelister(element)
Dim att As EA.Attribute
dim tag as EA.TaggedValue
dim utvekslingsalias
Session.Output(" ")
Session.Output("==== «"&element.Stereotype&"» "&element.Name&"")
Session.Output("Definisjon: "&getCleanDefinition(element.Notes)&"")
Session.Output(" ")

if element.TaggedValues.Count > 0 then
	Session.Output("===== Tagged Values")
	Session.Output("[cols=""20,80""]")
	Session.Output("|===")
	for each tag in element.TaggedValues								
		if tag.Value <> "" then	
			if tag.Name <> "persistence" and tag.Name <> "SOSI_melding" then
			'	Session.Output("|Tag: "&tag.Name&"")
			'	Session.Output("|Verdi: "&tag.Value&"")
				Session.Output("|"&tag.Name&"")
				Session.Output("|"&tag.Value&"")
				Session.Output(" ")			
			end if	
		end if
	next
	Session.Output("|===")
		
	for each tag in element.TaggedValues								
		if tag.Name = "SOSI_bildeAvModellelement" and tag.Value <> "" then
			diagCounter = diagCounter + 1
			Session.Output("[caption=""Figur "&diagCounter&": "",title="&tag.Name&"]")
		'	Session.Output("image::"&tag.Value&".png["&ThePackage.Name"."&tag.Name&"]")
			Session.Output("image::"&tag.Value&"["&tag.Value&"]")
		end if
	next
end if
if element.Attributes.Count > 0 then
Session.Output("===== Koder")
end if
utvekslingsalias = false
for each att in element.Attributes
	if att.Default <> "" then
		utvekslingsalias = true
	end if
next
if element.Attributes.Count > 1 then
if utvekslingsalias then
	Session.Output("[cols=""15,25,60""]")
	Session.Output("|===")
	Session.Output("|*Utvekslingsalias:* ")
	Session.Output("|*Kodenavn:* ")
	Session.Output("|*Definisjon:* ")
	Session.Output(" ")
	for each att in element.Attributes
		if att.Default <> "" then
			Session.Output("|"&att.Default&"")
		else
			Session.Output("|")
		end if
		Session.Output("|"&att.Name&"")
		Session.Output("|"&att.Notes&"")
	next
	Session.Output("|===")
else
	Session.Output("[cols=""20,80""]")
	Session.Output("|===")
	Session.Output("|*Navn:* ")
	Session.Output("|*Definisjon:* ")
	Session.Output(" ")
	for each att in element.Attributes
		Session.Output("|"&att.Name&"")
		Session.Output("|"&att.Notes&"")
	next
	Session.Output("|===")
end if
else
for each att in element.Attributes
	Session.Output("[cols=""20,80""]")
	Session.Output("|===")
	Session.Output("|Navn: ")
	Session.Output("|"&att.name&"")
	Session.Output(" ")
	Session.Output("|Definisjon: ")
	Session.Output("|"&att.Notes&"")
	if not att.Default = "" then
		Session.Output(" ")
		Session.Output("|Utvekslingsalias?: ")
		Session.Output("|"&att.Default&"")
	end if
	Session.Output("|===")
next
end if
End sub
'-----------------CodeList End-----------------


'-----------------Relasjoner-----------------
sub Relasjoner(element)
Dim generalizations
Dim con
Dim supplier
Dim client
Dim textVar

Session.Output("===== Roller")


'assosiasjoner
For Each con In element.Connectors
	If con.Type = "Association" or con.Type = "Aggregation" Then
		Session.Output("[cols=""20,80""]")
		Session.Output("|===")
		set supplier = Repository.GetElementByID(con.SupplierID)
		set client = Repository.GetElementByID(con.ClientID)
	'	Session.Output("|Type: ")
	'	Session.Output("|Assosiasjon ")
	'	Session.Output(" ")
		If supplier.elementID = element.elementID Then 'dette elementet er suppliersiden - implisitt at fraklasse er denne klassen
			textVar="|Til klasse"
			If con.ClientEnd.Navigable = "Navigable" Then 'Legg til info om klassen er navigerbar eller spesifisert ikke-navigerbar.
			'	textVar=textVar+" _(navigerbar)_:"
			ElseIf con.ClientEnd.Navigable = "Non-Navigable" Then 
				textVar=textVar+" _(ikke navigerbar)_:"
			Else 
				textVar=textVar+":" 
			End If
		'	Session.Output(textVar)
		'	Session.Output("|«" & client.Stereotype&"» "&client.Name)
		'	Session.Output(" ")
			If con.ClientEnd.Role <> "" Then
				Session.Output("|*Rollenavn:* ")
				Session.Output("|*" & con.ClientEnd.Role & "*")
				Session.Output(" ")
			'End If
				If con.ClientEnd.RoleNote <> "" Then
					Session.Output("|Definisjon: ")
					Session.Output("|" & con.ClientEnd.RoleNote)
					Session.Output(" ")
				End If
				If con.ClientEnd.Cardinality <> "" Then
					Session.Output("|Multiplisitet: ")
					Session.Output("|[" & con.ClientEnd.Cardinality&"]")
					Session.Output(" ")
				End If
				Session.Output(textVar)
				Session.Output("|«" & client.Stereotype&"» "&client.Name)
				if false then
				If con.SupplierEnd.Role <> "" Then
					Session.Output("|Fra rolle: ")
					Session.Output("|" & con.SupplierEnd.Role)
					Session.Output(" ")
				End If
				If con.SupplierEnd.RoleNote <> "" Then
					Session.Output("|Fra rolle definisjon: ")
					Session.Output("|" & con.SupplierEnd.RoleNote)
					Session.Output(" ")
				End If
				If con.SupplierEnd.Cardinality <> "" Then
					Session.Output("|Fra multiplisitet: ")
					Session.Output("|[" & con.SupplierEnd.Cardinality&"]")
					Session.Output(" ")
				End If
			End If
			end if
		Else 'dette elementet er clientsiden, (rollen er på target)
			textVar="|Til klasse"
			If con.SupplierEnd.Navigable = "Navigable" Then
			'	textVar=textVar+" _(navigerbar)_:"
			ElseIf con.SupplierEnd.Navigable = "Non-Navigable" Then
				textVar=textVar+" _(ikke-navigerbar)_:"
			Else
				textVar=textVar+":"
			End If
		'	Session.Output(textVar)
		'	Session.Output("|«" & supplier.Stereotype&"» "&supplier.Name)
			If con.SupplierEnd.Role <> "" Then
				Session.Output("|*Rollenavn:* ")
				Session.Output("|*" & con.SupplierEnd.Role & "*")
				Session.Output(" ")
			'	End If
				If con.SupplierEnd.RoleNote <> "" Then
					Session.Output("|Definisjon:")
					Session.Output("|" & con.SupplierEnd.RoleNote)
					Session.Output(" ")
				End If
				If con.SupplierEnd.Cardinality <> "" Then
					Session.Output("|Multiplisitet: ")
					Session.Output("|[" & con.SupplierEnd.Cardinality&"]")
					Session.Output(" ")
				End If
				Session.Output(textVar)
				Session.Output("|«" & supplier.Stereotype&"» "&supplier.Name)
				if false then
				If con.ClientEnd.Role <> "" Then
					Session.Output("|Fra rolle: ")
					Session.Output("|" & con.ClientEnd.Role)
					Session.Output(" ")
				End If
				If con.ClientEnd.RoleNote <> "" Then
					Session.Output("|Fra rolle definisjon: ")
					Session.Output("|" & con.ClientEnd.RoleNote)
					Session.Output(" ")
				End If
				If con.ClientEnd.Cardinality <> "" Then
					Session.Output("|Fra multiplisitet: ")
					Session.Output("|[" & con.ClientEnd.Cardinality&"]")
					Session.Output(" ")
				End If
			End If
			end if
		End If
		Session.Output("|===")
	End If
Next
if false then
'aggregeringer
For Each con In element.Connectors
	If con.Type = "Aggregation" Then
		Session.Output("[cols=""20,80""]")
		Session.Output("|===")
		set supplier = Repository.GetElementByID(con.SupplierID)
		set client = Repository.GetElementByID(con.ClientID)
		Session.Output("|Type: ")
		If con.clientend.aggregation = 1 Or con.supplierend.aggregation = 1 Then
			Session.Output("|Aggregering")
		ElseIf con.clientend.aggregation = 2 Or con.supplierend.aggregation = 2 Then
			Session.Output("|Komposisjon")
		End If
		Session.Output(" ")
		If supplier.elementID = element.elementID Then 'dette elementet er suppliersiden - implisitt at fraklasse er denne klassen
			textVar="|Til klasse"
			If con.clientend.aggregation = 0 Then 'motsatt side er komponent i denne klassen
				textVar=textVar+" _(del"
			Else
				textVar=textVar+" _(helhet"
			End If
			If con.ClientEnd.Navigable = "Navigable" Then 'Legg til info om klassen er navigerbar eller spesifisert ikke-navigerbar.
				textVar=textVar+", navigerbar)_:"
			ElseIf con.ClientEnd.Navigable = "Non-Navigable" Then 
				textVar=textVar+", ikke navigerbar)_:"
			Else 
				textVar=textVar+")_:" 
			End If
			Session.Output(textVar)
			Session.Output("|«" & client.Stereotype&"» "&client.Name)
			Session.Output(" ")
			If con.ClientEnd.Role <> "" Then
				Session.Output("|Til rolle: ")
				Session.Output("|" & con.ClientEnd.Role)
				Session.Output(" ")
			End If
			If con.ClientEnd.RoleNote <> "" Then
				Session.Output("|Til rolle definisjon: ")
				Session.Output("|" & con.ClientEnd.RoleNote)
				Session.Output(" ")
			End If
			If con.ClientEnd.Cardinality <> "" Then
				Session.Output("|Til multiplisitet: ")
				Session.Output("|[" & con.ClientEnd.Cardinality&"]")
				Session.Output(" ")
			End If
			If con.SupplierEnd.Role <> "" Then
				Session.Output("|Fra rolle: ")
				Session.Output("|" & con.SupplierEnd.Role)
				Session.Output(" ")
			End If
			If con.SupplierEnd.RoleNote <> "" Then
				Session.Output("|Fra rolle definisjon: ")
				Session.Output("|" & con.SupplierEnd.RoleNote)
				Session.Output(" ")
			End If
			If con.SupplierEnd.Cardinality <> "" Then
				Session.Output("|Fra multiplisitet: ")
				Session.Output("|[" & con.SupplierEnd.Cardinality&"]")
				Session.Output(" ")
			End If
		Else 'dette elementet er clientsiden
			textVar="|Til klasse"
			If con.supplierEnd.aggregation = 0 Then 'motsatt side er komponent i denne klassen
				textVar=textVar+" _(del"
			Else
				textVar=textVar+" _(helhet"
			End If
			If con.SupplierEnd.Navigable = "Navigable" Then 'Legg til info om klassen er navigerbar eller spesifisert ikke-navigerbar.
				textVar=textVar+", navigerbar)_:"
			ElseIf con.SupplierEnd.Navigable = "Non-Navigable" Then 
				textVar=textVar+", ikke navigerbar)_:"
			Else 
				textVar=textVar+")_:" 
			End If
			Session.Output(textVar)
			Session.Output("|«" & supplier.Stereotype&"» "&supplier.Name)
			If con.SupplierEnd.Role <> "" Then
				Session.Output("|Til rolle: ")
				Session.Output("|" & con.SupplierEnd.Role)
				Session.Output(" ")
			End If
			If con.SupplierEnd.RoleNote <> "" Then
				Session.Output("|Til rolle definisjon: ")
				Session.Output("|" & con.SupplierEnd.RoleNote)
				Session.Output(" ")
			End If
			If con.SupplierEnd.Cardinality <> "" Then
				Session.Output("|Til multiplisitet: ")
				Session.Output("|[" & con.SupplierEnd.Cardinality&"]")
				Session.Output(" ")
			End If
			If con.ClientEnd.Role <> "" Then
				Session.Output("|Fra rolle: ")
				Session.Output("|" & con.ClientEnd.Role)
				Session.Output(" ")
			End If
			If con.ClientEnd.RoleNote <> "" Then
				Session.Output("|Fra rolle definisjon: ")
				Session.Output("|" & con.ClientEnd.RoleNote)
				Session.Output(" ")
			End If
			If con.ClientEnd.Cardinality <> "" Then
				Session.Output("|Fra multiplisitet: ")
				Session.Output("|[" & con.ClientEnd.Cardinality&"]")
				Session.Output(" ")
			End If
		End If
		Session.Output("|===")
	End If
Next
end if

' Generaliseringer av pakken
generalizations = False
For Each con In element.Connectors
	If con.Type = "Generalization" Then
		set supplier = Repository.GetElementByID(con.SupplierID)
		set client = Repository.GetElementByID(con.ClientID)
		If supplier.ElementID=element.ElementID then 'dette er en generalisering
			If Not generalizations Then
				Session.Output("[cols=""20,80""]")
				Session.Output("|===")
				Session.Output("|Generalisering av:")
				textVar = "|«" + client.Stereotype + "» " + client.Name
				generalizations = True
			Else
				textVar = textVar + " +" + vbLF + "«" + client.Stereotype + "» " + client.Name
			End If
		End If
	End If
Next
If generalizations then
	Session.Output(textVar)
	Session.Output("|===")
End If

end sub
'-----------------Relasjoner End-----------------

'-----------------Funksjon for full path-----------------
function getPath(package)
	dim path
	dim parent
	if package.Element.Stereotype = "" then
		path = package.Name
	else
		path = "«" + package.Element.Stereotype + "» " + package.Name
	end if
	if not (ucase(package.Element.Stereotype)="APPLICATIONSCHEMA" or package.parentID = 0) then
		set parent = Repository.GetPackageByID(package.ParentID)
		path = getPath(parent) + "/" + path
	end if
	getPath = path
end function
'-----------------Funksjon for full path End-----------------


'-----------------Function getCleanDefinition Start-----------------
function getCleanDefinition(txt)
	'removes all formatting in notes fields, except crlf
    Dim res, tegn, i, u
    u=0
	getCleanDefinition = ""

		res = ""
		' loop gjennom alle tegn
		For i = 1 To Len(txt)
		  tegn = Mid(txt,i,1)
		  If tegn = "<" Then
				u = 1
			   'res = res + " "
		  Else 
			If tegn = ">" Then
				u = 0
			   'res = res + " "
				'If tegn = """" Then
				'  res = res + "'"
			Else
				  If tegn < " " and Asc(tegn) <> 10 and Asc(tegn) <> 13 Then
					res = res + " "
				  Else
					if u = 0 then
						res = res + Mid(txt,i,1)
					end if
				  End If
				'End If
			End If
		  End If
		  
		Next
		
	getCleanDefinition = res

end function
'-----------------Function getCleanDefinition End-----------------


function getTaggedValue(element,taggedValueName)
		dim i, existingTaggedValue
		getTaggedValue = ""
		for i = 0 to element.TaggedValues.Count - 1
			set existingTaggedValue = element.TaggedValues.GetAt(i)
			if existingTaggedValue.Name = taggedValueName then
				getTaggedValue = existingTaggedValue.Value
			end if
		next
end function

function getPackageTaggedValue(package,taggedValueName)
		dim i, existingTaggedValue
		getPackageTaggedValue = ""
		for i = 0 to package.element.TaggedValues.Count - 1
			set existingTaggedValue = package.element.TaggedValues.GetAt(i)
			if existingTaggedValue.Name = taggedValueName then
				getPackageTaggedValue = existingTaggedValue.Value
			end if
		next
end function

function getConnectorEndTaggedValue(connectorEnd,taggedValueName)
	getConnectorEndTaggedValue = ""
	if not connectorEnd is nothing and Len(taggedValueName) > 0 then
		dim existingTaggedValue as EA.RoleTag 
		dim i
		for i = 0 to connectorEnd.TaggedValues.Count - 1
			set existingTaggedValue = connectorEnd.TaggedValues.GetAt(i)
			if existingTaggedValue.Tag = taggedValueName then
				getConnectorEndTaggedValue = existingTaggedValue.Value
			end if 
		next
	end if 
end function 


OnProjectBrowserScript
