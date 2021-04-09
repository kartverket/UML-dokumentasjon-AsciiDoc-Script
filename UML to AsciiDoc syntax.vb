Option Explicit

!INC Local Scripts.EAConstants-VBScript

' Script Name: AsciiDocTest3
' Author: Tore Johnsen, Åsmund Tjora
' Purpose: Generate documentation in AsciiDoc syntax
' Date: 08.04.2021


' Project Browser Script main function
Sub OnProjectBrowserScript()

    Dim treeSelectedType
    treeSelectedType = Repository.GetTreeSelectedItemType()

    Select Case treeSelectedType

        Case otPackage
            ' Code for when a package is selected
            Dim thePackage As EA.Package
			set thePackage = Repository.GetTreeSelectedObject()
			Call ListAsciiDoc(thePackage)

        Case Else
            ' Error message
            Session.Prompt "This script does not support items of this type.", promptOK

    End Select

End Sub


Sub ListAsciiDoc(thePackage)
Dim element As EA.Element
Dim diag As EA.Diagram
Dim projectclass As EA.Project
set projectclass = Repository.GetProjectInterface()
Dim diagCounter
diagCounter = 0

Session.Output("=== "&thePackage.Name&"")
Session.Output("Definisjon: "&thePackage.Notes&"")

'-----------------Diagram-----------------
For Each diag In thePackage.Diagrams
	diagCounter = diagCounter + 1
	Call projectclass.PutDiagramImageToFile(diag.DiagramGUID, "" & diag.Name&".png", 1)
	Repository.CloseDiagram(diag.DiagramID)
	Session.Output("[caption=""Figur "&diagCounter&": "",title="&diag.Name&"]")
	Session.Output("image::"&diag.Name&".png["&diag.Name&"]")
Next

For each element in thePackage.Elements
	If Ucase(element.Stereotype) = "FEATURETYPE" Then
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
Session.Output("Definisjon: "&element.Notes&"")
Session.Output(" ")
numberSpecializations = 0
For Each con In element.Connectors
	set supplier = Repository.GetElementByID(con.SupplierID)
	If con.Type = "Generalization" And supplier.ElementID <> element.ElementID Then
		Session.Output("Spesialisering av «" & supplier.Stereotype&"» "&supplier.Name&"")
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
		Session.Output("Realisering av " & textVar &"::«" & supplier.Stereotype&"» "&supplier.Name)
		Session.Output(" ")
		numberSpecializations = numberSpecializations + 1
	end if
next

if element.Attributes.Count > 0 then
	Session.Output("===== Egenskaper")
	for each att in element.Attributes
		Session.Output("[cols=""20,80""]")
		Session.Output("|===")
		Session.Output("|*Navn:* ")
		Session.Output("|*"&att.name&"*")
		Session.Output(" ")
		Session.Output("|Definisjon: ")
		Session.Output("|"&att.Notes&"")
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
		Session.Output("|"&att.Type&"")			

		if att.TaggedValues.Count > 0 then
			Session.Output("|Tagged Values: ")
			Session.Output("|")
			for each tag in att.TaggedValues
				Session.Output(""&tag.Name& ": "&tag.Value&" + ")
			next
		end if
		Session.Output("|===")
	next
end if

if element.TaggedValues.Count > 0 then
	Session.Output("===== Tagged Values")
	Session.Output("[cols=""20,80""]")
	Session.Output("|===")
	for each tag in element.TaggedValues								
		Session.Output("|Tag: "&tag.Name&"")
		Session.Output("|Verdi: "&tag.Value&"")
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
Session.Output(" ")
Session.Output("==== «"&element.Stereotype&"» "&element.Name&"")
Session.Output("Definisjon: "&element.Notes&"")
Session.Output(" ")
Session.Output("===== Koder")
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
End sub
'-----------------CodeList End-----------------


'-----------------Relasjoner-----------------
sub Relasjoner(element)
Dim generalizations
Dim con
Dim supplier
Dim client
Dim textVar

Session.Output("===== Relasjoner")

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

'assosiasjoner
For Each con In element.Connectors
	If con.Type = "Association" Then
		Session.Output("[cols=""20,80""]")
		Session.Output("|===")
		set supplier = Repository.GetElementByID(con.SupplierID)
		set client = Repository.GetElementByID(con.ClientID)
		Session.Output("|Type: ")
		Session.Output("|Assosiasjon ")
		Session.Output(" ")
		If supplier.elementID = element.elementID Then 'dette elementet er suppliersiden - implisitt at fraklasse er denne klassen
			textVar="|Til klasse"
			If con.ClientEnd.Navigable = "Navigable" Then 'Legg til info om klassen er navigerbar eller spesifisert ikke-navigerbar.
				textVar=textVar+" _(navigerbar)_:"
			ElseIf con.ClientEnd.Navigable = "Non-Navigable" Then 
				textVar=textVar+" _(ikke navigerbar)_:"
			Else 
				textVar=textVar+":" 
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
			If con.SupplierEnd.Navigable = "Navigable" Then
				textVar=textVar+" _(navigerbar)_:"
			ElseIf con.SupplierEnd.Navigable = "Non-Navigable" Then
				textVar=textVar+" _(ikke-navigerbar)_:"
			Else
				textVar=textVar+":"
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

OnProjectBrowserScript