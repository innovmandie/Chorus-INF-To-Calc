REM  *****  BASIC  *****

Option Explicit

' =====================================================================
' 
' INF2Calc est une macro LibreOffice écrite en Python, compatible avec l'interpréteur 3.11.14 (64 bit) intégré à LibreOffice 25.8.4.2 (X86_64) 
' permettant l'ouverture dans Calc des restitutions Chorus INF exportées en Excel 2000 (faux fichier .xls, en réalité au format « Microsoft SpreadsheetML HTML »),
' après conversion automatique en fichier Excel (.xlsx) exploitable nativement par Calc.
' 
' Ce programme oooBasic installe ou désinstalle INF2Calc et la commande de menu
'
' Les procédures Install_INF2Calc et Remove_INF2Calc sont assignées aux boutons correspondants dans le corps du document ODT
'
' Il copie l'arborescence suivante dans "Mes Macros et boîtes de dialogue" (%USERPROFILE%\AppData\Roaming\LibreOffice\4\user\) : 
' Scripts/
' └── python/
'     ├── INF2Calc.py
'     └── pythonpath/
'         ├── openpyxl/
'         │   ├── __init__.py
'         │   └── ...
'         └── et_xmlfile/
'             ├── __init__.py
'             └── ...
' Il ajoute ou supprime la commande "Ouvrir Chorus INF" sous "Ouvir distant...." dans le menu "Fichier" de LibreOffice Calc
' La commande lance la macro python INF2Calc.py$INF2Calc
'
' La macro  "EmbeddingFromProfileToThisDocument" ne sert qu'à embarquer les script python du user profile vers le fichier ODT
' 
' =====================================================================


' ===============================================
'  INF2Calc - Installation / désinstallation via le document (ODT) en OooBasic
' ===============================================


' ====================
'     Bouton Installer INF2Calc
' ====================

Sub Install_INF2Calc()
On Error Goto Err_Install_INF2Calc

    Dim oSFA As Object
    Dim oRestart As Object
    Dim oCfg As Object
    Dim oSettings As Object
    Dim oConteneur As Object
    Dim iIndex As Integer
    Dim aItem(2) As New com.sun.star.beans.PropertyValue
    Dim sAddr As String
    Dim sUrlDoc As String
   
	If MsgBox("Voulez-vous installer INF2Calc ?" & Chr(13) & _
              "(Ceci remplacera toute installation INF2Calc existante)", _
              MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2, "Installation de INF2Calc") = IDYES Then 

		' Suppression des fichiers  d'une éventuelle ancienne installation

		oSFA = createUnoService("com.sun.star.ucb.SimpleFileAccess")
			
	    sAddr = UserProfile() & "/Scripts/python/INF2Calc.py"
	    If oSFA.exists(sAddr) Then
	    	oSFA.kill(sAddr)
		End If	
	   sAddr = UserProfile() & "/Scripts/python/pythonpath/openpyxl"
	    If oSFA.exists(sAddr) Then
	    	oSFA.kill(sAddr)
		End If	
	    sAddr = UserProfile() & "/Scripts/python/pythonpath/et_xmlfile"
	    If oSFA.exists(sAddr) Then
	    	oSFA.kill(sAddr)
		End If	
	    sAddr = UserProfile() & "/Scripts/python/pythonpath/__pycache__"
	    If oSFA.exists(sAddr) Then
	    	oSFA.kill(sAddr)
		End If	
	    sAddr = UserProfile() & "/Scripts/python/__pycache__"
	    If oSFA.exists(sAddr) Then
	    	oSFA.kill(sAddr)
		End If

		 sUrlDoc = UrlDoc()

 		' Copie des fichiers ThisDocument -> UserProfile
		CopyDirectory(sUrlDoc & "/Scripts/python/pythonpath/openpyxl",  UserProfile()  & "/Scripts/python/pythonpath/openpyxl")
		CopyDirectory(sUrlDoc & "/Scripts/python/pythonpath/et_xmlfile",  UserProfile()  & "/Scripts/python/pythonpath/et_xmlfile")
		
		CopyFile(sUrlDoc & "/Scripts/python/INF2Calc.py", UserProfile()  & "/Scripts/python")
	     
	    ' Ajout de la commande  "Ouvrir Chorus INF" dans le menu "Fichier"
	    oCfg = createUnoService("com.sun.star.ui.ModuleUIConfigurationManagerSupplier").getUIConfigurationManager("com.sun.star.sheet.SpreadsheetDocument")
	    oSettings = oCfg.getSettings("private:resource/menubar/menubar", True)
	
	    If Not FindCommandMenu(oSettings, "vnd.sun.star.script:INF2Calc.py$INF2Calc?language=Python&location=user", oConteneur, iIndex) Then
	    	'   "Ouvrir Chorus INF" n'est pas déjà présent dans le menu "Fichier"
	
		    If FindCommandMenu(oSettings, ".uno:OpenRemote", oConteneur, iIndex) Then
		    	
		    	'  "Ouvrir distant..." est  présent dans le menu "Fichier", on peut installer en-dessous
			    aItem(0).Name = "CommandURL"
			    aItem(0).Value = "vnd.sun.star.script:INF2Calc.py$INF2Calc?language=Python&location=user"
			    aItem(1).Name = "Label"
			    aItem(1).Value = "Ouvrir Chorus INF"
			    aItem(2).Name = "Type"
			    aItem(2).Value = 0
			
			    oConteneur.insertByIndex(iIndex + 1, aItem)
			    oCfg.replaceSettings("private:resource/menubar/menubar", oSettings)
			    oCfg.store()	

				' Redémarrage de LibreOffice
				If MsgBox("LibreOffice doit être redémarré pour que les modifications soient prises en compte." & Chr(13) & _
						    "Voulez-vous redémarrer LibreOffice maintenant ?",  _
						    MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1, _
						    "Installation de INF2Calc") = IDYES Then
					oRestart = GetDefaultContext().getValueByName("/singletons/com.sun.star.task.OfficeRestartManager")
					oRestart.requestRestart(Nothing)
				End If
					    	    
		    Else
		    
		        MsgBox "'Ouvrir distant...' introuvable dans le menu Fichier de Calc.", MB_ICONSTOP, "Installation de INF2Calc"
		    
		    End If
		    
		End If
	
	End If

Exit_Install_INF2Calc:
    Exit Sub

Err_Install_INF2Calc:
    MsgBox "Erreur " & Err & " : " & Error$ & " (ligne " & Erl & ")",  MB_ICONSTOP, "Installation de INF2Calc"
    Resume Exit_Install_INF2Calc

End Sub

' ====================
'     Bouton Désinstaller INF2Calc
' ====================

Sub Remove_INF2Calc()
On Error Goto Err_Remove_INF2Calc

    Dim oSFA As Object
    Dim oCfg As Object
    Dim oSettings As Object
    Dim oConteneur As Object
    Dim iIndex As Integer
    Dim sAddr As String
    
	If MsgBox("Voulez-vous désinstaller INF2Calc ?", _
              			MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2, _
              			"Désinstallation de INF2Calc") = IDYES Then 	    

		' Suppression des fichiers

		oSFA = createUnoService("com.sun.star.ucb.SimpleFileAccess")
			
	    sAddr = UserProfile() & "/Scripts/python/INF2Calc.py"
	    If oSFA.exists(sAddr) Then
	    	oSFA.kill(sAddr)
		End If	
	   sAddr = UserProfile() & "/Scripts/python/pythonpath/openpyxl"
	    If oSFA.exists(sAddr) Then
	    	oSFA.kill(sAddr)
		End If	
	    sAddr = UserProfile() & "/Scripts/python/pythonpath/et_xmlfile"
	    If oSFA.exists(sAddr) Then
	    	oSFA.kill(sAddr)
		End If	
	    sAddr = UserProfile() & "/Scripts/python/pythonpath/__pycache__"
	    If oSFA.exists(sAddr) Then
	    	oSFA.kill(sAddr)
		End If	
	    sAddr = UserProfile() & "/Scripts/python/__pycache__"
	    If oSFA.exists(sAddr) Then
	    	oSFA.kill(sAddr)
		End If

	    ' Suppression de la commande  "Ouvrir Chorus INF" du menu "Fichier"
	    oCfg = createUnoService("com.sun.star.ui.ModuleUIConfigurationManagerSupplier").getUIConfigurationManager("com.sun.star.sheet.SpreadsheetDocument")
	    oSettings = oCfg.getSettings("private:resource/menubar/menubar", True)

    	If FindCommandMenu(oSettings, "vnd.sun.star.script:INF2Calc.py$INF2Calc?language=Python&location=user", oConteneur, iIndex) Then 
		    oConteneur.removeByIndex(iIndex)
		    oCfg.replaceSettings("private:resource/menubar/menubar", oSettings)
		    oCfg.store()
		End If    
		
		' Redémarrage de LibreOffice non forcé    
	    MsgBox "La désinstallation sera effective au prochain redémarrage de LibreOffice.", MB_ICONINFORMATION, "Désinstallation de INF2Calc"
	
	End If
	
Exit_Remove_INF2Calc:
    Exit Sub

Err_Remove_INF2Calc:
    MsgBox "Erreur " & Err & " : " & Error$ & " (ligne " & Erl & ")",  MB_ICONSTOP, "Désinstallation de INF2Calc"
    Resume Exit_Remove_INF2Calc

End Sub

' ======================================================
'    Embarquement du code à installer dans le fichier ODT (usage unique pour déploiement)
'    Les fichiers et dossiers inutiles au déploiement sont filtrés
'    (voir IsExcludedDirectory / IsExcludedFile)
' ======================================================

Sub EmbeddingFromProfileToThisDocument()
On Error Goto Err_EmbeddingFromProfileToThisDocument

	' Copie FILTRÉE des fichiers UserProfile -> ThisDocument
	CopyDirectoryFiltered(UserProfile() & "/Scripts/python/pythonpath/openpyxl",  UrlDoc() & "/Scripts/python/pythonpath/openpyxl")
	CopyDirectoryFiltered(UserProfile() & "/Scripts/python/pythonpath/et_xmlfile",  UrlDoc() & "/Scripts/python/pythonpath/et_xmlfile")

	CopyFile(UserProfile() & "/Scripts/python/INF2Calc.py", UrlDoc() & "/Scripts/python")

    ThisComponent.setModified(True)

    If ThisComponent.hasLocation() And Not ThisComponent.isReadonly() Then
        ThisComponent.store()
        MsgBox "INF2Calc est embarqué et le document est enregistré.", _
	               MB_ICONINFORMATION, _
	               "Embarquement de INF2Calc"
    Else
        MsgBox "INF2Calc est embarqué." & Chr(13) & _
		               "Enregistrez le document pour conserver l'ajout (document jamais enregistré ou en lecture seule).", _
		               MB_ICONEXCLAMATION, _
		               "Embarquement de INF2Calc"
    End If

Exit_EmbeddingFromProfileToThisDocument:
    Exit Sub

Err_EmbeddingFromProfileToThisDocument:
    MsgBox "Erreur " & Err & " : " & Error$ & " (ligne " & Erl & ")",  MB_ICONSTOP, "Embarquement de INF2Calc"
    Resume Exit_EmbeddingFromProfileToThisDocument

End Sub


' Copie récursive d'UN dossier complet AVEC filtrage des éléments
' inutiles au déploiement. La destination est d'abord supprimée :
' vrai remplacement, sans reliquat.
Private Sub CopyDirectoryFiltered(sSourceDirectory As String, sTargetDirectory As String)

    Dim oSFA As Object
    Dim aContenu() As String
    Dim i As Integer
    Dim sNom As String

    oSFA = createUnoService("com.sun.star.ucb.SimpleFileAccess")

    If oSFA.exists(sTargetDirectory) Then
    	oSFA.kill(sTargetDirectory)
    End If
    CreerArborescence(sTargetDirectory)

    aContenu = oSFA.getFolderContents(sSourceDirectory, True)
    For i = LBound(aContenu) To UBound(aContenu)
        sNom = FileNameoutofPath(aContenu(i), "/")
        If oSFA.isFolder(aContenu(i)) Then
            If Not IsExcludedDirectory(sNom) Then
               CopyDirectoryFiltered(aContenu(i), sTargetDirectory & "/" & sNom)
            End If
        Else
            If Not IsExcludedFile(sNom) Then
               CopyFile(aContenu(i), sTargetDirectory)
            End If
        End If
    Next i

End Sub


' Dossiers à ne pas embarquer dans l'ODT
Private Function IsExcludedDirectory(sNom As String) As Boolean

    Dim sN As String
    sN = LCase(sNom)

    IsExcludedDirectory = ( _
        sN = "__pycache__" Or _
        sN = ".git" Or sN = ".hg" Or sN = ".svn" Or _
        sN = ".mypy_cache" Or sN = ".pytest_cache" Or _
        sN = ".idea" Or sN = ".vscode" )

End Function


' Fichiers à ne pas embarquer dans l'ODT
Private Function IsExcludedFile(sNom As String) As Boolean

    Dim sN As String
    sN = LCase(sNom)

    ' Par extension
    If Right(sN, 4) = ".pyc" Or Right(sN, 4) = ".pyo" _
        Or Right(sN, 4) = ".pyi" _
        Or Right(sN, 4) = ".bak" Or Right(sN, 4) = ".tmp" _
        Or Right(sN, 4) = ".log" Or Right(sN, 5) = ".orig" _
        Or Right(sN, 1) = "~" Then
        IsExcludedFile = True
        Exit Function
    End If

    ' Par nom (déchets des explorateurs de fichiers)
    IsExcludedFile = ( _
        sN = "thumbs.db" Or sN = "desktop.ini" Or sN = ".ds_store" )

End Function


Sub CopyFile(sSourceFile As String, sTargetDirectory As String)
	' sSourceFile = Chemin complet du FICHIER source 
	' sTargetDirectory = Chemin complet du DOSSIER destination 
	' Le nom de fichier destination est extrait de sSourceFile
	' Copie un seul fichier du UserProfile (en URL file:///) vers ThisDocument (en URL vnd.sun.star.tdoc:/) ou inversement
	' Crée l'arborescence
	' Remplace si existant

    Dim oSFA As Object
    Dim oIn As Object
    Dim sNom As String
    
    oSFA = createUnoService("com.sun.star.ucb.SimpleFileAccess")

	sNom = FileNameoutofPath(sSourceFile, "/")
     CreerArborescence(sTargetDirectory)

    oIn = oSFA.openFileRead(sSourceFile)
    oSFA.writeFile(sTargetDirectory & "/" & sNom, oIn)   ' crée OU remplace
    oIn.closeInput()
    
End Sub


' Copie récursive d'UN dossier complet (*.*, toutes extensions).
' La destination est d'abord supprimée : vrai remplacement, sans reliquat.
Sub CopyDirectory(sSourceDirectory As String, sTargetDirectory As String)

    Dim oSFA As Object
    Dim aContenu() As String
    Dim i As Integer
    Dim sNom As String
 
    oSFA = createUnoService("com.sun.star.ucb.SimpleFileAccess")

    If oSFA.exists(sTargetDirectory) Then 
    	oSFA.kill(sTargetDirectory)
    End If
    CreerArborescence(sTargetDirectory)

    aContenu = oSFA.getFolderContents(sSourceDirectory, True)
    For i = LBound(aContenu) To UBound(aContenu)
        sNom = FileNameoutofPath(aContenu(i), "/")
        If oSFA.isFolder(aContenu(i)) Then
            If sNom <> "__pycache__" Then      ' cache Python : inutile à déployer
               CopyDirectory(aContenu(i), sTargetDirectory & "/" & sNom)
            End If
        Else
          CopyFile(aContenu(i), sTargetDirectory)
        End If
    Next i
    
End Sub


' Crée l'arborescence niveau par niveau (file:// et tdoc:/)
Private Sub CreerArborescence(sUrlDossier As String)

    Dim oSFA As Object
    Dim sParent As String
    
    If Trim(sUrlDossier) = "" Then
        Error 76      ' Chemin introuvable : URL vide (racine mal construite ?)
    End If

    oSFA = createUnoService("com.sun.star.ucb.SimpleFileAccess")

    If Not oSFA.exists(sUrlDossier) Then
 	    sParent = DirectoryNameoutofPath(sUrlDossier, "/")
	    If sParent = "" Or sParent = sUrlDossier Then
	        Error 76                               ' Chemin introuvable : racine inaccessible
	    End If
	    
	    If Not oSFA.exists(sParent) Then 
	    	CreerArborescence(sParent)
	    End If	
	    oSFA.createFolder(sUrlDossier)
	    
    End If
    
End Sub


' =====================================================================
'   UTILITIES
' =====================================================================

' Recherche récursive d'une commande dans la barre de menus.
' Renvoie True et, par référence, le conteneur et l'index de l'élément.
Private Function FindCommandMenu(oConteneur As Object, sCmd As String, _
        ByRef oTrouve As Object, ByRef iTrouve As Integer) As Boolean
    Dim i As Integer, aItem As Variant, vSous As Variant
    FindCommandMenu = False
    For i = 0 To oConteneur.Count - 1
        aItem = oConteneur.getByIndex(i)
        If ProprieteItem(aItem, "CommandURL") = sCmd Then
            oTrouve = oConteneur
            iTrouve = i
            FindCommandMenu = True
            Exit Function
        End If
        vSous = ProprieteItem(aItem, "ItemDescriptorContainer")
        If Not IsEmpty(vSous) And Not IsNull(vSous) Then
            If FindCommandMenu(vSous, sCmd, oTrouve, iTrouve) Then
                FindCommandMenu = True
                Exit Function
            End If
        End If
    Next i
End Function


' Valeur d'une propriété dans un descripteur d'élément de menu
Private Function ProprieteItem(aItem As Variant, sNom As String) As Variant
    Dim i As Integer
    ProprieteItem = Empty
    For i = LBound(aItem) To UBound(aItem)
        If aItem(i).Name = sNom Then
            ProprieteItem = aItem(i).Value
            Exit Function
        End If
    Next i
End Function

' Racine du PROFIL, en URL file:///
Private Function UserProfile() As String
    Dim oPS As Object
    oPS = createUnoService("com.sun.star.util.PathSubstitution")
    UserProfile = oPS.getSubstituteVariableValue("$(user)")
End Function


' Racine du DOCUMENT, en URL vnd.sun.star.tdoc:/
Private Function UrlDoc() As String
    Dim oTdf As Object
    Dim sUrl As String
    oTdf = createUnoService("com.sun.star.frame.TransientDocumentsDocumentContentFactory")
    sUrl = oTdf.createDocumentContent(ThisComponent).getIdentifier().getContentIdentifier()
    If Right(sUrl, 1) = "/" Then
    	sUrl = Left(sUrl, Len(sUrl) - 1)
    End If
    UrlDoc = sUrl
End Function


' =====================================================================
'     FONCTIONS REPRISES DE LA BIBLIOTHÈQUE "Tools", MODULE "Strings",
'     livrée avec LibreOffice (Macros de l'application), licence MPL 2.0.
'     Intégrées telles quelles pour ne pas dépendre de
'     GlobalScope.BasicLibraries.loadLibrary("Tools") sur les postes.
' =====================================================================

' Origine : Tools/Strings - ArrayoutofString
' Retrieves an Array out of a String.
' The fields of the Array are separated by the parameter 'Separator', that is contained
' in the Array
' The Array MaxIndex delivers the highest Index of this Array
Function ArrayOutOfString(BigString, Separator as String, Optional MaxIndex as Integer)
	
	Dim LocList() as String
	
	LocList=Split(BigString,Separator)

	If not isMissing(MaxIndex) then 
		maxIndex=ubound(LocList())	
	End If
	
	ArrayOutOfString=LocList

End Function

' Origine : Tools/Strings - RTrimStr
' Deletes the String 'SmallString' out of the String 'BigString'
' in case SmallString's Position in BigString is right at the end
Function RTrimStr(ByVal BigString, SmallString as String) as String

	Dim SmallLen as Integer
	Dim BigLen as Integer
	
	SmallLen = Len(SmallString)
	BigLen = Len(BigString)
	If Instr(1,BigString, SmallString) <> 0 Then
		If Mid(BigString,BigLen + 1 - SmallLen, SmallLen) = SmallString Then
			RTrimStr = Mid(BigString,1,BigLen - SmallLen)
		Else
			RTrimStr = BigString
		End If
	Else
		RTrimStr = BigString
	End If

End Function

' Origine : Tools/Strings - FileNameoutofPath
'Retrieves the mere filename out of a whole path
Function FileNameoutofPath(ByVal Path as String, Optional Separator as String) as String
	
	Dim i as Integer
	Dim SepList() as String
	
	If IsMissing(Separator) Then
		Path = ConvertFromUrl(Path)
		Separator = GetPathSeparator()		
	End If
	SepList() = ArrayoutofString(Path, Separator,i)
	FileNameoutofPath = SepList(i)

End Function

' Origine : Tools/Strings - DirectoryNameoutofPath
Function DirectoryNameoutofPath(sPath as String, Separator as String) as String

	Dim LocFileName as String
	
	LocFileName = FileNameoutofPath(sPath, Separator)
	DirectoryNameoutofPath = RTrimStr(sPath, Separator & LocFileName)

End Function

