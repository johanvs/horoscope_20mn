Option Compare Database
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub download_all_20mn()

    Dim curDate As Date
    Dim url, cmd As String
    
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1


    'On initialise la date à aujourd'hui
    curDate = Date
    ChDir Application.CurrentProject.Path
    
    For i = 0 To 3000
    
        url = "http://pdf.20mn.fr/" & Format(curDate, "yyyy") & "/quotidien/" & Format(curDate, "yyyymmdd") & "_PAR.pdf"
        cmd = "aria2c64.exe -d 20mn " & url
        
        'Shell (cmd)
        wsh.Run cmd, windowStyle, waitOnReturn
        
        curDate = DateAdd("d", -1, curDate)
    
    Next

End Sub


Sub pdf_to_txt()

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(Application.CurrentProject.Path & "\20mn")
    Set fc = f.Files
    
    ChDir Application.CurrentProject.Path
    
    For Each f1 In fc
    
        If Dir("txt/" & f1.Name & ".txt") = "" Then
		
			'Ouverture du document
			
			'Remplacer par le chemin d'adobe reader
			Dim READERPath As String
			READERPath = """C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"" "
			
			Shell READERPath & "20mn\" & f1.Name, vbNormalFocus: DoEvents
			Sleep 2000
		
			'récupération du contenu dans le presse papier
			SendKeys "^a", True
			Sleep 2000
		
			SendKeys "^c", True
			Sleep 3000
		
			'fermeture du document PDF
			Shell "taskkill /f /IM AcroRd32*"
			Sleep 1000
		
			'puis on enregistre le contenu du presse papier dans un fichier txt
			Shell "notepad.exe /W txt/" & f1.Name & ".txt", vbNormalFocus
			Sleep 1000
			
			SendKeys "{enter}", True
			Sleep 500
			
			SendKeys "^v", True
			Sleep 500
			
			SendKeys "^s", True
			Sleep 500
			
			SendKeys "%{F4}", True
			Sleep 500
        
        End If
    Next

End Sub


Sub txt_to_database()

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(Application.CurrentProject.Path & "\txt")
    Set fc = f.Files
    
    ChDir Application.CurrentProject.Path
    
    CurrentDb.Execute "DROP TABLE horoscopes"
    CurrentDb.Execute "CREATE TABLE horoscopes (id AUTOINCREMENT, jour DATE, signe INTEGER, description TEXT(255));"
    
    For Each f1 In fc
        
        Dim day As Date
        day = Mid(f1.Name, 7, 2) & "/" & Mid(f1.Name, 5, 2) & "/" & Mid(f1.Name, 1, 4)
        
        'Ajouter la référence "Microsoft Scripting Runtime"
        Dim fso As FileSystemObject: Set fso = New FileSystemObject
        Set txtStream = fso.OpenTextFile("txt/" & f1.Name, ForReading, False, TristateTrue)

        Do While Not txtStream.AtEndOfStream
            Dim line As String
            line = txtStream.ReadLine
            
            If line = "HOROSCOPE" Then
            
                For signe = 1 To 12
                    Dim content As String
                    'nom du signe
                    txtStream.ReadLine
                    content = ""
                    For ligne = 1 To 3
                        content = content & Trim(txtStream.ReadLine) & " "
                    Next
                    content = RTrim(content)
                    CurrentDb.Execute "INSERT INTO horoscopes (jour, signe, description) VALUES (""" & day & """," & signe & ",""" & content & """)"
                Next
            End If
            
        Loop
        txtStream.Close
    Next

End Sub

