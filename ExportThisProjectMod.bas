Attribute VB_Name = "ExportThisProjectMod"
'
' __        __        _    _          ____
' \ \      / /__  ___| | _| |_   _   / ___|_____   _____ _ __ __ _  __ _  ___
'  \ \ /\ / / _ \/ _ \ |/ / | | | | | |   / _ \ \ / / _ \ '__/ _` |/ _` |/ _ \
'   \ V  V /  __/  __/   <| | |_| | | |__| (_) \ V /  __/ | | (_| | (_| |  __/
'    \_/\_/ \___|\___|_|\_\_|\__, |  \____\___/ \_/ \___|_|  \__,_|\__, |\___|
'   ___  _ __    / ___|___  _|___/_ _(_) |                         |___/
'  / _ \| '_ \  | |   / _ \| '__/ _` | | |
' | (_) | | | | | |__| (_) | | | (_| | | |
'  \___/|_| |_|  \____\___/|_|  \__,_|_|_|
'
'
'01010111 01100101 01100101 01101011 01101100 01111001  01000011 01101111 01110110 01100101 01110010 01100001 01100111 01100101
'01101111 01101110  01000011 01101111 01110010 01100001 01101001 01101100
' FORREST SOFTWARE
' Copyright (c) 2016 Mateusz Forrest Milewski
'
' Permission is hereby granted, free of charge,
' to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


Global Const REPO_PATH = "C:\WORKSPACE\dev\coverage\wcoc\repo\"

Private Sub export_this_project()
    
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim VBComps As VBIDE.VBComponents
    Dim CodeMod As VBIDE.CodeModule
    
    Set VBProj = ThisWorkbook.VBProject
    Set VBComps = VBProj.VBComponents
    
    For Each VBComp In VBComps
        
        If VBComp.Type = vbext_ct_StdModule Then
            txt = VBComp.Name & ".bas"
            VBComp.Export CStr(REPO_PATH) & txt
            Debug.Print txt
            
        ElseIf VBComp.Type = vbext_ct_ClassModule Then
            txt = VBComp.Name & ".cls"
            VBComp.Export CStr(REPO_PATH) & txt
            Debug.Print txt
            
        ElseIf VBComp.Type = vbext_ct_MSForm Then
            txt = VBComp.Name & ".frm"
            VBComp.Export CStr(REPO_PATH) & txt
            Debug.Print txt
            
        End If
         
    Next VBComp
    
    MsgBox "ready!", vbInformation

End Sub


Private Sub import_this_project()
    
    
    remove_current_implementation
    
    
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.file
    Set objFSO = New Scripting.FileSystemObject
    
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim VBComps As VBIDE.VBComponents
    Dim CodeMod As VBIDE.CodeModule
    
    Set VBProj = ThisWorkbook.VBProject
    Set VBComps = VBProj.VBComponents
    
    For Each objFile In objFSO.GetFolder(XWIZ.REPO_PATH).Files
        ' body
        ' ==============================================================
        
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            VBComps.Import objFile.Path
        End If
        
        ' ==============================================================
    Next objFile
    
    MsgBox "ready!", vbInformation

End Sub


Private Sub remove_current_implementation()
    
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim VBComps As VBIDE.VBComponents
    Dim CodeMod As VBIDE.CodeModule
    
    Set VBProj = ThisWorkbook.VBProject
    Set VBComps = VBProj.VBComponents
    
    For Each VBComp In VBComps
        
        If VBComp.Type = vbext_ct_Document Then
            txt = VBComp.Name
            Debug.Print txt & " zostaje"
            
        ElseIf VBComp.Type = vbext_ct_ActiveXDesigner Then
            txt = VBComp.Name
            Debug.Print txt & " zostaje"

        ElseIf CStr(VBComp.Name) = "ExportThisProjectMod" Then
            txt = VBComp.Name
            Debug.Print txt & " zostaje"
        Else
            
            VBComps.Remove VBComp
        End If
         
    Next VBComp
    
    MsgBox "ready!", vbInformation

End Sub
