Attribute VB_Name = "Main"
Option Explicit

Dim swApp As Object
Dim gCurDoc As ModelDoc2
Dim gCurDocName As String
Dim gCurConf As String
Dim gComponents As Dictionary
Dim gKeys() As String
Dim gFSO As FileSystemObject
Dim gCurDirMask As String

Const COL_NAME = 0
Const COL_CONF = 1
Const COL_COUNT = 2

Sub Main()
    Set swApp = Application.SldWorks
    Set gFSO = New FileSystemObject
    Set gComponents = New Dictionary
    
    Set gCurDoc = swApp.ActiveDoc
    If gCurDoc Is Nothing Then Exit Sub
    If gCurDoc.GetType <> swDocASSEMBLY Then
        MsgBox "Запускать в сборках!", vbCritical
        Exit Sub
    End If
   
    gCurDirMask = LCase(GetFolderPath(gCurDoc.GetPathName) & "*")
    gCurDocName = gFSO.GetBaseName(gCurDoc.GetPathName)
    gCurConf = gCurDoc.GetActiveConfiguration.Name
    
    ResearchComponents
    
    MainForm.Caption = "Детали " & gCurDocName & " (" & gCurConf & ") "
    MainForm.Show
End Sub

Function GetFolderPath(pathName As String) As String
    GetFolderPath = Left(pathName, InStrRev(pathName, "\"))
End Function

Function ResearchComponents() 'mask for button
    Dim onlyInCurrentDir As Boolean
    
    onlyInCurrentDir = MainForm.chkCurDir.Value
    gComponents.RemoveAll
    SearchComponents gCurDoc, onlyInCurrentDir, gCurDocName, gCurConf
    If gComponents.count > 0 Then
        ReDim gKeys(gComponents.count - 1)
    End If
    FilterAndPrint
End Function

Function FilterAndPrint() 'mask for button
    Dim key_ As Variant
    Dim index As Long
    Dim info As ComponentInfo
    Dim masks As Variant
    
    index = -1
    If gComponents.count > 0 Then
        masks = CreateFilters
        For Each key_ In gComponents.keys
            Set info = gComponents(key_)
            If CheckUserFilter(info.baseName, masks) Then
                index = index + 1
                gKeys(index) = key_
            End If
        Next
        If index >= 0 Then
            QuickSort gKeys, LBound(gKeys), index
        End If
    End If
    PrintComponents index
End Function

Function CheckUserFilter(baseName As String, masks As Variant) As Boolean
    Dim i As Integer
    
    CheckUserFilter = True
    If Not IsArrayEmpty(masks) Then
        For i = LBound(masks) To UBound(masks)
            If Not LCase(baseName) Like masks(i) Then
                CheckUserFilter = False
                Exit Function
            End If
        Next
    End If
End Function

Function CreateFilters() As String()
    Dim filter As String
    Dim words As Variant
    Dim i As Integer
    
    filter = MainForm.txtFilter.Value
    words = Split(filter, " ")
    If Not IsArrayEmpty(words) Then
        For i = LBound(words) To UBound(words)
            words(i) = LCase("*" & words(i) & "*")
        Next
    End If
    CreateFilters = words
End Function

Function IsArrayEmpty(ByRef anArray As Variant) As Boolean
    Dim i As Integer
  
    On Error GoTo ArrayIsEmpty
    IsArrayEmpty = LBound(anArray) > UBound(anArray)
    Exit Function
ArrayIsEmpty:
    IsArrayEmpty = True
End Function

Sub SearchComponents(asm As AssemblyDoc, onlyInCurrentDir As Boolean, asmName As String, asmConf As String)
    Dim comp_ As Variant
    Dim comp As Component2
    Dim doc As ModelDoc2
    Dim subAsmName As String
    Dim subAsmConf As String
    
    For Each comp_ In asm.GetComponents(True)
        Set comp = comp_
        If comp.IsSuppressed Then  'погашен
            GoTo NextFor
        End If
        Set doc = comp.GetModelDoc2
        If doc Is Nothing Then  'не найден
            GoTo NextFor
        End If
        If doc.GetType = swDocASSEMBLY Then
            If Not onlyInCurrentDir Or (LCase(doc.GetPathName) Like gCurDirMask) Then
                subAsmName = gFSO.GetBaseName(comp.GetPathName)
                subAsmConf = comp.ReferencedConfiguration
                SearchComponents doc, onlyInCurrentDir, subAsmName, subAsmConf
            End If
        Else  'doc is part
            AddComponent comp, asmName, asmConf
        End If
NextFor:
    Next
End Sub

Function CreateKey(baseName As String, conf As String) As String
    CreateKey = LCase(baseName & "/*@@*/" & conf)
End Function

Sub AddComponent(comp As Component2, asmName As String, asmConf As String)
    Dim baseName As String
    Dim conf As String
    Dim key As String
    Dim item As ComponentInfo
    
    baseName = gFSO.GetBaseName(comp.GetPathName)
    conf = comp.ReferencedConfiguration
    key = CreateKey(baseName, conf)
    
    If Not gComponents.Exists(key) Then
        Set item = New ComponentInfo
        item.baseName = baseName
        item.conf = conf
        item.totalCount = 0
        Set item.where = New Dictionary
        gComponents.Add key, item
    End If
        
    AddWherePartIsUsed key, asmName, asmConf
    'If MsgBox(comp.GetParent Is Nothing, vbOKCancel) = vbCancel Then End
    gComponents(key).totalCount = gComponents(key).totalCount + 1
End Sub

Sub AddWherePartIsUsed(key As String, asmName As String, asmConf As String)
    Dim asmKey As String
    Dim item As WhereInfo
    
    asmKey = CreateKey(asmName, asmConf)
    If Not gComponents(key).where.Exists(asmKey) Then
        Set item = New WhereInfo
        item.asmName = asmName
        item.conf = asmConf
        item.count = 0
        gComponents(key).where.Add asmKey, item
    End If
    
    gComponents(key).where(asmKey).count = gComponents(key).where(asmKey).count + 1
End Sub

Function PrintComponents(topKeysBound As Long) 'mask for button
    Dim info As ComponentInfo
    Dim i As Integer
    
    With MainForm.lstDeps
        .Clear
        For i = 0 To topKeysBound
            Set info = gComponents(gKeys(i))
                .AddItem
                .List(.ListCount - 1, COL_NAME) = info.baseName
                .List(.ListCount - 1, COL_CONF) = info.conf
                .List(.ListCount - 1, COL_COUNT) = Str(info.totalCount)
        Next
    End With
End Function

Sub ShowWhereIsPartUsed(index As Integer)
    Dim key As String
    Dim info As ComponentInfo
    Dim asmKey_ As Variant
    Dim text As String
    
    key = gKeys(index)
    Set info = gComponents(key)
    For Each asmKey_ In info.where
        text = text & info.where(asmKey_).asmName & " (" & info.where(asmKey_).conf & ") --" & Str(info.where(asmKey_).count) & vbNewLine
    Next
    MsgBox text, , info.conf
End Sub

Function ExitApp()  'mask for button
    Unload MainForm
    End
End Function
