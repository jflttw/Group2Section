Option Strict On
'Imports OpenSTAADUI
'Imports Microsoft.Office.Interop.Excel
Module Module1

    Public Sub Main()
        '''''''''''''''''''''''''''''''' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Get the application object and set openstaad
        Dim ObjStaad As OpenSTAADUI.OpenSTAAD
        ObjStaad = CType(GetObject(, "StaadPro.OpenSTAAD"), OpenSTAADUI.OpenSTAAD)
        'set Geometry class
        Dim ObjGeometry As OpenSTAADUI.OSGeometryUI
        ObjGeometry = CType(ObjStaad.Geometry, OpenSTAADUI.OSGeometryUI)
        'set Property class
        Dim ObjProperty As OpenSTAADUI.OSPropertyUI
        ObjProperty = CType(ObjStaad.Property, OpenSTAADUI.OSPropertyUI)
        Dim ObjRng As Microsoft.Office.Interop.Excel.Range
        'check std model exist
        Dim strFileName As String = "*.std" 'std file name here
        Dim bIncludePath As Boolean = True
        Dim tempFile As Object = CObj(strFileName)
        ObjStaad.GetSTAADFile(tempFile, bIncludePath)
        strFileName = CType(tempFile, String)

        If strFileName = "" Then
            MsgBox("This app requires STAAD.Pro to have a model open")
            Exit Sub
        End If
        '''''''''''''''''''''''''''''''''get group name'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'set group type
        Dim Grouptype As Integer
        Grouptype = 2 '1node ,2beam
        'local temp variable
        Dim Grouptypetemp As Object
        Grouptypetemp = CObj(Grouptype)
        'get total number of groups
        Dim GroupNO As Integer
        GroupNO = CInt(ObjGeometry.GetGroupCount(Grouptypetemp))


        ''group temp variable
        Dim GroupArr() As String
        ReDim GroupArr(GroupNO - 1)
        'local temp variable
        Dim GroupArrtemp As Object
        GroupArrtemp = CObj(GroupArr)
        'get group list arry
        ObjGeometry.GetGroupNames(Grouptypetemp, GroupArrtemp)
        ' Cast temp variable back to the original type
        GroupArr = CType(GroupArrtemp, String())




        ''Select beams form each group
        'Dim BeamNos() As Integer
        'ReDim BeamNos(GroupBeamNo - 1)
        'BeamNos = EntityList
        'Dim BeamNostemp As Object
        'BeamNostemp = CObj(BeamNos)
        'ObjGeometry.SelectMultipleBeams(BeamNostemp)


        '''''''''''''''''''''''''''''''' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Get the application object and set Application
        Dim ObjExcel As Microsoft.Office.Interop.Excel.Application
        ObjExcel = CType(GetObject(, "excel.Application"), Microsoft.Office.Interop.Excel.Application)
        'set activeworkbook name space
        Dim Objworkbook As Microsoft.Office.Interop.Excel.Workbook
        Objworkbook = CType(ObjExcel.ActiveWorkbook, Microsoft.Office.Interop.Excel.Workbook)
        Dim Objsheet As Microsoft.Office.Interop.Excel.Worksheet
        Objsheet = CType(Objworkbook.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        'fill cell A and B column with Group name
        Objsheet.Cells(1, 1) = "GroupName"
        Objsheet.Cells(1, 2) = "SectionName"
        Dim i As Integer

        For i = 1 To GroupNO
            Objsheet.Cells(1 + i, 1) = GroupArr(i - 1)
        Next

        '''''''''''''''''''''''''''''''' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Console.WriteLine("Complete 'SectionName(Column B)' and Press anykey to assign sections!")
        Console.ReadKey()
        '''''''''''''''''''''''''''''''' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'assign section from each group
        Dim GroupBeamNo As Integer
        Dim GroupNewName As String

        For i = 1 To GroupNO


            ObjRng = CType(Objsheet.Cells(i + 1, 2), Microsoft.Office.Interop.Excel.Range)
            GroupNewName = CStr(ObjRng.Value)
            'set section protperty in able
            Dim Country As Integer
            Country = 7 ';;;;;;;;;;;;;need to be fill
            Dim Countrytemp As Object
            Countrytemp = CObj(Country)
            Dim SectionName As String
            SectionName = GroupNewName
            Dim SectionNametemp As Object
            SectionNametemp = CObj(SectionName)
            ObjProperty.CreateBeamPropertyFromTable(Countrytemp, SectionNametemp, 0, 0, 0)

            'set group type
            Dim GroupTXT As String
            GroupTXT = GroupArr(i - 1) 'section name
            'local temp variable
            Dim GroupTXTtemp As Object
            GroupTXTtemp = CObj(GroupTXT)
            'get total number of groups
            GroupBeamNo = CInt(ObjGeometry.GetGroupEntityCount(GroupTXTtemp))
            'get beams nos form each group
            Dim Groupname As String
            Groupname = GroupArr(i - 1)
            Dim Groupnametemp As Object
            Groupnametemp = CObj(Groupname)
            Dim EntityList() As Integer
            ReDim EntityList(GroupBeamNo - 1)
            Dim EntityListtemp As Object
            EntityListtemp = CObj(EntityList)
            ObjGeometry.GetGroupEntities(Groupnametemp, EntityListtemp)
            EntityList = CType(EntityListtemp, Integer())
            ' assign section protperty to beam 
            Dim nBeamNo() As Integer
            ReDim nBeamNo(GroupBeamNo - 1)
            nBeamNo = EntityList
            Dim nBeamNotemp As Object
            nBeamNotemp = CObj(nBeamNo)
            Dim nProperty As Integer
            nProperty = i
            Dim nPropertytemp As Object
            nPropertytemp = CObj(nProperty)
            ObjProperty.AssignBeamProperty(nBeamNotemp, nPropertytemp)

        Next



        'Console.WriteLine(GroupArr(3))
        'Console.ReadKey()






    End Sub

End Module

'GetGroupEntities
'GetGroupEntityCount

'Property.CreateBeamPropertyFromTable
'Property.AssignBeamProperty

'GroupName SectionName
'_IPE240 IPE240
'_IPE200 IPE200
'_CHS1010.6X6.3 101.6X6.3CHS
'_HEB240 HE240B
