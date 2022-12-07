' Module source code
Private Declare PtrSafe Function auto_open Lib "~/dev/custom_code.dll" (ByVal app As LongPtr) As Long

Public Sub Auto_Open()
    ' execute code when PowerPoint starts
    auto_open ObjPtr(Application)
End Sub
