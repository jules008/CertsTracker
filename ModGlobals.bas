Attribute VB_Name = "ModGlobals"
'===============================================================
' Module ModGlobals
'===============================================================
' v1.0.0 - Initial Version
' v1.0.1 - Moved Range constants to sheets
' v1.0.2 - Added Enum Column headings
'---------------------------------------------------------------
' Date - 23 Mar 20
'===============================================================
Option Explicit

Public Const VERSION As String = "1.3.0"
Public Const ISS_DATE As String = "25 Mar 20"
Public Const RNG_SSN As String = "B2:B500"
Public Const RNG_CREW_COUNT As String = "A:A"
Public Const RNG_NAMES As String = "A:A"
Public Const RNG_LAST_COL As String = "AX"
Public Const NO_COURSES As Integer = 39
Public Const PERS_DET_NO_COLS As Integer = 10
Public Const SEC_KEY As String = "2683174"
Public Const EXPORT_FILE_PATH As String = "G:\Development Areas\Certification Tracker\Library\"
Public Const LIBRARY_FILE_PATH As String = "G:\Development Areas\Certification Tracker\Library\"
Public Const PROJECT_FILE_NAME As String = "Certification Tracker"
Public Const APP_NAME As String = "Certification Tracker"
Public USER_LEVEL As EnumUserLvl

Enum EnumUserLvl
    BasicLvl = 1
    AdminLvl
    DevLvl
End Enum

Enum EnumStatus
    Active = 1
    InActive
End Enum

Enum EnumRole
    FI = 1
    Dispatch
    Firefighter
    DriverOp
    CrewManager
    StationCaptain
    ACTraining
    ACHealthandSafety
    ACFirePrevention
    ACOps
    DeputyChief
    FireChief
End Enum

Enum EnumQual
    CPR = 1
    PPEProgram
    EMR
    Munitions
    IS100_IS700
    IS200_IS800
    HazmatAW
    HazmatOps
    FirefighterI
    FirefighterII
    TelecommunicatorI
    TelecommunicatorII
    LGVCatC
    DrvrOpPumper
    DrvrOPMWS
    HazmatTech
    FireOfficerI
    FireInpsectorI
    FireInstructorI
    IncidentSafetyOfficer
    FireOfficerII
    FireInspectorII
    FireInstructorII
    HazmatIC
    NIMS300400
    FireOfficerIII
    FireInspectorIII
    FireInstructorIII
    FireOfficerIV
    EMT
    HealthSafetyOfficer
    HazmatWMDIC
    RescueTechnicianI
    RescueTechnicianII
    PlansExaminer
    MSASCBAServicer
    WMD
    LGVCatCE
    FireEducator
End Enum

Enum EnumReport
    FFtoDO = 1
    DOtoCM
    CMtoSC
    SCtoAC
End Enum

Enum EnumTriState
    Yes = 1
    No
    Blank
End Enum

Enum EnumRW
    eRead = 1
    EWrite
    EClear
End Enum

Enum EnumExpiryStatus
    Valid = 1
    Expired
End Enum
    
Enum EnumCols
    eName = 1
    eGrade
    ePosition
    eDoB
    eFINNo
    eDoDRef
    eContract
    eWatch
    eSSN
    eStatus
End Enum

Public Function QualConvEnum(Qual As EnumQual) As String
    Select Case Qual
        Case CPR
            QualConvEnum = "CPR"
        Case EMR
            QualConvEnum = "EMR"
        Case Munitions
            QualConvEnum = "Munitions"
        Case IS100_IS700
            QualConvEnum = "IS100 & IS700"
        Case IS200_IS800
            QualConvEnum = "IS200 & IS800"
        Case HazmatAW
            QualConvEnum = "Hazmat Awareness"
        Case HazmatOps
            QualConvEnum = "Hazmat Ops"
        Case FirefighterI
            QualConvEnum = "Firefighter I"
        Case FirefighterII
            QualConvEnum = "Firefighter II"
        Case TelecommunicatorI
            QualConvEnum = "Telecommunicator I"
        Case TelecommunicatorII
            QualConvEnum = "Telecommunicator II"
        Case LGVCatC
            QualConvEnum = "LGV Cat C"
        Case DrvrOpPumper
            QualConvEnum = "Driver Op Pumper"
        Case DrvrOPMWS
            QualConvEnum = "Driver Op MWS"
        Case HazmatTech
            QualConvEnum = "Hazmat Tech"
        Case FireOfficerI
            QualConvEnum = "Fire Officer I"
        Case FireInpsectorI
            QualConvEnum = "Fire Inpsector I"
        Case FireInstructorI
            QualConvEnum = "Fire Instructor I"
        Case IncidentSafetyOfficer
            QualConvEnum = "Incident Safety Officer"
        Case FireOfficerII
            QualConvEnum = "Fire Officer II"
        Case FireInspectorII
            QualConvEnum = "Fire Inpsector II"
        Case FireInstructorII
            QualConvEnum = "Fire Instructor II"
        Case HazmatIC
            QualConvEnum = "Hazmat IC"
        Case NIMS300400
            QualConvEnum = "NIMS 300 400"
        Case FireOfficerIII
            QualConvEnum = "Fire Officer III"
        Case FireInspectorIII
            QualConvEnum = "Fire Inpsector III"
        Case FireInstructorIII
            QualConvEnum = "Fire Instructor III"
        Case FireOfficerIV
            QualConvEnum = "Fire Officer IV"
        Case EMT
            QualConvEnum = "EMT"
        Case HealthSafetyOfficer
            QualConvEnum = "Health Safety Officer"
        Case HazmatWMDIC
            QualConvEnum = "Hazmat WMD IC"
        Case RescueTechnicianI
            QualConvEnum = "Rescue Technician I"
        Case RescueTechnicianII
            QualConvEnum = "Rescue Technician II"
        Case PlansExaminer
            QualConvEnum = "Plans Examiner"
        Case MSASCBAServicer
            QualConvEnum = "MSA SCBA Servicer"
        Case WMD
            QualConvEnum = "WMD"
        Case LGVCatCE
            QualConvEnum = "LGV Cat CE"
    End Select
End Function
