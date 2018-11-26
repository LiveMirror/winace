VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmpItemBuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Change Log:
'300716 rjk: Create to add commodities changes easy for server release 4.3.34.

'local variable(s) to hold property value(s)
Private mvarType As String 'local copy
Private mvarID As String 'local copy
Private mvarDesc As String 'local copy
Private mvarCost As Long 'local copy
Private mvavTech As Integer 'local copy
Private mvarAvail As Integer 'local copy
Private mvarLCM As Integer 'local copy
Private mvarHCM As Integer 'local copy
Private mvarMil As Integer 'local copy
Private mvarGun As Integer 'local copy
Private mvarOil As Integer 'local copy
Private mvarRad As Integer 'local copy
Private mvarCiv As Integer 'local copy
Private mvarPet As Integer 'local copy
Private mvarShell As Integer 'local copy
Private mvarIron As Integer 'local copy
Private mvarDust As Integer 'local copy
Private mvarBar As Integer 'local copy
Private mvarFood As Integer 'local copy
Private mvarUW As Integer 'local copy
Private mvarStat(1 To 20) As Integer 'local copy
Private mvarCap(1 To 20) As String 'local copy
Private mvarChrIndex As Integer 'local copy
Private mvarBase_Spd As Integer 'local copy
Private mvarBase_Frnge As Integer 'local copy
Private mvarBase_Att As Integer 'local copy
Private mvarBase_Def As Integer 'local copy
Private mvarBase_Acc As Integer 'local copy
Private mvarBase_Load As Integer 'local copy
Private mvarBase_Vul As Integer 'local copy
Private mvarBase_Dam As Integer 'local copy
Private mvarBase_Aaf As Integer 'local copy
Private mvarBWork As Integer 'local copy

'Private Sub Class_Initialize()
'End Sub


Public Property Let BuildType(ByVal vType As String)
mvarType = vType
End Property

Public Property Get BuildType() As String
BuildType = mvarType
End Property

Public Property Let id(ByVal vID As String)
mvarID = vID
End Property

Public Property Get id() As String
id = mvarID
End Property

Public Property Let desc(ByVal vDesc As String)
mvarDesc = vDesc
End Property

Public Property Get desc() As String
desc = mvarDesc
End Property

Public Property Let Cost(ByVal vCost As Long)
mvarCost = vCost
End Property

Public Property Get Cost() As Long
Cost = mvarCost
End Property

Public Property Let Tech(ByVal vTech As Integer)
mvavTech = vTech
End Property

Public Property Get Tech() As Integer
Tech = mvavTech
End Property

Public Property Let avail(ByVal vAvail As Integer)
mvarAvail = vAvail
End Property

Public Property Get avail() As Integer
avail = mvarAvail
End Property

Public Property Let lcm(ByVal vLCM As Integer)
mvarLCM = vLCM
End Property

Public Property Get lcm() As Integer
lcm = mvarLCM
End Property

Public Property Let hcm(ByVal vHCM As Integer)
mvarHCM = vHCM
End Property

Public Property Get hcm() As Integer
hcm = mvarHCM
End Property

Public Property Let Mil(ByVal vMil As Integer)
mvarMil = vMil
End Property

Public Property Get Mil() As Integer
Mil = mvarMil
End Property

Public Property Let gun(ByVal vGun As Integer)
mvarGun = vGun
End Property

Public Property Get gun() As Integer
gun = mvarGun
End Property

Public Property Let Oil(ByVal vOil As Integer)
mvarOil = vOil
End Property

Public Property Get Oil() As Integer
Oil = mvarOil
End Property

Public Property Let Rad(ByVal vRad As Integer)
mvarRad = vRad
End Property

Public Property Get Rad() As Integer
Rad = mvarRad
End Property

Public Property Let Civ(ByVal vCiv As Integer)
mvarCiv = vCiv
End Property

Public Property Get Civ() As Integer
Civ = mvarCiv
End Property

Public Property Let Pet(ByVal vPet As Integer)
mvarPet = vPet
End Property

Public Property Get Pet() As Integer
Pet = mvarPet
End Property

Public Property Let Shell(ByVal vShell As Integer)
mvarShell = vShell
End Property

Public Property Get Shell() As Integer
Shell = mvarShell
End Property

Public Property Let Iron(ByVal vIron As Integer)
mvarIron = vIron
End Property

Public Property Get Iron() As Integer
Iron = mvarIron
End Property

Public Property Let Dust(ByVal vDust As Integer)
mvarDust = vDust
End Property

Public Property Get Dust() As Integer
Dust = mvarDust
End Property

Public Property Let Bar(ByVal vBar As Integer)
mvarBar = vBar
End Property

Public Property Get Bar() As Integer
Bar = mvarBar
End Property

Public Property Let Food(ByVal vFood As Integer)
mvarFood = vFood
End Property

Public Property Get Food() As Integer
Food = mvarFood
End Property

Public Property Let UW(ByVal vUW As Integer)
mvarUW = vUW
End Property

Public Property Get UW() As Integer
UW = mvarUW
End Property

Public Property Let Stat(ByVal vIndex As Integer, ByVal vStat As Integer)
mvarStat(vIndex) = vStat
End Property

Public Property Get Stat(ByVal vIndex As Integer) As Integer
Stat = mvarStat(vIndex)
End Property

Public Property Let Cap(ByVal vIndex As Integer, ByVal vCap As String)
mvarCap(vIndex) = vCap
End Property

Public Property Get Cap(ByVal vIndex As Integer) As String
Cap = mvarCap(vIndex)
End Property

Public Property Let ChrIndex(ByVal vChrIndex As Integer)
mvarChrIndex = vChrIndex
End Property

Public Property Get ChrIndex() As Integer
ChrIndex = mvarChrIndex
End Property

Public Property Let Base_Spd(ByVal vBase_Spd As Integer)
mvarBase_Spd = vBase_Spd
End Property

Public Property Get Base_Spd() As Integer
Base_Spd = mvarBase_Spd
End Property

Public Property Let Base_Frnge(ByVal vBase_Frnge As Integer)
mvarBase_Frnge = vBase_Frnge
End Property

Public Property Get Base_Frnge() As Integer
Base_Frnge = mvarBase_Frnge
End Property

Public Property Let Base_Att(ByVal vBase_Att As Integer)
mvarBase_Att = vBase_Att
End Property

Public Property Get Base_Att() As Integer
Base_Att = mvarBase_Att
End Property

Public Property Let Base_Def(ByVal vBase_Def As Integer)
mvarBase_Def = vBase_Def
End Property

Public Property Get Base_Def() As Integer
Base_Def = mvarBase_Def
End Property

Public Property Let Base_Acc(ByVal vBase_Acc As Integer)
mvarBase_Acc = vBase_Acc
End Property

Public Property Get Base_Acc() As Integer
Base_Acc = mvarBase_Acc
End Property

Public Property Let Base_Load(ByVal vBase_Load As Integer)
mvarBase_Load = vBase_Load
End Property

Public Property Get Base_Load() As Integer
Base_Load = mvarBase_Load
End Property

Public Property Let Base_Vul(ByVal vBase_Vul As Integer)
mvarBase_Vul = vBase_Vul
End Property

Public Property Get Base_Vul() As Integer
Base_Vul = mvarBase_Vul
End Property

Public Property Let Base_Dam(ByVal vBase_Dam As Integer)
mvarBase_Dam = vBase_Dam
End Property

Public Property Get Base_Dam() As Integer
Base_Dam = mvarBase_Dam
End Property

Public Property Let Base_Aaf(ByVal vBase_Aaf As Integer)
mvarBase_Aaf = vBase_Aaf
End Property

Public Property Get Base_Aaf() As Integer
Base_Aaf = mvarBase_Aaf
End Property

Public Property Let BWork(ByVal vBWork As Integer)
mvarBWork = vBWork
End Property

Public Property Get BWork() As Integer
BWork = mvarBWork
End Property

Public Sub InsertOrUpdateDB()
Dim strErrorField As String
Dim hldBuildIndex As String
Dim iIndex As Integer

strErrorField = "Build Infrastructure"
hldBuildIndex = rsBuild.Index
rsBuild.Index = "PrimaryKey"
    
If rsBuild.BOF And rsBuild.EOF Then
    rsBuild.AddNew
    rsBuild.Fields("type") = mvarType
Else
    rsBuild.Seek "=", mvarType, mvarID
    If rsBuild.NoMatch Then
        rsBuild.AddNew
        rsBuild.Fields("type") = mvarType
    Else
        rsBuild.Edit
    End If
End If
rsBuild.Fields("id") = mvarID
rsBuild.Fields("desc") = mvarDesc
rsBuild.Fields("cost") = mvarCost
rsBuild.Fields("lcm") = mvarLCM
rsBuild.Fields("hcm") = mvarHCM
rsBuild.Fields("mil") = mvarMil
rsBuild.Fields("gun") = mvarGun
rsBuild.Fields("oil") = mvarOil
rsBuild.Fields("rad") = mvarRad
rsBuild.Fields("civ") = mvarCiv
rsBuild.Fields("pet") = mvarPet
rsBuild.Fields("shell") = mvarShell
rsBuild.Fields("iron") = mvarIron
rsBuild.Fields("dust") = mvarDust
rsBuild.Fields("bar") = mvarBar
rsBuild.Fields("food") = mvarFood
rsBuild.Fields("uw") = mvarUW

For iIndex = 1 To 20
rsBuild.Fields("stat" + CStr(iIndex)) = mvarStat(iIndex)
rsBuild.Fields("cap" + CStr(iIndex)) = mvarCap(iIndex)
Next
rsBuild.Fields("chr_i") = mvarChrIndex
rsBuild.Fields("base spd") = mvarBase_Spd
rsBuild.Fields("base frnge") = mvarBase_Frnge
rsBuild.Fields("base att") = mvarBase_Att
rsBuild.Fields("base def") = mvarBase_Def
rsBuild.Fields("base acc") = mvarBase_Acc
rsBuild.Fields("base load") = mvarBase_Load
rsBuild.Fields("base vul") = mvarBase_Vul
rsBuild.Fields("base dam") = mvarBase_Dam
rsBuild.Fields("base aaf") = mvarBase_Aaf
rsBuild.Fields("bwork") = mvarBWork
rsBuild.Update
rsBuild.Index = hldBuildIndex

End Sub

