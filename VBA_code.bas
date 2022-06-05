Attribute VB_Name = "Module1"
'defining function that relates pressure and concentration for individual compound

Function Tothcalc(A, B, C, p)
'calculation of amount adsorbed using toth equation
Tothcalc = A * p / (B + p ^ C) ^ (1 / C)
End Function

Function integ(P_max, A, B, C)

n = 1000
j = 1
integ = 0
p = 0
pa = P_max / n
p1 = 10


integ = (Tothcalc(A, B, C, pa) / 2)
p = pa


Do Until j = n
integ = integ + (Tothcalc(A, B, C, p) / p + Tothcalc(A, B, C, (p + pa)) / (p + pa)) * pa / 2
p1 = integ
j = j + 1
p = p + pa
Loop

integ = p1
j = 1

End Function



Sub Toth()
'returns toth parameters from isotherm data points

'opens macro that deletes previous information, so solver can iterate faster
Reset_sheet

'initialize parameters
Range("f3").Value = 1
Range("g3").Value = 0.5
Range("h3").Value = 1
Range("I3").Value = 1
Range("J3").Value = 0.5
Range("k3").Value = 1


'determine number of points for the two sets of data

Psize1 = 1
p = 1
C = 1

Do While C > 0
p = Range("a" & (Psize1 + 6))

If p = 0 Then
C = 0
ElseIf p > 0 Then
Psize1 = Psize1 + 1
End If

Loop

Psize2 = 1
p = 1
C = 1

Do While C > 0
p = Range("c" & (Psize2 + 6)) * 10000
p = p * 1000

If p = 0 Then
C = 0
ElseIf p > 0 Then
Psize2 = Psize2 + 1
End If

Loop


j = 1

'loop to insert formula into cells for component B
Do Until (j = Psize1 + 1)
Range("f" & (j + 6)).Formula = "=Tothcalc(f3,g3,h3, A" & (j + 5) & ")"

j = j + 1

Loop

I = 1

'loop to insert formula into cells for component A
Do Until (I = Psize2 + 1)
Range("i" & (I + 6)).Formula = "=Tothcalc(i3,j3,k3, C" & (I + 5) & ")"
I = I + 1
Loop

j = 1


'loop to calculate squared error between fit and data
Do Until (j = Psize1 + 1)
Range("g" & (j + 6)).Formula = "=(b" & (j + 5) & "-f" & (j + 6) & ")^2"
Range("h" & (j + 6)).Formula = "=(b" & (j + 5) & "-average(b$6:b$" & (Psize1 + 5) & "))^2"
j = j + 1
Loop

I = 1

Do Until (I = Psize2 + 1)
Range("J" & (I + 6)).Formula = "=(d" & (I + 5) & "-I" & (I + 6) & ")^2"
Range("k" & (I + 6)).Formula = "=(d" & (I + 5) & "-average(d$6:d$" & (Psize2 + 5) & "))^2"
I = I + 1
Loop


Range("g" & (j + 6)).Formula = "=sum(g7:g" & (j + 5) & ")"
Range("f" & (j + 6)).Value = "sum of errors"
Range("f4").Formula = "R^2"
Range("f5").Formula = "=1-g" & (j + 6) & "/sum(h7:h" & (j + 5) & ")"


Range("J" & (I + 6)).Formula = "=sum(J7:J" & (I + 5) & ")"
Range("I" & (I + 6)).Value = "sum of errors"
Range("i4").Formula = "R^2"
Range("i5").Formula = "=1-j" & (I + 6) & "/sum(k7:k" & (I + 5) & ")"
Range("r8").Select
'plotting chart with fit and data

    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.ChartType = xlXYScatterSmoothNoMarkers
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(1).XValues = "='Component A Isotherm'!$A$6:$A$" & (j + 4)
    ActiveChart.SeriesCollection(1).Values = "='Component A Isotherm'!$f$7:$f$" & (j + 5)
    ActiveChart.SeriesCollection(1).Name = "=""toth fit B"""
   ActiveChart.SeriesCollection.NewSeries
   ActiveChart.SeriesCollection(2).XValues = "='Component A Isotherm'!$A$6:$A$" & (j + 4)
    ActiveChart.SeriesCollection(2).Values = "='Component A Isotherm'!$b$6:$b$" & (j + 4)
    ActiveChart.SeriesCollection(2).Name = "=""isotherm data B"""
  ActiveChart.SeriesCollection(2).ChartType = xlXYScatter
   ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(3).XValues = "='Component A Isotherm'!$c$6:$c$" & (I + 4)
    ActiveChart.SeriesCollection(3).Values = "='Component A Isotherm'!$i$7:$i$" & (I + 5)
    ActiveChart.SeriesCollection(3).Name = "=""toth fit A"""
      ActiveChart.SeriesCollection(3).ChartType = xlXYScatterSmoothNoMarkers
   ActiveChart.SeriesCollection.NewSeries
   ActiveChart.SeriesCollection(4).XValues = "='Component A Isotherm'!$c$6:$c$" & (I + 4)
    ActiveChart.SeriesCollection(4).Values = "='Component A Isotherm'!$d$6:$d$" & (I + 4)
    ActiveChart.SeriesCollection(4).Name = "=""isotherm data A"""
  ActiveChart.SeriesCollection(4).ChartType = xlXYScatter


 'applying solver
 
      SolverOk SetCell:="$g$" & (j + 6), MaxMinVal:=2, ValueOf:=0, ByChange:="$f$3,$g$3,$h$3"
    SolverAdd CellRef:="$G$3", Relation:=3, FormulaText:="0.001"
    SolverOk SetCell:="$g$" & (j + 6), MaxMinVal:=2, ValueOf:=0, ByChange:="$f$3,$g$3,$h$3"
    SolverAdd CellRef:="$G$3", Relation:=3, FormulaText:="0.001"
   SolverSolve
    
 
    SolverOk SetCell:="$j$" & (I + 6), MaxMinVal:=2, ValueOf:=0, ByChange:="$i$3,$j$3,$k$3"
    SolverAdd CellRef:="$j$3", Relation:=3, FormulaText:="0.001"
        
    SolverOk SetCell:="$j$" & (I + 6), MaxMinVal:=2, ValueOf:=0, ByChange:="$i$3,$j$3,$k$3"
    SolverAdd CellRef:="$j$3", Relation:=3, FormulaText:="0.001"
    SolverSolve
    
  'calculating adsorbed phase composition
  
  'input pressure to obtain spreading pressure
    p = InputBox("Select maximum pressure to determine composition", "Pressure", 50)
    Range("'SP'!c4").Value = p
    compA = InputBox("Select gas composition of A", "Component A composition", 0.5)
    Range("'SP'!d6").Value = compA
   'write formula that calculate spreading pressure using numerical integration
   
    Range("'SP'!f8").Formula = "=integ(D8,'Component A Isotherm'!I3,'Component A Isotherm'!J3,'Component A Isotherm'!K3)"
    Range("'SP'!g8").Formula = "=integ(C8,'Component A Isotherm'!F3,'Component A Isotherm'!G3,'Component A Isotherm'!h3)"
    Range("'SP'!g10").Formula = "=abs(g8-f8)"
    Range("'SP'!c8").Value = p + 0.1
    Range("'SP'!d16").Formula = "=(D13/tothcalc('Component A Isotherm'!I3,'Component A Isotherm'!J3,'Component A Isotherm'!K3,SP!D8)+SP!D14/tothcalc('Component A Isotherm'!F3,'Component A Isotherm'!G3,'Component A Isotherm'!H3,SP!C8))^-1"
    
Sheets("SP").Select
    
      SolverOk SetCell:="'SP'!$g$10", MaxMinVal:=2, ValueOf:=0, ByChange:="'SP'!$c$8"
           SolverAdd CellRef:="$C$8", Relation:=3, FormulaText:="$c$4"
      SolverOk SetCell:="'SP'!$g$10", MaxMinVal:=2, ValueOf:=0, ByChange:="'SP'!$c$8"
           SolverAdd CellRef:="$C$8", Relation:=3, FormulaText:="$c$4"
    SolverSolve
    
    
    
    
    
    
    
End Sub



Sub Reset_sheet()
'
' Macro1 Macro
'

'
    Range("e7:Q7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("F5").Select
    Selection.ClearContents
    Range("I5").Select
    Selection.ClearContents
    Range("G14").Select
    Sheets("SP").Select
    Range("f8:g8").Select
    Selection.ClearContents
    Range("d16").Select
    Selection.ClearContents
    Sheets("Component A Isotherm").Select
    
End Sub

Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    SolverOk SetCell:="$J$15", MaxMinVal:=2, ValueOf:=0, ByChange:="$I$3,$J$3,$K$3" _
        , Engine:=1, EngineDesc:="GRG Nonlinear"
    SolverOk SetCell:="$J$15", MaxMinVal:=2, ValueOf:=0, ByChange:="$I$3,$J$3,$K$3" _
        , Engine:=1, EngineDesc:="GRG Nonlinear"
    SolverSolve
End Sub
