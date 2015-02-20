Sub calculation_density()
'Name of worksheet containing initial data (DYRESM output)
source1 = "temperature"     
source2 = "conductivity"
'Name of worksheet for resulting table
target = "density"       
schmidt = "Schmidt stability"
finalvalue = 1000
'Number of measurements in vertical profile
f = 16
'Number of days of measurements
sp = 370           
'Volume of Lake Ammersee [m³]
totalvolume = 1750010000
'Surface area of Lake Ammersee [km²]     
surface = 46600000#    
a = 2
b = 3
c = 4
E = 3
For d = 1 To sp
For i = 1 To f
Sheets(source1).Select
Cells(a, b).Select
'Columns containing temperature measurements (tempms)
tempms = Cells(a, b)
Selection.Copy
'Enter date to table of Schmidt stability 
If a = 2 Then Sheets(schmidt).Select
Cells(a, b).Select
ActiveSheet.Paste
Sheets(target).Select
Cells(a, c + 1).Select
ActiveSheet.Paste
End If
Sheets(source2).Select
Cells(a, b).Select
'Columns containing conductivity measurements (cms)
cms = Cells(a, b)
Sheets(target).Select
c = c + 1
Cells(a, c).Select
texponent3 = tempms * tempms * tempms
texponent2 = tempms * tempms
'Circumvent the date
If a > 2 Then
'Calculation of density and mean density [kg m-3   10-3] as well as depth at which mean density  occurs [m]
term3 = ((0.059385 * (tempms * tempms * tempms) - 8.56272 * (tempms * tempms) + 65.4891 * tempms) * 0.000001 + 0.99984298) + 0.64 * 0.000001 * (cms * (-0.00001222651 * (tempms * tempms * tempms) + 0.00114842 * (tempms * tempms) - 0.0541369 * tempms + 1.72118))
ActiveCell.FormulaR1C1 = term3
sum_density = sum_dichte + term3
End If
a = a + 1
c = c - 1
Next i
c = c + 1
Cells(a, c).Select
mean_density = sum_density / (f - 1)
ActiveCell.FormulaR1C1 = mean_density
schmidtstability = diff * totalvolume * volweightedsum * 1000 * 9.81
schmidtstabilsurface = schmidtstability / surface
Sheets(schmidt).Select
Cells(3, E).Select
ActiveCell.FormulaR1C1 = schmidtstabilsurface
E = E + 1
volwithdensitysum = 0
voldensitysum = 0
s = 0
s0 = 0
diff = 0
volwithsum = 0
schmidtstability = 0
volweightedsum = 0
sum_density = 0
a = 2
b = b + 1
Next d
End Sub

Sub schmidtstability()
'Ensure that Sub calculation_density runs first    
source1 = "temperature"
source2 = "conductivity"
target = "density"       
schmidt = "Schmidt_stability"
finalvalue = 1000
f = 16
sp = 370           
totalvolume = 1750010000
surface = 46600000#    
a = 2
c = 6
E = 3
For d = 1 To sp
For i = 1 To f
If a > 2 Then 
'Calculation of Schmidt stability [KJ m-2]
Sheets(target).Select
mean_density = Cells(18, c)
depth_mean_density = Cells(19, c)
term = (Cells(a, c) – mean_density) * (Cells(a, 2) * (Cells(a, 3) – depth_mean_density))
termsum = term + termsum
End If
a = a + 1
Next i
schmidtstability = (termsum / surface) * 9.81
Sheets(schmidt).Select
Cells(3, E).Select
ActiveCell.FormulaR1C1 = schmidtstability
E = E + 1
schmidtstability = 0
a = 2
c = c + 1
termsum = 0
Next d
End Sub

