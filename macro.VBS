
Private Sub Calculate_Clearance_or_Interferance_Click() 'This sub calculates Cleareance or Interferance when a user clicks on the comand button

'initialise contact diameter variable with their respective values from the data worksheet
Contact_Diameter = Worksheets("Data").Range("H4")

'initialise upper shaft tolerance variable with their respective values from the data worksheet
Shaft_Tol_Upper = Worksheets("Data").Range("H8")

'initialise lower shaft tolerance variable with their respective values from the data worksheet
Shaft_Tol_Lower = Worksheets("Data").Range("H9")

'initialise upper hole tolerance variable with their respective values from the data worksheet
Hole_Tol_Upper = Worksheets("Data").Range("H6")

'initialise lower hole tolerance variable with their respective values from the data worksheet
Hole_Tol_Lower = Worksheets("Data").Range("H7")

'define maximum shaft diameter
Max_Shaft_Dia = Contact_Diameter + Shaft_Tol_Upper

'define minimum shaft diameter
Min_Shaft_Dia = Contact_Diameter + Shaft_Tol_Lower

'define maximum hole diameter
Max_Hole_Dia = Contact_Diameter + Hole_Tol_Upper

'define minimum hole diameter
Min_Hole_Dia = Contact_Diameter + Hole_Tol_Lower
        
'If minimum hole diameter is greater than maximum shaft diameter, then:
If Min_Hole_Dia > Max_Shaft_Dia Then
    
    Max_Clearance = Max_Hole_Dia - Min_Shaft_Dia    'maximum clearence is equal to maximum hole diameter less minimum shaft diameter
    Min_Clearance = Min_Hole_Dia - Max_Shaft_Dia    'minimum clearence is equal to minimum hole diameter less maximum shaft diameter

    'Define the format of maximum and minimum clearence to three decimal places
    Worksheets("Data").Range("M3") = Format(Max_Clearance, "#,##0.000")
    Worksheets("Data").Range("M4") = Format(Min_Clearance, "#,##0.000")
    
    'Assign macro Clearence in cell A8 of the data worksheet
    Worksheets("Data").Range("A8") = "CLEARANCE"
    
    'Create a dialog box to show the maximum and minimum clearence for the assembly to 3 decimal places with units
    MsgBox "The maximum Clearance for this assembly is, " & Format(Max_Clearance, "#,##0.000") & " mm. " & "The minimum Clearance for this assembly is, " & Format(Min_Clearance, "#,##0.000") & " mm."
    
End If  'end if condition
        
'If minimum shaft diameter is greater than maximum hole diameter, then:
If Min_Shaft_Dia > Max_Hole_Dia Then

    Max_Interferance = Max_Shaft_Dia - Min_Hole_Dia 'maximum interferance is equal to maximum shaft diameter less minimum hiole diameter
    Min_Interferance = Min_Shaft_Dia - Max_Hole_Dia 'minimum interferance is equal to minimum shaft diameter less maximum hole diameter

    'Define the format of maximum and minimum interference to three decimal places
    Worksheets("Data").Range("M6") = Format(Max_Interferance, "#,##0.000")
    Worksheets("Data").Range("M7") = Format(Min_Interferance, "#,##0.000")
    
    'Assign macro Interferance in cell A8 of the data worksheet
    Worksheets("Data").Range("A8") = "INTERFERANCE"
 
    'Create a dialog box to show the maximum and minimum interferance for the assembly to 3 decimal places with units
    MsgBox "The maximum Interferance for this assembly is, " & Format(Max_Interferance, "#,##0.") & " mm. " & "The minimum Interferance for this assembly is, " & Format(Min_Interferance, "#,##0.000") & " mm."
0
End If  'end if condition
                
        
End Sub


Private Sub Display_Axial_Force_Change()

'Display Axial Force
Display_Axial_Force.Value = Worksheets("Data").Range("R3") & Format(Max_Clearance, "#,##0.000") & " kN "

End Sub

Private Sub Display_Factor_Safety_Change()

'Check if the factor of Safety is less than 1
If Safety_Factor < 1 Then
    
    'Create a dialog box to show warning
    'MsgBox "WARNING DANGER, THE HUB COMPONENT WILL BURST AT THIS STRESS "

End If

'Check if the factor of Safety is graeter than 1 but less than two
If Safety_Factor > 1 And Safety_Factor < 2 Then
    
    'Create a dialog box to show warning
    MsgBox "THE FACTOR OF SAFETY IS," & Worksheets("Data").Range("P11").Value & " BY CLICKING OK YOU AGREE TO PROCEED EITH THIS VALUE."

End If

'Display factor of safety
Display_Factor_Safety.Value = Worksheets("Data").Range("P11").Value & Format(Max_Clearance, "#,##0.0")

End Sub

Private Sub Display_Nominal_Inner_Diameter_Change()

'Display Nominal inner diameter
Display_Nominal_Inner_Diameter.Value = Worksheets("Data").Range("H3") & Format(Max_Clearance, "#,##0.00") & " mm "

End Sub

Private Sub Display_Nominal_Shaft_Diameter_Change()

'Display Nominal shaft diameter
Display_Nominal_Shaft_Diameter.Value = Worksheets("Data").Range("H4") & Format(Max_Clearance, "#,##0.00") & " mm "

End Sub

Private Sub Display_Hub_Outer_Diameter_Change()

'Display Hub outer diameter
Display_Hub_Outer_Diameter.Value = Worksheets("Data").Range("H5") & Format(Max_Clearance, "#,##0.00") & " mm "

End Sub

Private Sub Display_Torque_Change()

'Display Torque
Display_Torque.Value = Worksheets("Data").Range("R4") & Format(Max_Clearance, "#,##0.00") & " kNm "

End Sub

Private Sub Select_Type_of_Fit_Change() 'This sub enables selection of the type of fit

Dim Type_Fit As String  'Declare the variable Type_Fit of type string

'initialise contact diameter variable with their respective values from the data worksheet
Contact_Diameter = Worksheets("Data").Range("H4")

'initialise type of fit variable with their respective values from the data worksheet
Type_Fit = Worksheets("Data").Range("A7")

'If contact diameter is greater than 50 but less or equal to 65, then:
If Contact_Diameter > 50 And Contact_Diameter <= 65 Then
    
    Select Case Type_Fit
    
        Case "Easy Running" 'If Easy Running selected
            'Initialise the upper tolerance of the hole variable with the value in cell D17 of the Data worksheet
            Hole_Tol_Upper = Worksheets("Data").Range("D17")
            
            'Initialise the lower tolerance of the hole variable with the value in cell E17 of the Data worksheet
            Hole_Tol_Lower = Worksheets("Data").Range("E17")
            
            'Initialise the upper tolerance of the shaft variable with the value in cell D18 of the Data worksheet
            Shaft_Tol_Upper = Worksheets("Data").Range("D18")
            
            'Initialise the lower tolerance of the shaft variable with the value in cell E18 of the Data worksheet
            Shaft_Tol_Lower = Worksheets("Data").Range("E18")
                                                        
        Case "Precision Location"   'If Precision Location case selected
            
            'Initialise the upper tolerance of the hole variable with the value in cell D19 of the Data worksheet
            Hole_Tol_Upper = Worksheets("Data").Range("D19")
            
            'Initialise the lower tolerance of the hole variable with the value in cell E19 of the Data worksheet
            Hole_Tol_Lower = Worksheets("Data").Range("E19")
            
            'Initialise the upper tolerance of the shaft variable with the value in cell D20 of the Data worksheet
            Shaft_Tol_Upper = Worksheets("Data").Range("D20")
            
            'Initialise the lower tolerance of the shaft variable with the value in cell E20 of the Data worksheet
            Shaft_Tol_Lower = Worksheets("Data").Range("E20")
                            
        Case "Press Fit"    'If Press Fit case selected
            
            'Initialise the upper tolerance of the hole variable with the value in cell D21 of the Data worksheet
            Hole_Tol_Upper = Worksheets("Data").Range("D21")
            
            'Initialise the lower tolerance of the hole variable with the value in cell E21 of the Data worksheet
            Hole_Tol_Lower = Worksheets("Data").Range("E21")
            
            'Initialise the upper tolerance of the shaft variable with the value in cell D22 of the Data worksheet
            Shaft_Tol_Upper = Worksheets("Data").Range("D22")
            
            'Initialise the lower tolerance of the shaft variable with the value in cell E22 of the Data worksheet
            Shaft_Tol_Lower = Worksheets("Data").Range("E22")
                            
        Case "Heavy Press Fit"  'If Heavy Press Fit case selected
            
            'Initialise the upper tolerance of the hole variable with the value in cell D23 of the Data worksheet
            Hole_Tol_Upper = Worksheets("Data").Range("D23")
            
            'Initialise the lower tolerance of the hole variable with the value in cell E23 of the Data worksheet
            Hole_Tol_Lower = Worksheets("Data").Range("E23")
            
            'Initialise the upper tolerance of the shaft variable with the value in cell D24 of the Data worksheet
            Shaft_Tol_Upper = Worksheets("Data").Range("D24")
            
            'Initialise the lower tolerance of the shaft variable with the value in cell E24 of the Data worksheet
            Shaft_Tol_Lower = Worksheets("Data").Range("E24")
    
    End Select  'End selecting case of type fit
    
End If  'end if condition

'If contact diameter is greater than 65 but less of equal to 80, then:
If Contact_Diameter > 65 And Contact_Diameter <= 80 Then
    
    Select Case Type_Fit
        
        Case "Easy Running" 'If Easy Running case selected
            
            'Initialise the upper tolerance of the hole variable with the value in cell G17 of the Data worksheet
            Hole_Tol_Upper = Worksheets("Data").Range("G17")
            
            'Initialise the lower tolerance of the hole variable with the value in cell H17 of the Data worksheet
            Hole_Tol_Lower = Worksheets("Data").Range("H17")
            
            'Initialise the upper tolerance of the shaft variable with the value in cell G18 of the Data worksheet
            Shaft_Tol_Upper = Worksheets("Data").Range("G18")
            
            'Initialise the lower tolerance of the shaft variable with the value in cell H18 of the Data worksheet
            Shaft_Tol_Lower = Worksheets("Data").Range("H18")
                                                        
        Case "Precision Location"   'If Precision Location case selected
            
            'Initialise the upper tolerance of the hole variable with the value in cell G19 of the Data worksheet
            Hole_Tol_Upper = Worksheets("Data").Range("G19")
            
            'Initialise the lower tolerance of the hole variable with the value in cell H19 of the Data worksheet
            Hole_Tol_Lower = Worksheets("Data").Range("H19")
            
            'Initialise the upper tolerance of the shaft variable with the value in cell G20 of the Data worksheet
            Shaft_Tol_Upper = Worksheets("Data").Range("G20")
            
            'Initialise the lower tolerance of the shaft variable with the value in cell H20 of the Data worksheet
            Shaft_Tol_Lower = Worksheets("Data").Range("H20")
                            
        Case "Press Fit"    'If Press Fit case selected'
            
            'Initialise the upper tolerance of the hole variable with the value in cell G21 of the Data worksheet
            Hole_Tol_Upper = Worksheets("Data").Range("G21")
            
            'Initialise the lower tolerance of the hole variable with the value in cell H21 of the Data worksheet
            Hole_Tol_Lower = Worksheets("Data").Range("H21")
            
            'Initialise the upper tolerance of the shaft variable with the value in cell G22 of the Data worksheet
            Shaft_Tol_Upper = Worksheets("Data").Range("G22")
            
            'Initialise the lower tolerance of the shaft variable with the value in cell H22 of the Data worksheet
            Shaft_Tol_Lower = Worksheets("Data").Range("H22")
                            
        Case "Heavy Press Fit"  'If Heavy Press Fit case selected
            
            'Initialise the upper tolerance of the hole variable with the value in cell G23 of the Data worksheet
            Hole_Tol_Upper = Worksheets("Data").Range("G23")
            
            'Initialise the lower tolerance of the hole variable with the value in cell H23 of the Data worksheet
            Hole_Tol_Lower = Worksheets("Data").Range("H23")
            
            'Initialise the upper tolerance of the shaft variable with the value in cell G24 of the Data worksheet
            Shaft_Tol_Upper = Worksheets("Data").Range("G24")
            
            'Initialise the lower tolerance of the shaft variable with the value in cell H24 of the Data worksheet
            Shaft_Tol_Lower = Worksheets("Data").Range("H24")
    
    End Select  'End selecting type of fit

End If  'end if condition


'If contact diameter is greater than 80 but less of equal to 100, then:
If Contact_Diameter > 80 And Contact_Diameter <= 100 Then
    
    Select Case Type_Fit
        
        Case "Easy Running" 'If Easy Running case selected
            
            'Initialise the upper tolerance of the hole variable with the value in cell J17 of the Data worksheet
            Hole_Tol_Upper = Worksheets("Data").Range("J17")
            
            'Initialise the lower tolerance of the hole variable with the value in cell K17 of the Data worksheet
            Hole_Tol_Lower = Worksheets("Data").Range("K17")
            
            'Initialise the upper tolerance of the shaft variable with the value in cell J18 of the Data worksheet
            Shaft_Tol_Upper = Worksheets("Data").Range("J18")
            
            'Initialise the lower tolerance of the shaft variable with the value in cell K18 of the Data worksheet
            Shaft_Tol_Lower = Worksheets("Data").Range("K18")
                                                        
        Case "Precision Location"   'If Precision Location case selected
            
            'Initialise the upper tolerance of the hole variable with the value in cell J19 of the Data worksheet
            Hole_Tol_Upper = Worksheets("Data").Range("J19")
            
            'Initialise the lower tolerance of the hole variable with the value in cell K19 of the Data worksheet
            Hole_Tol_Lower = Worksheets("Data").Range("K19")
            
            'Initialise the upper tolerance of the shaft variable with the value in cell J20 of the Data worksheet
            Shaft_Tol_Upper = Worksheets("Data").Range("J20")
            
            'Initialise the lower tolerance of the shaft variable with the value in cell K20 of the Data worksheet
            Shaft_Tol_Lower = Worksheets("Data").Range("K20")
                            
        Case "Press Fit"    'If Press Fit case selected'
            
            'Initialise the upper tolerance of the hole variable with the value in cell J21 of the Data worksheet
            Hole_Tol_Upper = Worksheets("Data").Range("J21")
            
            'Initialise the lower tolerance of the hole variable with the value in cell K21 of the Data worksheet
            Hole_Tol_Lower = Worksheets("Data").Range("K21")
            
            'Initialise the upper tolerance of the shaft variable with the value in cell J22 of the Data worksheet
            Shaft_Tol_Upper = Worksheets("Data").Range("J22")
            
            'Initialise the lower tolerance of the shaft variable with the value in cell K22 of the Data worksheet
            Shaft_Tol_Lower = Worksheets("Data").Range("K22")
                            
        Case "Heavy Press Fit"  'If Heavy Press Fit case selected
            
            'Initialise the upper tolerance of the hole variable with the value in cell J23 of the Data worksheet
            Hole_Tol_Upper = Worksheets("Data").Range("J23")
            
            'Initialise the lower tolerance of the hole variable with the value in cell K23 of the Data worksheet
            Hole_Tol_Lower = Worksheets("Data").Range("K23")
            
            'Initialise the upper tolerance of the shaft variable with the value in cell J24 of the Data worksheet
            Shaft_Tol_Upper = Worksheets("Data").Range("J24")
            
            'Initialise the lower tolerance of the shaft variable with the value in cell K24 of the Data worksheet
            Shaft_Tol_Lower = Worksheets("Data").Range("K24")
    
    End Select  'End selecting type of fit

End If  'end if condition


'If contact diameter is greater than 100 but less of equal to 120, then:
If Contact_Diameter > 100 And Contact_Diameter <= 120 Then
    
    Select Case Type_Fit
        
        Case "Easy Running" 'If Easy Running case selected
            
            'Initialise the upper tolerance of the hole variable with the value in cell M17 of the Data worksheet
            Hole_Tol_Upper = Worksheets("Data").Range("M17")
            
            'Initialise the lower tolerance of the hole variable with the value in cell N17 of the Data worksheet
            Hole_Tol_Lower = Worksheets("Data").Range("N17")
            
            'Initialise the upper tolerance of the shaft variable with the value in cell M18 of the Data worksheet
            Shaft_Tol_Upper = Worksheets("Data").Range("M18")
            
            'Initialise the lower tolerance of the shaft variable with the value in cell N18 of the Data worksheet
            Shaft_Tol_Lower = Worksheets("Data").Range("N18")
                                                        
        Case "Precision Location"   'If Precision Location case selected
            
            'Initialise the upper tolerance of the hole variable with the value in cell M19 of the Data worksheet
            Hole_Tol_Upper = Worksheets("Data").Range("M19")
            
            'Initialise the lower tolerance of the hole variable with the value in cell N19 of the Data worksheet
            Hole_Tol_Lower = Worksheets("Data").Range("N19")
            
            'Initialise the upper tolerance of the shaft variable with the value in cell M20 of the Data worksheet
            Shaft_Tol_Upper = Worksheets("Data").Range("M20")
            
            'Initialise the lower tolerance of the shaft variable with the value in cell N20 of the Data worksheet
            Shaft_Tol_Lower = Worksheets("Data").Range("N20")
                            
        Case "Press Fit"    'If Press Fit case selected'
            
            'Initialise the upper tolerance of the hole variable with the value in cell M21 of the Data worksheet
            Hole_Tol_Upper = Worksheets("Data").Range("M21")
            
            'Initialise the lower tolerance of the hole variable with the value in cell N21 of the Data worksheet
            Hole_Tol_Lower = Worksheets("Data").Range("N21")
            
            'Initialise the upper tolerance of the shaft variable with the value in cell M22 of the Data worksheet
            Shaft_Tol_Upper = Worksheets("Data").Range("M22")
            
            'Initialise the lower tolerance of the shaft variable with the value in cell N22 of the Data worksheet
            Shaft_Tol_Lower = Worksheets("Data").Range("N22")
                            
        Case "Heavy Press Fit"  'If Heavy Press Fit case selected
            
            'Initialise the upper tolerance of the hole variable with the value in cell M23 of the Data worksheet
            Hole_Tol_Upper = Worksheets("Data").Range("M23")
            
            'Initialise the lower tolerance of the hole variable with the value in cell N23 of the Data worksheet
            Hole_Tol_Lower = Worksheets("Data").Range("N23")
            
            'Initialise the upper tolerance of the shaft variable with the value in cell M24 of the Data worksheet
            Shaft_Tol_Upper = Worksheets("Data").Range("M24")
            
            'Initialise the lower tolerance of the shaft variable with the value in cell N24 of the Data worksheet
            Shaft_Tol_Lower = Worksheets("Data").Range("N24")
    
    End Select  'End selecting type of fit

End If  'end if condition

'Define the format of upper tolerance of hole to three decimal places
Worksheets("Data").Range("H6") = Format(Hole_Tol_Upper, "#,##0.000")

'Define the format of lower tolerance of hole to three decimal places
Worksheets("Data").Range("H7") = Format(Hole_Tol_Lower, "#,##0.000")

'Define the format of upper tolerance of shaft to three decimal places
Worksheets("Data").Range("H8") = Format(Shaft_Tol_Upper, "#,##0.000")

'Define the format of lower tolerance of shaft to three decimal places
Worksheets("Data").Range("H9") = Format(Shaft_Tol_Lower, "#,##0.000")


End Sub

Private Sub Display_Fit_Type_Change()

'Display type of fit (Clearance or Interference)
Display_Fit_Type.Text = Worksheets("Data").Range("A7") & " ( " & Worksheets("Data").Range("A8") & " ) "

End Sub

Private Sub Select_Shaft_Material_Change()  'This sub enables selection of shaft material change

Dim Shaft_Mat, Hub_Mat As String    'Declare the variables Shaft_Mat and Hub_Mat of type string
Dim Ei, Vi, Si As Single    'Declare variables Ei, Vi and Si of type single

'Initialise the shaft material variable with the value in cell B4 of the Data worksheet
Shaft_Mat = Worksheets("Data").Range("B4")

Select Case Shaft_Mat   'Select shaft material
        
        Case "Aluminium"    'If Aluminium selected
            
            'Initialise the variable Ei with the value in cell B11 of the Data worksheet
            Ei = Worksheets("Data").Range("B11")
            
            'Initialise the variable Vi with the value in cell B12 of the Data worksheet
            Vi = Worksheets("Data").Range("B12")
            
            'Initialise the variable Si with the value in cell B13 of the Data worksheet
            Si = Worksheets("Data").Range("B13")
                                                        
        Case "Bronze"   'If Bronze selected
            
            'Initialise the variable Ei with the value in cell C11 of the Data worksheet
            Ei = Worksheets("Data").Range("C11")
            
            'Initialise the variable Vi with the value in cell C12 of the Data worksheet
            Vi = Worksheets("Data").Range("C12")
            
            'Initialise the variable Si with the value in cell C13 of the Data worksheet
            Si = Worksheets("Data").Range("C13")
                            
        Case "Nickel Steel" 'If Nickel Steel selected
            
            'Initialise the variable Ei with the value in cell D11 of the Data worksheet
            Ei = Worksheets("Data").Range("D11")
            
            'Initialise the variable Vi with the value in cell D12 of the Data worksheet
            Vi = Worksheets("Data").Range("D12")
            
            'Initialise the variable Si with the value in cell D13 of the Data worksheet
            Si = Worksheets("Data").Range("D13")
            
        Case "Stainless Steel"  'If Stainless Steel selected
            
            'Initialise the variable Ei with the value in cell E11 of the Data worksheet
            Ei = Worksheets("Data").Range("E11")
            
            'Initialise the variable Vi with the value in cell E12 of the Data worksheet
            Vi = Worksheets("Data").Range("E12")
            
            'Initialise the variable Si with the value in cell E13 of the Data worksheet
            Si = Worksheets("Data").Range("E13")
            
        Case "Carbon Steel" 'If Carbon Steel selected
            
            'Initialise the variable Ei with the value in cell F11 of the Data worksheet
            Ei = Worksheets("Data").Range("F11")
            
            'Initialise the variable Vi with the value in cell F12 of the Data worksheet
            Vi = Worksheets("Data").Range("F12")
            
            'Initialise the variable Si with the value in cell F13 of the Data worksheet
            Si = Worksheets("Data").Range("F13")

End Select  'End selecting shaft material

   'Assign macro Ei in cell C4 of the data worksheet
   Worksheets("Data").Range("C4") = Ei
   
   'Assign macro Vi in cell D4 of the data worksheet
   Worksheets("Data").Range("D4") = Vi
   
   'Assign macro Si in cell E4 of the data worksheet
   Worksheets("Data").Range("E4") = Si
   

End Sub

Private Sub Select_Hub_Material_Change()    'This sub enables selection of hub material change

Dim Shaft_Mat, Hub_Mat As String    'Declare the variables Shaft_Mat and Hub_Mat of type string
Dim Ei, Vi As Single    'Declare variables Ei, Vi and Si of type single

'Initialise the hub material variable with the value in cell B3 of the Data worksheet
Hub_Mat = Worksheets("Data").Range("B3")

Select Case Hub_Mat 'Select hub material
        
        Case "Aluminium"    'If Aluminium selected
            
            'Initialise the variable Eo with the value in cell B11 of the Data worksheet
            Eo = Worksheets("Data").Range("B11")
            
            'Initialise the variable Vo with the value in cell B12 of the Data worksheet
            Vo = Worksheets("Data").Range("B12")
            
            'Initialise the variable So with the value in cell B13 of the Data worksheet
            So = Worksheets("Data").Range("B13")
                                                        
        Case "Bronze"   'If Bronze selected
            
            'Initialise the variable Eo with the value in cell C11 of the Data worksheet
            Eo = Worksheets("Data").Range("C11")
            
            'Initialise the variable Vo with the value in cell C12 of the Data worksheet
            Vo = Worksheets("Data").Range("C12")
            
            'Initialise the variable So with the value in cell C13 of the Data worksheet
            So = Worksheets("Data").Range("C13")
                            
        Case "Nickel Steel" 'If Nickel Steel selected
            
            'Initialise the variable Eo with the value in cell D11 of the Data worksheet
            Eo = Worksheets("Data").Range("D11")
            
            'Initialise the variable Vo with the value in cell D12 of the Data worksheet
            Vo = Worksheets("Data").Range("D12")
            
            'Initialise the variable So with the value in cell D13 of the Data worksheet
            So = Worksheets("Data").Range("D13")
            
        Case "Stainless Steel"  'If Stainless Steel selected
            
            'Initialise the variable Eo with the value in cell E11 of the Data worksheet
            Eo = Worksheets("Data").Range("E11")
            
            'Initialise the variable Vo with the value in cell E12 of the Data worksheet
            Vo = Worksheets("Data").Range("E12")
            
            'Initialise the variable So with the value in cell E13 of the Data worksheet
            So = Worksheets("Data").Range("E13")
            
        Case "Carbon Steel" 'If Carbon Steel selected
            
            'Initialise the variable Eo with the value in cell F11 of the Data worksheet
            Eo = Worksheets("Data").Range("F11")
            
            'Initialise the variable Vo with the value in cell F12 of the Data worksheet
            Vo = Worksheets("Data").Range("F12")
            
            'Initialise the variable So with the value in cell F13 of the Data worksheet
            So = Worksheets("Data").Range("F13")

End Select  'End selecting hub material
    
    'Assign macro Eo in cell C3 of the data worksheet
    Worksheets("Data").Range("C3") = Eo
    
    'Assign macro Vo in cell D3 of the data worksheet
    Worksheets("Data").Range("D3") = Vo
    
    'Assign macro So in cell E3 of the data worksheet
    Worksheets("Data").Range("E3") = So

End Sub

Private Sub Enter_Diameters_Click() 'Must be the same as the object name. This Sub enables the user to input the shaft diameters

    'Declare the variables as single floating point numbers
    Dim Shaft_Dia_Outer, Shaft_Dia_Inner, Hub_Dia_Outer, Contact_Diameter As Single

    'Intitialise the variables to zero
    Shaft_Dia_Outer = 0
    Shaft_Dia_Inner = 0
    Hub_Dia_Outer = 0

    'Read Operation - Enter the value of the variables
    Shaft_Dia_Outer = InputBox("Units mm", "Enter Shaft Outer Diameter")

    'Logical Operation - Make sure values are within the range of the programme
    Do While Shaft_Dia_Outer < 50 Or Shaft_Dia_Outer > 120
    
        'Display message - Shaft diameter is outside the range of the programme
        MsgBox "The shaft diameter entered, " & (Shaft_Dia_Outer) & "mm, is outside the range of this programme, Please enter a value between 50mm and 120 mm"
        
        'Prompt Read Operation again - Enter the value of the variables
        Shaft_Dia_Outer = InputBox("Units mm", "Enter Shaft Outer Diameter")
        
    Loop    'Reapeat until the values are within the range of the programme
    
        'Write Operation - Store the value entered by the user into a cell, on the worksheet named Data
        Worksheets("Data").Range("H4") = Shaft_Dia_Outer
        Worksheets("Data").Range("O30") = Shaft_Dia_Outer & " mm"
    
   
        'Read Operation - Enter Shaft Inner Diameter
        Shaft_Dia_Inner = InputBox("Units mm", "Enter Shaft Inner Diameter")
        
        'Write Operation - Store the value of contact diameter entered by the user into cell H4, on the worksheet named Data
        Contact_Diameter = Worksheets("Data").Range("H4")
    
    'Logical Operation - Make sure the value of shaft inner diameter is not very close to or does not exceed the shaft nominal diameter
    Do While Shaft_Dia_Inner > (Contact_Diameter * 0.8)
         
        'Display message - Shaft inner diameter is is close to or exceeds the nominal shaft diameter
        MsgBox "The shaft inner diameter entered, " & (Shaft_Dia_Inner) & "mm, is very close to or exceeds the nominal shaft diameter of " & (Shaft_Dia_Outer) & "mm, Please enter a safe diameter"
        
        'Prompt Read Operation again - Enter New Shaft Inner Diameter
        Shaft_Dia_Inner = InputBox("Units mm", "Enter New Shaft Inner Diameter")
        
    Loop    'Reapeat until shaft inner diameter is not very close to or does not exceed the shaft nominal diameter
    
        'Write Operation - Store the value entered by the user into a cell, on the worksheet named Data
        Worksheets("Data").Range("H3") = Shaft_Dia_Inner
        Worksheets("Data").Range("O31") = Shaft_Dia_Inner & " mm"
    
        'Read Operation - Enter Hub Outer Diameter
        Hub_Dia_Outer = InputBox("Units mm", "Enter Hub Outer Diameter")
        
        'Write Operation - Store the value of contact diameter entered by the user into cell H4, on the worksheet named Data
        Contact_Diameter = Worksheets("Data").Range("H4")
    
    'Logical Operation - Make sure the value of hub outer diamter is greater than 1.5 times the nominal shaft outer diameter
    Do While Hub_Dia_Outer < (Contact_Diameter * 1.5)
        
        'Display message - The hub diameter entered is less than 1.5 times the nominal shaft outer diameter
        MsgBox "The hub diameter entered, " & (Hub_Dia_Outer) & "mm, is less than 1.5 times the nominal shaft diameter of " & (Shaft_Dia_Outer) & "mm, Please enter a safe hub diameter"
        
        'Prompt Read Operation again - Enter New Hub Outer Diameter
        Hub_Dia_Outer = InputBox("Units mm", "Enter New Hub Outer Diameter")
    
    Loop    'Reapeat until hub outer diamter is greater than 1.5 times the nominal shaft outer diameter
    
        'Write Operation - Store the value entered by the user into a cell, on the worksheet named Data
        Worksheets("Data").Range("H5") = Hub_Dia_Outer
        Worksheets("Data").Range("O32") = Hub_Dia_Outer & " mm"


End Sub


Private Sub Calculate_Stresses_Click()  'This sub calculates the stresses when the comand button is clicked

'Declare the variables as single floating point numbers
Dim T1, T2, T3, T4 As Single
Dim Dc, Di, Dout, Vo, Vi, Eo, Ei, Pc_max, Pc_min As Single
Dim Max_Stress, Min_Stress As Single

'Initialise the variables with the value a cell of the worksheet named data
Max_Interferance = Worksheets("Data").Range("M6")
Min_Interferance = Worksheets("Data").Range("M7")

Dc = Worksheets("Data").Range("H4")
Di = Worksheets("Data").Range("H3")
Dout = Worksheets("Data").Range("H5")

Eo = Worksheets("Data").Range("C3")
Vo = Worksheets("Data").Range("D3")
Ei = Worksheets("Data").Range("C4")
Vi = Worksheets("Data").Range("D4")

'Define T1, T2, T3, and T4
'/1000 brings GN/m^2 to N/mm^2

T1 = (Dc * Dc) + (Di * Di) / ((Ei / 1000) * ((Dc * Dc) - (Di * Di)))

T2 = (Dout * Dout) + (Dc * Dc) / ((Eo / 1000) * ((Dout * Dout) - (Dc * Dc)))

T3 = -(Vi / (Ei))

T4 = (Vo / (Eo))

'Define Pc_max and Pc-min
Pc_max = Max_Interferance / Dc * (T1 + T2 + T3 + T4)

Pc_min = Min_Interferance / Dc * (T1 + T2 + T3 + T4)

'Store the values of Pc_max and Pc_min in a cell on the worksheet named Data
Worksheets("Data").Range("M9") = Pc_max
Worksheets("Data").Range("M10") = Pc_min

'Define Max_Stress and Min_Stress
Max_Stress = (Pc_max) * (((((Dout * Dout) + (Dc * Dc))) / (((Dout * Dout) - (Dc * Dc)))) + Vo)

Min_Stress = (Pc_min) * (((((Dout * Dout) + (Dc * Dc))) / (((Dout * Dout) - (Dc * Dc)))) + Vo)

'Format the values of maximum abnd minimum stress to 1 decimal place
Worksheets("Data").Range("M11") = Format(Max_Stress, "#,##0.0")
Worksheets("Data").Range("M12") = Format(Min_Stress, "#,##0.0")

Worksheets("Data").Range("Q30") = "Max is  " & Format(Max_Stress, "#,##0.0") & " MN/m^2"
Worksheets("Data").Range("Q31") = "Min is  " & Format(Min_Stress, "#,##0.0") & " MN/m^2"

Hub_Mat_Yield_Stress = Worksheets("Data").Range("E3")

'Define factor of safety
Factor_of_Safety = Hub_Mat_Yield_Stress / Max_Stress

'Format the value of factor of safety to 1 decimal place
Worksheets("Data").Range("P11") = Format(Factor_of_Safety, "#,##0.0")

'Display message - The maximum and minimum stress
MsgBox "The Stress in the Hub with the Maximum Stress is, " & Format(Max_Stress, "#,##0.000") & " MN/m^2. " & " The Streess in the Hub with the Minimum Stress is, " & Format(Min_Stress, "#,##0.000") & " MN/m^2." & " The Factor of Safety for the Hub Material is, " & Format(Factor_of_Safety, "#,##0.0")


End Sub

Private Sub Max_Stress_Change() 'This Sub is for the maximum stress change

'Display Maximum stress
Max_Stress.Value = Worksheets("Data").Range("M11").Value & Format(Max_Clearance, "#,##0.000") & " MN/m^2 "

End Sub

Private Sub Min_Stress_Change() 'This Sub is for the minimum stress change

'Display Minimum stress
Min_Stress.Value = Worksheets("Data").Range("M12").Value & Format(Max_Clearance, "#,##0.000") & " MN/m^2  "

End Sub


Private Sub Axial_Force_Click() ' Thuis sub calculates axial force when a comand button is clicked

Dim Axial_Force As Single   ''Declare the variable as single floating point numbers

'Store the values in a cell on the worksheet named Data
Max_Contact_Pressure = Worksheets("Data").Range("M9")
Min_Contact_Pressure = Worksheets("Data").Range("M10")
Contact_Diameter = Worksheets("Data").Range("H4")

'Read Operation - Enter the contact length of the shaft and the hub
Contact_Length = InputBox("Units mm", "Enter the contact length of the shaft and the hub")

'Read Operation - Enter the Coefficient of Friction between the shaft and the hub
Coeff_Friction = InputBox("", "Enter the Coefficient of Friction between the shaft and the hub")

'Define Pi
Pi = 3.14159

'Define maximum and minimum axial force
Max_Axial_Force = (Pi * Contact_Diameter * Contact_Length * Coeff_Friction * Max_Contact_Pressure) / 1000

Min_Axial_Force = (Pi * Contact_Diameter * Contact_Length * Coeff_Friction * Min_Contact_Pressure) / 1000

'Fromat the value of Maximum axial force to a whole number
Worksheets("Data").Range("R3") = Format(Max_Axial_Force, "#,##0")

'Define maximum torque
Max_Torque = Min_Axial_Force * 1000 * (Contact_Diameter * 0.001 / 2)

'Fromat the value of Maximum torque to a whole number
Worksheets("Data").Range("R4") = Format(Max_Torque, "#,##0")

'Display message - The Maximum Axial Force to Assemble the Components and Maximum Transmissable Torque with this Assembly
MsgBox "The Maximum Axial Force to Assemble the Components is, " & Format(Max_Axial_Force, "#,##0") & " kN " & " The Maximum Transmissable Torque with this Assembly is, " & Format(Max_Torque, "#,##0") & " Nm"

End Sub