Global targetDC As String

Sub T_SetTarget()
'Change this value if you want to target a SPECIFIC DC or you are proxying macro communication
targetDC = ""
End Sub

Sub A1_BuildReconWorksheets()

If Not (Evaluate("ISREF('ADSubnets'!A1)")) Then
    Sheets.Add.Name = "ADSubnets"
    
    Sheets("ADSubnets").Range("A1").ColumnWidth = 50
    Sheets("ADSubnets").Range("B1").ColumnWidth = 20
    
    Sheets("ADSubnets").Cells(1, 1).Value = "Name"
    Sheets("ADSubnets").Cells(1, 2).Value = "SiteObject"
    
    Sheets("ADSubnets").Range("A1:B1").Font.Bold = True
    Sheets("ADSubnets").Range("A1:B1").Interior.ColorIndex = 15
    
End If

If Not (Evaluate("ISREF('ADSites'!A1)")) Then
    Sheets.Add.Name = "ADSites"
    
    Sheets("ADSites").Range("A1").ColumnWidth = 50
    Sheets("ADSites").Range("B1").ColumnWidth = 20
    
    Sheets("ADSites").Cells(1, 1).Value = "Name"
    Sheets("ADSites").Cells(1, 2).Value = "SiteObjectBL"
    
    Sheets("ADSites").Range("A1:B1").Font.Bold = True
    Sheets("ADSites").Range("A1:B1").Interior.ColorIndex = 15
    
End If

If Not (Evaluate("ISREF('TrustedDomains'!A1)")) Then
    Sheets.Add.Name = "TrustedDomains"
    
    Sheets("TrustedDomains").Range("A1").ColumnWidth = 50
    Sheets("TrustedDomains").Range("B1").ColumnWidth = 15
    Sheets("TrustedDomains").Range("C1").ColumnWidth = 15
    Sheets("TrustedDomains").Range("D1").ColumnWidth = 15
    Sheets("TrustedDomains").Range("E1").ColumnWidth = 15
    
    Sheets("TrustedDomains").Cells(1, 1).Value = "Name"
    Sheets("TrustedDomains").Cells(1, 2).Value = "TrustAttributes"
    Sheets("TrustedDomains").Cells(1, 3).Value = "TrustDirection"
    Sheets("TrustedDomains").Cells(1, 4).Value = "TrustPartner"
    Sheets("TrustedDomains").Cells(1, 5).Value = "TrustType"
    
    Sheets("TrustedDomains").Range("A1:E1").Font.Bold = True
    Sheets("TrustedDomains").Range("A1:E1").Interior.ColorIndex = 15
    
End If

If Not (Evaluate("ISREF('DFSRoot'!A1)")) Then
    Sheets.Add.Name = "DFSRoot"
    
    Sheets("DFSRoot").Range("A1").ColumnWidth = 50
    
    Sheets("DFSRoot").Cells(1, 1).Value = "DFSRoot"
    
    Sheets("DFSRoot").Range("A1").Font.Bold = True
    Sheets("DFSRoot").Range("A1").Interior.ColorIndex = 15
    
End If

If Not (Evaluate("ISREF('QueryGPO'!A1)")) Then
    Sheets.Add.Name = "QueryGPO"
    
    Sheets("QueryGPO").Range("A1").ColumnWidth = 20
    Sheets("QueryGPO").Range("B1").ColumnWidth = 120
    Sheets("QueryGPO").Range("C1").ColumnWidth = 20
    
    Sheets("QueryGPO").Cells(1, 1).Value = "Resource Type"
    Sheets("QueryGPO").Cells(1, 2).Value = "Policy Location/UNC Path/Shortcut URL"
    Sheets("QueryGPO").Cells(1, 3).Value = "Filename/Drive Letter"
    
    Sheets("QueryGPO").Range("A1:C1").Font.Bold = True
    Sheets("QueryGPO").Range("A1:C1").Interior.ColorIndex = 15
    
End If

If Not (Evaluate("ISREF('QuerySQL'!A1)")) Then
    Sheets.Add.Name = "QuerySQL"
    
    Sheets("QuerySQL").Range("A1").ColumnWidth = 20
    Sheets("QuerySQL").Range("B1").ColumnWidth = 18
    Sheets("QuerySQL").Range("C1").ColumnWidth = 3
    Sheets("QuerySQL").Range("D1").ColumnWidth = 6
    Sheets("QuerySQL").Range("E1").ColumnWidth = 35
    Sheets("QuerySQL").Range("F1").ColumnWidth = 19
    Sheets("QuerySQL").Range("G1").ColumnWidth = 19
    Sheets("QuerySQL").Range("H1").ColumnWidth = 11
    Sheets("QuerySQL").Range("I1").ColumnWidth = 50
    Sheets("QuerySQL").Range("J1").ColumnWidth = 50
    Sheets("QuerySQL").Range("K1").ColumnWidth = 17
    Sheets("QuerySQL").Range("L1").ColumnWidth = 50
    
    Sheets("QuerySQL").Cells(1, 1).Value = "ComputerName"
    Sheets("QuerySQL").Cells(1, 2).Value = "UserAccountControl"
    Sheets("QuerySQL").Cells(1, 3).Value = "DC"
    Sheets("QuerySQL").Cells(1, 4).Value = "RODC"
    Sheets("QuerySQL").Cells(1, 5).Value = "OperatingSystem"
    Sheets("QuerySQL").Cells(1, 6).Value = "PasswordLastSetHigh"
    Sheets("QuerySQL").Cells(1, 7).Value = "PasswordLastSetLow"
    Sheets("QuerySQL").Cells(1, 8).Value = "PwdLastSet"
    Sheets("QuerySQL").Cells(1, 9).Value = "Comment"
    Sheets("QuerySQL").Cells(1, 10).Value = "Description"
    Sheets("QuerySQL").Cells(1, 11).Value = "Connection Status"
    Sheets("QuerySQL").Cells(1, 12).Value = "Databases"
    
    Sheets("QuerySQL").Range("A1:L1").Font.Bold = True
    Sheets("QuerySQL").Range("A1:L1").Interior.ColorIndex = 15
    
    Columns("B").Hidden = True
    Columns("F:G").Hidden = True
    
End If

If Not (Evaluate("ISREF('QueryLAPS'!A1)")) Then
    Sheets.Add.Name = "QueryLAPS"
    
    Sheets("QueryLAPS").Range("A1").ColumnWidth = 20
    Sheets("QueryLAPS").Range("B1").ColumnWidth = 18
    Sheets("QueryLAPS").Range("C1").ColumnWidth = 3
    Sheets("QueryLAPS").Range("D1").ColumnWidth = 5
    Sheets("QueryLAPS").Range("E1").ColumnWidth = 35
    Sheets("QueryLAPS").Range("F1").ColumnWidth = 19
    Sheets("QueryLAPS").Range("G1").ColumnWidth = 19
    Sheets("QueryLAPS").Range("H1").ColumnWidth = 10
    Sheets("QueryLAPS").Range("I1").ColumnWidth = 50
    Sheets("QueryLAPS").Range("J1").ColumnWidth = 50
    Sheets("QueryLAPS").Range("K1").ColumnWidth = 30
    
    Sheets("QueryLAPS").Cells(1, 1).Value = "ComputerName"
    Sheets("QueryLAPS").Cells(1, 2).Value = "UserAccountControl"
    Sheets("QueryLAPS").Cells(1, 3).Value = "DC"
    Sheets("QueryLAPS").Cells(1, 4).Value = "RODC"
    Sheets("QueryLAPS").Cells(1, 5).Value = "OperatingSystem"
    Sheets("QueryLAPS").Cells(1, 6).Value = "PasswordLastSetHigh"
    Sheets("QueryLAPS").Cells(1, 7).Value = "PasswordLastSetLow"
    Sheets("QueryLAPS").Cells(1, 8).Value = "PwdLastSet"
    Sheets("QueryLAPS").Cells(1, 9).Value = "Comment"
    Sheets("QueryLAPS").Cells(1, 10).Value = "Description"
    Sheets("QueryLAPS").Cells(1, 11).Value = "LAPS Password"
    
    Sheets("QueryLAPS").Range("A1:K1").Font.Bold = True
    Sheets("QueryLAPS").Range("A1:K1").Interior.ColorIndex = 15
    
    Columns("B").Hidden = True
    Columns("F:G").Hidden = True
    
End If

If Not (Evaluate("ISREF('ADComputers'!A1)")) Then
    Sheets.Add.Name = "ADComputers"
    
    Sheets("ADComputers").Range("A1").ColumnWidth = 20
    Sheets("ADComputers").Range("B1").ColumnWidth = 18
    Sheets("ADComputers").Range("C1").ColumnWidth = 3
    Sheets("ADComputers").Range("D1").ColumnWidth = 5
    Sheets("ADComputers").Range("E1").ColumnWidth = 35
    Sheets("ADComputers").Range("F1").ColumnWidth = 19
    Sheets("ADComputers").Range("G1").ColumnWidth = 19
    Sheets("ADComputers").Range("H1").ColumnWidth = 10
    Sheets("ADComputers").Range("I1").ColumnWidth = 50
    Sheets("ADComputers").Range("J1").ColumnWidth = 50
    Sheets("ADComputers").Range("K1").ColumnWidth = 26.5
    
    Sheets("ADComputers").Cells(1, 1).Value = "ComputerName"
    Sheets("ADComputers").Cells(1, 2).Value = "UserAccountControl"
    Sheets("ADComputers").Cells(1, 3).Value = "DC"
    Sheets("ADComputers").Cells(1, 4).Value = "RODC"
    Sheets("ADComputers").Cells(1, 5).Value = "OperatingSystem"
    Sheets("ADComputers").Cells(1, 6).Value = "PasswordLastSetHigh"
    Sheets("ADComputers").Cells(1, 7).Value = "PasswordLastSetLow"
    Sheets("ADComputers").Cells(1, 8).Value = "PwdLastSet"
    Sheets("ADComputers").Cells(1, 9).Value = "Comment"
    Sheets("ADComputers").Cells(1, 10).Value = "Description"
    Sheets("ADComputers").Cells(1, 11).Value = "Trusted to Auth for Delegation"
    
    Sheets("ADComputers").Range("A1:K1").Font.Bold = True
    Sheets("ADComputers").Range("A1:K1").Interior.ColorIndex = 15
    
    Columns("B").Hidden = True
    Columns("F:G").Hidden = True
    
End If

If Not (Evaluate("ISREF('QueryGroup'!A1)")) Then
    Sheets.Add.Name = "QueryGroup"
    
    Sheets("QueryGroup").Cells(1, 1).Value = "GroupNameToQuery:"
    Sheets("QueryGroup").Cells(3, 1).Value = "Name:"
    Sheets("QueryGroup").Cells(4, 1).Value = "Comment:"
    Sheets("QueryGroup").Cells(5, 1).Value = "Description:"
    Sheets("QueryGroup").Cells(6, 1).Value = "Members:"
    
    Sheets("QueryGroup").Range("A1").Font.Bold = True
    Sheets("QueryGroup").Range("A1").Interior.ColorIndex = 15
    Sheets("QueryGroup").Range("A1").HorizontalAlignment = xlRight
    Sheets("QueryGroup").Range("A1").ColumnWidth = 25
    Sheets("QueryGroup").Range("A3:A6").Font.Bold = True
    Sheets("QueryGroup").Range("A3:A6").Interior.ColorIndex = 15
    Sheets("QueryGroup").Range("A3:A6").HorizontalAlignment = xlRight
    Sheets("QueryGroup").Range("A3:A6").ColumnWidth = 25
    Sheets("QueryGroup").Range("B1:B6").ColumnWidth = 100
    
End If

If Not (Evaluate("ISREF('ADGroupInfo'!A1)")) Then
    Sheets.Add.Name = "ADGroupInfo"
    
    Sheets("ADGroupInfo").Range("A1").ColumnWidth = 12
    
    Sheets("ADGroupInfo").Cells(1, 1).Value = "GroupName"
    Sheets("ADGroupInfo").Cells(2, 1).Value = "Members"
    
    Sheets("ADGroupInfo").Range("A1:A2").Font.Bold = True
    Sheets("ADGroupInfo").Range("A1:A2").Interior.ColorIndex = 15
    
    Sheets("ADGroupInfo").Range("B1:ZZ1").Interior.ColorIndex = 19

End If

If Not (Evaluate("ISREF('ADGroups'!A1)")) Then
    Sheets.Add.Name = "ADGroups"
    
    Sheets("ADGroups").Range("A1").ColumnWidth = 70
    Sheets("ADGroups").Range("B1").ColumnWidth = 50
    Sheets("ADGroups").Range("C1").ColumnWidth = 50
    Sheets("ADGroups").Range("D1").ColumnWidth = 50
    
    Sheets("ADGroups").Cells(1, 1).Value = "GroupName"
    Sheets("ADGroups").Cells(1, 2).Value = "First Member"
    Sheets("ADGroups").Cells(1, 3).Value = "Comment"
    Sheets("ADGroups").Cells(1, 4).Value = "Description"
    
    Sheets("ADGroups").Range("A1:D1").Font.Bold = True
    Sheets("ADGroups").Range("A1:D1").Interior.ColorIndex = 15
    
End If

If Not (Evaluate("ISREF('QueryUser'!A1)")) Then
    Sheets.Add.Name = "QueryUser"
    
    Sheets("QueryUser").Cells(1, 1).Value = "UsernameToQuery:"
    Sheets("QueryUser").Cells(3, 1).Value = "Name:"
    Sheets("QueryUser").Cells(4, 1).Value = "SamAccountName:"
    Sheets("QueryUser").Cells(5, 1).Value = "UserPrincipalName:"
    Sheets("QueryUser").Cells(6, 1).Value = "BadPwdCount:"
    Sheets("QueryUser").Cells(7, 1).Value = "Mail:"
    Sheets("QueryUser").Cells(8, 1).Value = "Phone:"
    Sheets("QueryUser").Cells(9, 1).Value = "Mobile:"
    Sheets("QueryUser").Cells(10, 1).Value = "Manager:"
    Sheets("QueryUser").Cells(11, 1).Value = "LogonCount:"
    Sheets("QueryUser").Cells(12, 1).Value = "MemberOf:"
    
    Sheets("QueryUser").Range("A1").Font.Bold = True
    Sheets("QueryUser").Range("A1").Interior.ColorIndex = 15
    Sheets("QueryUser").Range("A1").HorizontalAlignment = xlRight
    Sheets("QueryUser").Range("A1").ColumnWidth = 25
    Sheets("QueryUser").Range("A3:A12").Font.Bold = True
    Sheets("QueryUser").Range("A3:A12").Interior.ColorIndex = 15
    Sheets("QueryUser").Range("A3:A12").HorizontalAlignment = xlRight
    Sheets("QueryUser").Range("A3:A12").ColumnWidth = 25
    Sheets("QueryUser").Range("B1:B12").ColumnWidth = 100
    
End If

If Not (Evaluate("ISREF('ADUserUACInfo'!A1)")) Then
    Sheets.Add.Name = "ADUserUACInfo"
    
    Sheets("ADUserUACInfo").Range("A1").ColumnWidth = 25
    Sheets("ADUserUACInfo").Range("B1").ColumnWidth = 8.5
    Sheets("ADUserUACInfo").Range("C1").ColumnWidth = 8.5
    Sheets("ADUserUACInfo").Range("D1").ColumnWidth = 10
    Sheets("ADUserUACInfo").Range("E1").ColumnWidth = 10
    Sheets("ADUserUACInfo").Range("F1").ColumnWidth = 10
    Sheets("ADUserUACInfo").Range("G1").ColumnWidth = 10
    Sheets("ADUserUACInfo").Range("H1").ColumnWidth = 10
    Sheets("ADUserUACInfo").Range("I1").ColumnWidth = 10
    Sheets("ADUserUACInfo").Range("J1").ColumnWidth = 10
    Sheets("ADUserUACInfo").Range("K1").ColumnWidth = 10
    
    Sheets("ADUserUACInfo").Cells(1, 1).Value = "UserName"
    Sheets("ADUserUACInfo").Cells(1, 2).Value = "UAC"
    Sheets("ADUserUACInfo").Cells(1, 3).Value = "Disabled"
    Sheets("ADUserUACInfo").Cells(1, 4).Value = "Password Not Required"
    Sheets("ADUserUACInfo").Cells(1, 5).Value = "PreAuth Not Required"
    Sheets("ADUserUACInfo").Cells(1, 6).Value = "Reversible Encryption Allowed"
    Sheets("ADUserUACInfo").Cells(1, 7).Value = "Password Never Expires"
    Sheets("ADUserUACInfo").Cells(1, 8).Value = "Trusted  for  Delegation"
    Sheets("ADUserUACInfo").Cells(1, 9).Value = "Use DES Key Only"
    Sheets("ADUserUACInfo").Cells(1, 10).Value = "Password Expired"
    Sheets("ADUserUACInfo").Cells(1, 11).Value = "Trusted to Auth for Delegation"
    
    Sheets("ADUserUACInfo").Range("A1:K1").Font.Bold = True
    Sheets("ADUserUACInfo").Range("A1:K1").Interior.ColorIndex = 15
    Sheets("ADUserUACInfo").Range("B1:K1").HorizontalAlignment = xlCenter
    Sheets("ADUserUACInfo").Range("A1:K1").WrapText = True
    
End If

If Not (Evaluate("ISREF('ADUsers'!A1)")) Then
    Sheets.Add.Name = "ADUsers"
    
    Sheets("ADUsers").Range("A1").ColumnWidth = 25
    Sheets("ADUsers").Range("B1").ColumnWidth = 18
    Sheets("ADUsers").Range("C1").ColumnWidth = 12
    Sheets("ADUsers").Range("D1").ColumnWidth = 19
    Sheets("ADUsers").Range("E1").ColumnWidth = 19
    Sheets("ADUsers").Range("F1").ColumnWidth = 10
    Sheets("ADUsers").Range("G1").ColumnWidth = 15
    Sheets("ADUsers").Range("H1").ColumnWidth = 20
    Sheets("ADUsers").Range("I1").ColumnWidth = 20
    Sheets("ADUsers").Range("J1").ColumnWidth = 60
    Sheets("ADUsers").Range("K1").ColumnWidth = 100
    Sheets("ADUsers").Range("L1").ColumnWidth = 100
    Sheets("ADUsers").Range("M1").ColumnWidth = 50
    Sheets("ADUsers").Range("N1").ColumnWidth = 17
    Sheets("ADUsers").Range("O1").ColumnWidth = 17
    Sheets("ADUsers").Range("P1").ColumnWidth = 18
    Sheets("ADUsers").Range("Q1").ColumnWidth = 18
    Sheets("ADUsers").Range("R1").ColumnWidth = 20
    Sheets("ADUsers").Range("S1").ColumnWidth = 100
    
    Sheets("ADUsers").Cells(1, 1).Value = "UserName"
    Sheets("ADUsers").Cells(1, 2).Value = "UserAccountControl"
    Sheets("ADUsers").Cells(1, 3).Value = "Disabled"
    Sheets("ADUsers").Cells(1, 4).Value = "PasswordLastSetHigh"
    Sheets("ADUsers").Cells(1, 5).Value = "PasswordLastSetLow"
    Sheets("ADUsers").Cells(1, 6).Value = "PwdLastSet"
    Sheets("ADUsers").Cells(1, 7).Value = "UserPassword"
    Sheets("ADUsers").Cells(1, 8).Value = "UnixUserPassword"
    Sheets("ADUsers").Cells(1, 9).Value = "UnicodePassword"
    Sheets("ADUsers").Cells(1, 10).Value = "Comment"
    Sheets("ADUsers").Cells(1, 11).Value = "Description"
    Sheets("ADUsers").Cells(1, 12).Value = "Info"
    Sheets("ADUsers").Cells(1, 13).Value = "Email"
    Sheets("ADUsers").Cells(1, 14).Value = "EmployeeId"
    Sheets("ADUsers").Cells(1, 15).Value = "EmployeeNumber"
    Sheets("ADUsers").Cells(1, 16).Value = "TelephoneNumber"
    Sheets("ADUsers").Cells(1, 17).Value = "Mobile"
    Sheets("ADUsers").Cells(1, 18).Value = "Manager"
    Sheets("ADUsers").Cells(1, 19).Value = "ServicePrincipalName"
    
    Sheets("ADUsers").Range("A1:S1").Font.Bold = True
    Sheets("ADUsers").Range("A1:S1").Interior.ColorIndex = 15
    Sheets("ADUsers").Range("A1:S1").HorizontalAlignment = xlLeft
    
    Columns("B").Hidden = True
    Columns("D:E").Hidden = True
        
End If

If Not (Evaluate("ISREF('ADInfo'!A1)")) Then
    Sheets.Add.Name = "ADInfo"
    
    Sheets("ADInfo").Range("A1").ColumnWidth = 15
    Sheets("ADInfo").Range("B1").ColumnWidth = 6
    Sheets("ADInfo").Range("C1").ColumnWidth = 6
    Sheets("ADInfo").Range("D1").ColumnWidth = 5
    Sheets("ADInfo").Range("E1").ColumnWidth = 5
    Sheets("ADInfo").Range("F1").ColumnWidth = 9
    Sheets("ADInfo").Range("G1").ColumnWidth = 9
    Sheets("ADInfo").Range("H1").ColumnWidth = 5
    Sheets("ADInfo").Range("I1").ColumnWidth = 5
    Sheets("ADInfo").Range("J1").ColumnWidth = 5
    Sheets("ADInfo").Range("K1").ColumnWidth = 5
    Sheets("ADInfo").Range("L1").ColumnWidth = 10
    Sheets("ADInfo").Range("M1").ColumnWidth = 22
    
    Sheets("ADInfo").Cells(1, 1).Value = "Domain Name"
    Sheets("ADInfo").Cells(1, 2).Value = "LockoutDurationHigh"
    Sheets("ADInfo").Cells(1, 3).Value = "LockoutDurationLow"
    Sheets("ADInfo").Cells(1, 4).Value = "LockoutObservationHigh"
    Sheets("ADInfo").Cells(1, 5).Value = "LockoutObservationLow"
    Sheets("ADInfo").Cells(1, 6).Value = "LockoutThreshold"
    Sheets("ADInfo").Cells(1, 7).Value = "MinimumPasswordLength"
    Sheets("ADInfo").Cells(1, 8).Value = "MinimumPasswordAgeHigh"
    Sheets("ADInfo").Cells(1, 9).Value = "MinimumPasswordAgeLow"
    Sheets("ADInfo").Cells(1, 10).Value = "MaximumPasswordAgeHigh"
    Sheets("ADInfo").Cells(1, 11).Value = "MaximumPasswordAgeLow"

    Sheets("ADInfo").Cells(3, 2).Value = "=IF(C2 < 0, B2 +2, B2)"
    Sheets("ADInfo").Cells(3, 3).Value = "=C2"
    Sheets("ADInfo").Cells(3, 4).Value = "=IF(E2 < 0,D2+1,D2)"
    Sheets("ADInfo").Cells(3, 5).Value = "=E2"
    Sheets("ADInfo").Cells(3, 8).Value = "=IF(I2 < 0, H2+1,H2)"
    Sheets("ADInfo").Cells(3, 9).Value = "=I2"

    Sheets("ADInfo").Range("B4:C4").Merge
    Sheets("ADInfo").Range("D4:E4").Merge
    Sheets("ADInfo").Range("H4:I4").Merge
    Sheets("ADInfo").Range("J4:K4").Merge
    
    Sheets("ADInfo").Cells(4, 2).Value = "=-(((B2)*2^32) + C2)/(10000000*60)"
    Sheets("ADInfo").Cells(4, 4).Value = "=-(((D3)*2^32)+E3)/(10000000*60)"
    Sheets("ADInfo").Cells(4, 8).Value = "=-(((H3)*2^32)+I3)/(10000000 *60 *60 *24)"
    Sheets("ADInfo").Cells(4, 10).Value = "=QUOTIENT(QUOTIENT(QUOTIENT(QUOTIENT(-((J2 * 2^32) + K2),10000000),60),60),24)"
    
    Sheets("ADInfo").Range("A5:K5").Merge
    Sheets("ADInfo").Cells(5, 1).Value = "***** The Formulas Above Are Used to Calculate Stored Domain Password Policy Attribut Information - DO NOT CHANGE *****"
    Sheets("ADInfo").Range("A5").Font.Bold = True
    Sheets("ADInfo").Range("A5").Interior.ColorIndex = 6
    Sheets("ADInfo").Range("A5").HorizontalAlignment = xlCenter
    
    Rows("1:5").EntireRow.Hidden = True
        
    Sheets("ADInfo").Range("B6:C6").Merge
    Sheets("ADInfo").Range("D6:E6").Merge
    Sheets("ADInfo").Range("H6:I6").Merge
    Sheets("ADInfo").Range("J6:K6").Merge
    Sheets("ADInfo").Range("B7:C7").Merge
    Sheets("ADInfo").Range("D7:E7").Merge
    Sheets("ADInfo").Range("H7:I7").Merge
    Sheets("ADInfo").Range("J7:K7").Merge
    
    Sheets("ADInfo").Cells(6, 1).Value = "Domain Name"
    Sheets("ADInfo").Cells(6, 2).Value = "Lockout Duration"
    Sheets("ADInfo").Cells(6, 4).Value = "Lockout Observation Window"
    Sheets("ADInfo").Cells(6, 6).Value = "Lockout Threshold"
    Sheets("ADInfo").Cells(6, 7).Value = "Minimum Password Length"
    Sheets("ADInfo").Cells(6, 8).Value = "Minimum Password Age"
    Sheets("ADInfo").Cells(6, 10).Value = "Maximum Password Age"
    Sheets("ADInfo").Cells(6, 13).Value = "Machine Account Quota"
    Sheets("ADInfo").Cells(10, 13).Value = "Functional Level"
    
    Sheets("ADInfo").Range("A6:K6").Font.Bold = True
    Sheets("ADInfo").Range("A6:K6").Interior.ColorIndex = 15
    Sheets("ADInfo").Range("A6:K6").HorizontalAlignment = xlCenter
    Sheets("ADInfo").Range("A6:K6").WrapText = True

    Sheets("ADInfo").Range("M6").Font.Bold = True
    Sheets("ADInfo").Range("M6").Interior.ColorIndex = 15
    Sheets("ADInfo").Range("M6").HorizontalAlignment = xlCenter
    Sheets("ADInfo").Range("M6").WrapText = True

    Sheets("ADInfo").Range("M10").Font.Bold = True
    Sheets("ADInfo").Range("M10").Interior.ColorIndex = 15
    Sheets("ADInfo").Range("M10").HorizontalAlignment = xlCenter
    Sheets("ADInfo").Range("M10").WrapText = True
    
    Sheets("ADInfo").Cells(7, 1).Value = "=A2"
    Sheets("ADInfo").Cells(7, 2).Value = "=IF(B2=-2147483648,IF(C2=0,""Never Unlock"",B4))"
    Sheets("ADInfo").Cells(7, 4).Value = "=D4"
    Sheets("ADInfo").Cells(7, 6).Value = "=F2"
    Sheets("ADInfo").Cells(7, 7).Value = "=G2"
    Sheets("ADInfo").Cells(7, 8).Value = "=H4"
    Sheets("ADInfo").Cells(7, 10).Value = "=J4"
    
    Sheets("ADInfo").Range("A10:K10").Merge
    Sheets("ADInfo").Cells(10, 1).Value = "Fine-Grained Password Policies"
    Sheets("ADInfo").Range("A10").Font.Bold = True
    Sheets("ADInfo").Range("A10").Interior.ColorIndex = 15
    Sheets("ADInfo").Range("A10").HorizontalAlignment = xlCenter
    

End If

If Not (Evaluate("ISREF('ProgramFiles'!A1)")) Then
    Sheets.Add.Name = "ProgramFiles"
    
    Sheets("ProgramFiles").Cells(1, 1).Value = "Program Files:"
    Sheets("ProgramFiles").Cells(1, 2).Value = "Program Files (x86):"
    Sheets("ProgramFiles").Range("A1:B1").Font.Bold = True
    Sheets("ProgramFiles").Range("A1:B1").Interior.ColorIndex = 15
    Sheets("ProgramFiles").Range("A1:B1").HorizontalAlignment = xlLeft
    Sheets("ProgramFiles").Range("A1:B1").ColumnWidth = 50

End If

If Not (Evaluate("ISREF('UserInfo'!A1)")) Then
    Sheets.Add.Name = "UserInfo"
    
    Sheets("UserInfo").Cells(1, 1).Value = "Name:"
    Sheets("UserInfo").Cells(2, 1).Value = "SamAccountName:"
    Sheets("UserInfo").Cells(3, 1).Value = "UserPrincipalName:"
    Sheets("UserInfo").Cells(4, 1).Value = "BadPwdCount:"
    Sheets("UserInfo").Cells(5, 1).Value = "Mail:"
    Sheets("UserInfo").Cells(6, 1).Value = "Phone:"
    Sheets("UserInfo").Cells(7, 1).Value = "Mobile:"
    Sheets("UserInfo").Cells(8, 1).Value = "Manager:"
    Sheets("UserInfo").Cells(9, 1).Value = "LogonCount:"
    Sheets("UserInfo").Cells(10, 1).Value = "MemberOf:"
    Sheets("UserInfo").Range("A1:A10").Font.Bold = True
    Sheets("UserInfo").Range("A1:A10").Interior.ColorIndex = 15
    Sheets("UserInfo").Range("A1:A10").HorizontalAlignment = xlRight
    Sheets("UserInfo").Range("A1:A10").ColumnWidth = 25
    Sheets("UserInfo").Range("B1:B10").ColumnWidth = 100
    
End If

If Not (Evaluate("ISREF('HostInfo'!A1)")) Then
    Sheets.Add.Name = "HostInfo"
    
    Sheets("HostInfo").Cells(1, 1).Value = "User:"
    Sheets("HostInfo").Cells(2, 1).Value = "Computer:"
    Sheets("HostInfo").Cells(3, 1).Value = "Domain:"
    Sheets("HostInfo").Cells(4, 1).Value = "DNS Domain:"
    Sheets("HostInfo").Cells(5, 1).Value = "Domain Controller:"
    Sheets("HostInfo").Cells(6, 1).Value = "Operating System:"
    Sheets("HostInfo").Range("A1:A6").Font.Bold = True
    Sheets("HostInfo").Range("A1:A6").Interior.ColorIndex = 15
    Sheets("HostInfo").Range("A1:A6").HorizontalAlignment = xlRight
    Sheets("HostInfo").Range("A1:B6").ColumnWidth = 25
    
    Sheets("HostInfo").Range("A9:B9").Merge
    Sheets("HostInfo").Cells(9, 1).Value = "Running Processes:"
    Sheets("HostInfo").Range("A9").Font.Bold = True
    Sheets("HostInfo").Range("A9").Interior.ColorIndex = 15
    Sheets("HostInfo").Range("A9").HorizontalAlignment = xlCenter
    
End If

End Sub

Sub F2_HideReconWorksheets()
    Sheets("HostInfo").Visible = False
    Sheets("UserInfo").Visible = False
    Sheets("ProgramFiles").Visible = False
    Sheets("ADInfo").Visible = False
    Sheets("ADusers").Visible = False
    Sheets("ADUserUACInfo").Visible = False
    Sheets("QueryUser").Visible = False
    Sheets("ADGroups").Visible = False
    Sheets("ADGroupInfo").Visible = False
    Sheets("QueryGroup").Visible = False
    Sheets("ADComputers").Visible = False
    Sheets("DFSRoot").Visible = False
    Sheets("TrustedDomains").Visible = False
    Sheets("ADSites").Visible = False
    Sheets("ADSubnets").Visible = False
End Sub

Sub F3_UnHideReconWorksheets()
    Sheets("HostInfo").Visible = True
    Sheets("UserInfo").Visible = True
    Sheets("ProgramFiles").Visible = True
    Sheets("ADInfo").Visible = True
    Sheets("ADusers").Visible = True
    Sheets("ADUserUACInfo").Visible = True
    Sheets("QueryUser").Visible = False
    Sheets("ADGroups").Visible = True
    Sheets("ADGroupInfo").Visible = True
    Sheets("QueryGroup").Visible = False
    Sheets("ADComputers").Visible = True
    Sheets("DFSRoot").Visible = True
    Sheets("TrustedDomains").Visible = False
    Sheets("ADSites").Visible = False
    Sheets("ADSubnets").Visible = False
End Sub

Sub A2_Collect()
    A1_BuildReconWorksheets
    H1_HostInfo
    H2_ProgramFiles
    H3_UserInfo
    AD1_ADUsers
    AD2_ADInfo
    AD3_ADGroups
    AD4_ADComputers
    AD5_DFSRoot
    AD6_TrustedDomains
    AD7_ADSites
    AD8_ADSubnets
    F1_ApplyConditionalFormatting
    ActiveWorkbook.Save
End Sub

Sub H1_HostInfo()
    'Initialize Variables
    Dim strUname As String
    Dim strCname As String
    Dim strDomain As String
    Dim strDomainDC As String
    Dim strDnsDomain As String
    Dim strOS As String
    Dim strComputer As String
    Dim objServices As Object, objProcessSet As Object, Process As Object
    Dim intProcStartRow As Long
    
    'Gather User Information
    strUname = Environ$("username")
    strCname = Environ$("computername")
    strDomain = Environ$("userdomain")
    strDnsDomain = Environ$("userdnsdomain")
    strDomainDC = Environ$("logonServer")
    strOS = Environ$("os")

    'Populate target sheet
    Sheets("HostInfo").Cells(1, 2).Value = strUname
    Sheets("HostInfo").Cells(2, 2).Value = strCname
    Sheets("HostInfo").Cells(3, 2).Value = strDomain
    Sheets("HostInfo").Cells(4, 2).Value = strDnsDomain
    Sheets("HostInfo").Cells(5, 2).Value = strDomainDC
    Sheets("HostInfo").Cells(6, 2).Value = strOS

    'Gather Running Process Information
    strComputer = "."
    intProcStartRow = 10

    Set objServices = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set objProcessSet = objServices.ExecQuery("SELECT Name, ProcessID FROM Win32_Process", , 48)


    For Each Process In objProcessSet
        Sheets("HostInfo").Cells(intProcStartRow, 1).Value = Process.Properties_("ProcessID").Value
            Sheets("HostInfo").Cells(intProcStartRow, 2).Value = Process.Properties_("Name").Value
        intProcStartRow = intProcStartRow + 1
    Next
End Sub

Sub H2_ProgramFiles()
    'Initialize Variables
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oDir As Object
    Dim pfCol As Integer
    Dim pfRow As Integer

    Set oFSO = CreateObject("Scripting.FileSystemObject")

    pfCol = 1
    pfRow = 2
    Set oFolder = oFSO.GetFolder("C:\Program Files")
    For Each oDir In oFolder.SubFolders
        Sheets("ProgramFiles").Cells(pfRow, pfCol) = oDir.Name
        pfRow = pfRow + 1
    Next oDir

    pfCol = pfCol + 1
    pfRow = 2
    Set oFolder = oFSO.GetFolder("C:\Program Files (x86)")
    For Each oDir In oFolder.SubFolders
        Sheets("ProgramFiles").Cells(pfRow, pfCol) = oDir.Name
        pfRow = pfRow + 1
    Next oDir
End Sub

Sub H3_UserInfo()
    Dim objRootDSE, objConn, objCmd, objRs As Object
    Dim strUname As String
    Dim strBase As String
    Dim strFilter As String
    Dim strAttrib As String
    Dim strScope As String
    Dim intRow As Long
    Const ADS_CHASE_REFERRALS_ALWAYS = &H60
    
    strUname = Environ$("username")

    If targetDC <> "" Then
        strBase = "<LDAP://" & targetDC & ">"
    Else
        Set rootDSE = GetObject("LDAP://RootDSE")
        strBase = "<LDAP://" & rootDSE.Get("defaultNamingContext") & ">"
    End If
    strFilter = "(&(objectClass=user)(objectCategory=Person)(samAccountName=" & strUname & "))"
    strAttrib = "name,samAccountName,userPrincipalName,badPwdCount,mail,telephoneNumber,mobile,manager,logonCount,memberOf"
    strScope = "subtree"

    Set objConn = CreateObject("ADODB.Connection")
    objConn.Provider = "ADsDSOObject"
    objConn.Open "Active Directory Provider"

    Set objCmd = CreateObject("ADODB.Command")
    Set objCmd.ActiveConnection = objConn
    objCmd.Properties("Chase referrals") = ADS_CHASE_REFERRALS_ALWAYS
    objCmd.Properties("Page Size") = 500
    objCmd.CommandText = strBase & ";" & strFilter & ";" & strAttrib & ";" & strScope

    Set objRs = objCmd.Execute
    Do Until objRs.EOF
        Sheets("UserInfo").Cells(1, 2).Value = objRs.Fields("name").Value
        Sheets("UserInfo").Cells(2, 2).Value = objRs.Fields("samAccountName").Value
        Sheets("UserInfo").Cells(3, 2).Value = objRs.Fields("userPrincipalName").Value
        Sheets("UserInfo").Cells(4, 2).Value = objRs.Fields("badPwdCount").Value
        Sheets("UserInfo").Cells(5, 2).Value = objRs.Fields("mail").Value
        Sheets("UserInfo").Cells(6, 2).Value = objRs.Fields("telephoneNumber").Value
        Sheets("UserInfo").Cells(7, 2).Value = objRs.Fields("mobile").Value
        Sheets("UserInfo").Cells(8, 2).Value = objRs.Fields("manager").Value
        Sheets("UserInfo").Cells(9, 2).Value = objRs.Fields("logonCount").Value
        objGroups = objRs.Fields("memberOf")
    
    If Not IsNull(objGroups) Then
            intRow = 10
            For Each objGroup In objGroups
                Sheets("UserInfo").Cells(intRow, 2).Value = objGroup
                intRow = intRow + 1
            Next
    End If

    objRs.MoveNext
    Loop
    objRs.Close
    objConn.Close
End Sub

Sub AD1_ADUsers()
    Dim objRootDSE, objConn, objCmd, objRs As Object
    Dim strUname As String
    Dim strBase As String
    Dim strFilter As String
    Dim strAttrib As String
    Dim strScope As String
    Dim intRow As Long
    Const ADS_CHASE_REFERRALS_ALWAYS = &H60

    If targetDC <> "" Then
        strBase = "<LDAP://" & targetDC & ">"
    Else
        Set rootDSE = GetObject("LDAP://RootDSE")
        strBase = "<LDAP://" & rootDSE.Get("defaultNamingContext") & ">"
    End If

    strFilter = "(&(objectClass=user)(objectCategory=Person)(pwdLastSet=*))"
    strAttrib = "samAccountName,userAccountControl,pwdLastSet,userPassword,unixUserPassword,unicodePwd,comment,description,info,mail,employeeId,employeeNumber,telephoneNumber,mobile,manager,servicePrincipalName"
    strScope = "subtree"

    Set objConn = CreateObject("ADODB.Connection")
    objConn.Provider = "ADsDSOObject"
    objConn.Open "Active Directory Provider"

    Set objCmd = CreateObject("ADODB.Command")
    Set objCmd.ActiveConnection = objConn
    objCmd.Properties("Chase referrals") = ADS_CHASE_REFERRALS_ALWAYS
    objCmd.Properties("Page Size") = 500
    objCmd.CommandText = strBase & ";" & strFilter & ";" & strAttrib & ";" & strScope
    

    intRow = 2
    Dim strPassLast As Long
    Dim dateLastSet As Date
    Set objRs = objCmd.Execute
    Do Until objRs.EOF
        Sheets("ADUsers").Cells(intRow, 1).Value = objRs.Fields("samAccountName").Value
        Sheets("ADUsers").Cells(intRow, 2).Value = objRs.Fields("userAccountControl").Value
        Sheets("ADUsers").Cells(intRow, 3).Value = "=MID(BASE(B" & intRow & ",2,32),31,1)"
        Sheets("ADUsers").Cells(intRow, 4).Value = objRs.Fields("pwdLastSet").Value.HighPart
        Sheets("ADUsers").Cells(intRow, 5).Value = objRs.Fields("pwdLastSet").Value.LowPart
        Sheets("ADUsers").Cells(intRow, 6).Value = "=IF(D" & intRow & "=0,""Expired"",((D" & intRow & " * 2^32) + E" & intRow & ")/(8.64*10^11)-109205)"
        Sheets("ADUsers").Cells(intRow, 7).Value = objRs.Fields("userPassword").Value
        Sheets("ADUsers").Cells(intRow, 8).Value = objRs.Fields("unixUserPassword").Value
        Sheets("ADUsers").Cells(intRow, 9).Value = objRs.Fields("unicodePwd").Value
        Sheets("ADUsers").Cells(intRow, 10).Value = objRs.Fields("comment").Value
        Sheets("ADUsers").Cells(intRow, 11).Value = objRs.Fields("description").Value
        Sheets("ADUsers").Cells(intRow, 12).Value = objRs.Fields("info").Value
        Sheets("ADUsers").Cells(intRow, 13).Value = objRs.Fields("mail").Value
        Sheets("ADUsers").Cells(intRow, 15).Value = objRs.Fields("employeeId").Value
        Sheets("ADUsers").Cells(intRow, 16).Value = objRs.Fields("employeeNumber").Value
        Sheets("ADUsers").Cells(intRow, 13).Value = objRs.Fields("telephoneNumber").Value
        Sheets("ADUsers").Cells(intRow, 13).Value = objRs.Fields("mobile").Value
        Sheets("ADUsers").Cells(intRow, 15).Value = objRs.Fields("manager").Value
        Sheets("ADUsers").Cells(intRow, 14).Value = objRs.Fields("servicePrincipalName").Value
        
        Sheets("ADUserUACInfo").Cells(intRow, 1).Value = objRs.Fields("samAccountName").Value
        Sheets("ADUserUACInfo").Cells(intRow, 2).Value = objRs.Fields("userAccountControl").Value
        Sheets("ADUserUACInfo").Cells(intRow, 3).Value = "=MID(BASE(B" & intRow & ",2,32),31,1)"
        Sheets("ADUserUACInfo").Cells(intRow, 4).Value = "=MID(BASE(B" & intRow & ",2,32),27,1)"
        Sheets("ADUserUACInfo").Cells(intRow, 5).Value = "=MID(BASE(B" & intRow & ",2,32),10,1)"
        Sheets("ADUserUACInfo").Cells(intRow, 6).Value = "=MID(BASE(B" & intRow & ",2,32),25,1)"
        Sheets("ADUserUACInfo").Cells(intRow, 7).Value = "=MID(BASE(B" & intRow & ",2,32),16,1)"
        Sheets("ADUserUACInfo").Cells(intRow, 8).Value = "=MID(BASE(B" & intRow & ",2,32),13,1)"
        Sheets("ADUserUACInfo").Cells(intRow, 9).Value = "=MID(BASE(B" & intRow & ",2,32),11,1)"
        Sheets("ADUserUACInfo").Cells(intRow, 10).Value = "=MID(BASE(B" & intRow & ",2,32),9,1)"
        Sheets("ADUserUACInfo").Cells(intRow, 11).Value = "=MID(BASE(B" & intRow & ",2,32),8,1)"
        intRow = intRow + 1
    objRs.MoveNext
    Loop
    objRs.Close
    objConn.Close
    Sheets("ADUsers").Range("F2", "F" & intRow).NumberFormat = "yyyy-mm-dd"
End Sub

Sub AD2_ADInfo()
    Dim objRootDSE, objConn, objCmd, objRs As Object
    Dim strUname As String
    Dim strBase As String
    Dim strFilter As String
    Dim strAttrib As String
    Dim strScope As String
    Dim intRow As Long
    Const ADS_CHASE_REFERRALS_ALWAYS = &H60

    If targetDC <> "" Then
        strBase = "<LDAP://" & targetDC & ">"
    Else
        Set rootDSE = GetObject("LDAP://RootDSE")
        strBase = "<LDAP://" & rootDSE.Get("defaultNamingContext") & ">"
    End If
    strFilter = "(&(objectClass=Domain)(objectCategory=Domain))"
    strAttrib = "name,lockoutDuration,lockoutObservationWindow,lockoutThreshold,minPwdLength,minPwdAge,maxPwdAge,ms-DS-MachineAccountQuota,msDS-Behavior-Version"
    strScope = "subtree"

    Set objConn = CreateObject("ADODB.Connection")
    objConn.Provider = "ADsDSOObject"
    objConn.Open "Active Directory Provider"

    Set objCmd = CreateObject("ADODB.Command")
    Set objCmd.ActiveConnection = objConn
    objCmd.Properties("Chase referrals") = ADS_CHASE_REFERRALS_ALWAYS
    objCmd.Properties("Page Size") = 500
    objCmd.CommandText = strBase & ";" & strFilter & ";" & strAttrib & ";" & strScope

    intRow = 2
    Dim lngDateHigh As Currency
    Dim lngDateLow As Currency
    Dim dateResult As Currency

    Set objRs = objCmd.Execute
    Do Until objRs.EOF
        Sheets("ADInfo").Cells(intRow, 1).Value = objRs.Fields("name").Value
        Sheets("ADInfo").Cells(intRow, 2).Value = objRs.Fields("lockoutDuration").Value.HighPart
        Sheets("ADInfo").Cells(intRow, 3).Value = objRs.Fields("lockoutDuration").Value.LowPart
        Sheets("ADInfo").Cells(intRow, 4).Value = objRs.Fields("lockoutObservationWindow").Value.HighPart
        Sheets("ADInfo").Cells(intRow, 5).Value = objRs.Fields("lockoutObservationWindow").Value.LowPart
        Sheets("ADInfo").Cells(intRow, 6).Value = objRs.Fields("lockoutThreshold").Value
        Sheets("ADInfo").Cells(intRow, 7).Value = objRs.Fields("minPwdLength").Value
        Sheets("ADInfo").Cells(intRow, 8).Value = objRs.Fields("minPwdAge").Value.HighPart
        Sheets("ADInfo").Cells(intRow, 9).Value = objRs.Fields("minPwdAge").Value.LowPart
        Sheets("ADInfo").Cells(intRow, 10).Value = objRs.Fields("maxPwdAge").Value.HighPart
        Sheets("ADInfo").Cells(intRow, 11).Value = objRs.Fields("maxPwdAge").Value.LowPart
        Sheets("ADInfo").Cells(7, 13).Value = objRs.Fields("ms-DS-MachineAccountQuota").Value
        Sheets("ADInfo").Cells(11, 13).Value = objRs.Fields("msDS-Behavior-Version").Value
        intRow = intRow + 1
        objRs.MoveNext
    Loop
    objRs.Close
    
    strFilter2 = "(objectClass=msDS-PasswordSettings)"
    strAttrib2 = "name,msDS-LockoutThreshold,msDs-minimumPasswordLength,msDS-psoAppliesTo"
    objCmd.CommandText = strBase & ";" & strFilter2 & ";" & strAttrib2 & ";" & strScope
    
    intRow = 12
    Set objRs = objCmd.Execute
    Do Until objRs.EOF
        Sheets("ADInfo").Cells(intRow, 1).Value = objRs.Fields("name").Value
        Sheets("ADInfo").Cells(intRow, 2).Value = objRs.Fields("msDS-LockoutThreshold").Value
        Sheets("ADInfo").Cells(intRow, 6).Value = objRs.Fields("msDS-minimumPasswordLength").Value
        Sheets("ADInfo").Cells(intRow, 9).Value = objRs.Fields("msDS-psoAppliesTo").Value.LowPart
        intRow = intRow + 1
        objRs.MoveNext
    Loop
    objRs.Close
    objConn.Close
End Sub

Sub AD3_ADGroups()
    Dim objRootDSE, objConn, objCmd, objRs As Object
    Dim strUname As String
    Dim strBase As String
    Dim strFilter As String
    Dim strAttrib As String
    Dim strScope As String
    Dim intRow As Long
    Const ADS_CHASE_REFERRALS_ALWAYS = &H60

    If targetDC <> "" Then
        strBase = "<LDAP://" & targetDC & ">"
    Else
        Set rootDSE = GetObject("LDAP://RootDSE")
        strBase = "<LDAP://" & rootDSE.Get("defaultNamingContext") & ">"
    End If
    strFilter = "(&(objectClass=Group)(objectCategory=Group))"
    strAttrib = "samAccountName,member,comment,description,ADsPath"
    strScope = "subtree"

    Set objConn = CreateObject("ADODB.Connection")
    objConn.Provider = "ADsDSOObject"
    objConn.Open "Active Directory Provider"

    Set objCmd = CreateObject("ADODB.Command")
    Set objCmd.ActiveConnection = objConn
    objCmd.Properties("Chase referrals") = ADS_CHASE_REFERRALS_ALWAYS
    objCmd.Properties("Page Size") = 500
    objCmd.CommandText = strBase & ";" & strFilter & ";" & strAttrib & ";" & strScope

    intRow = 2
    gpCol = 2

    Set objRs = objCmd.Execute
    Do Until objRs.EOF
        Sheets("ADGroups").Cells(intRow, 1).Value = objRs.Fields("samAccountName").Value
        Sheets("ADGroups").Cells(intRow, 2).Value = objRs.Fields("member").Value
        Sheets("ADGroups").Cells(intRow, 3).Value = objRs.Fields("comment").Value
        Sheets("ADGroups").Cells(intRow, 4).Value = objRs.Fields("description").Value
    
        If objRs.Fields("samAccountName").Value Like "*Admin*" Or objRs.Fields("samAccountName").Value Like "*VPN*" Then
            Set Group = GetObject(objRs.Fields("ADsPath"))
            gpRow = 2
            Sheets("ADGroupInfo").Cells(1, gpCol).Value = objRs.Fields("samAccountName").Value
            For Each member In Group.Members
                Sheets("ADGroupInfo").Cells(gpRow, gpCol) = member.samAccountName
                gpRow = gpRow + 1
            Next
            Sheets("ADGroups").Cells(intRow, 2).Value = strMember
            gpCol = gpCol + 1
        End If
    
        intRow = intRow + 1
        objRs.MoveNext
    Loop
    objRs.Close
    objConn.Close
End Sub

Sub AD4_ADComputers()
    Dim objRootDSE, objConn, objCmd, objRs As Object
    Dim strUname As String
    Dim strBase As String
    Dim strFilter As String
    Dim strAttrib As String
    Dim strScope As String
    Dim intRow As Long
    Const ADS_CHASE_REFERRALS_ALWAYS = &H60

    If targetDC <> "" Then
        strBase = "<LDAP://" & targetDC & ">"
    Else
        Set rootDSE = GetObject("LDAP://RootDSE")
        strBase = "<LDAP://" & rootDSE.Get("defaultNamingContext") & ">"
    End If
    strFilter = "(&(objectClass=Computer)(objectCategory=Computer)(pwdLastSet=*))"
    strAttrib = "samAccountName,userAccountControl,operatingSystem,pwdLastSet,comment,description,logoncount"
    strScope = "subtree"

    Set objConn = CreateObject("ADODB.Connection")
    objConn.Provider = "ADsDSOObject"
    objConn.Open "Active Directory Provider"

    Set objCmd = CreateObject("ADODB.Command")
    Set objCmd.ActiveConnection = objConn
    objCmd.Properties("Chase referrals") = ADS_CHASE_REFERRALS_ALWAYS
    objCmd.Properties("Page Size") = 500
    objCmd.CommandText = strBase & ";" & strFilter & ";" & strAttrib & ";" & strScope

    intRow = 2

    Set objRs = objCmd.Execute
    Do Until objRs.EOF
        Sheets("ADComputers").Cells(intRow, 1).Value = objRs.Fields("samAccountName").Value
        Sheets("ADComputers").Cells(intRow, 2).Value = objRs.Fields("userAccountControl").Value
        Sheets("ADComputers").Cells(intRow, 3).Value = "=MID(BASE(B" & intRow & ",2,32),19,1)"
        Sheets("ADComputers").Cells(intRow, 4).Value = "=MID(BASE(B" & intRow & ",2,32),6,1)"
        Sheets("ADComputers").Cells(intRow, 5).Value = objRs.Fields("operatingSystem").Value
        Sheets("ADComputers").Cells(intRow, 6).Value = objRs.Fields("pwdLastSet").Value.HighPart
        Sheets("ADComputers").Cells(intRow, 7).Value = objRs.Fields("pwdLastSet").Value.LowPart
        Sheets("ADComputers").Cells(intRow, 8).Value = "=((F" & intRow & " *2^32) + G" & intRow & ")/(8.64*10^11) -109205"
        Sheets("ADComputers").Cells(intRow, 9).Value = objRs.Fields("comment").Value
        Sheets("ADComputers").Cells(intRow, 10).Value = objRs.Fields("description").Value
        Sheets("ADComputers").Cells(intRow, 11).Value = "=MID(BASE(B" & intRow & ",2,32),8,1)"
        Sheets("ADComputers").Cells(intRow, 12).Value = "=MID(BASE(B" & intRow & ",2,32),27,1)"
        Sheets("ADComputers").Cells(intRow, 13).Value = objRs.Fields("logoncount").Value

        intRow = intRow + 1
        objRs.MoveNext
    Loop
    objRs.Close
    objConn.Close
    Sheets("ADComputers").Range("H2", "H" & intRow).NumberFormat = "yyyy-mm-dd"
End Sub

Sub AD5_DFSRoot()
    Dim objRootDSE, objConn, objCmd, objRs As Object
    Dim strUname As String
    Dim strBase As String
    Dim strFilter As String
    Dim strAttrib As String
    Dim strScope As String
    Dim intRow As Long
    Const ADS_CHASE_REFERRALS_ALWAYS = &H60

    If targetDC <> "" Then
        strBase = "<LDAP://" & targetDC & ">"
    Else
        Set rootDSE = GetObject("LDAP://RootDSE")
        strBase = "<LDAP://" & rootDSE.Get("defaultNamingContext") & ">"
    End If
    strFilter = "(&(name=*)(msDFS-Propertiesv2=*))"
    strAttrib = "name"
    strScope = "subtree"

    Set objConn = CreateObject("ADODB.Connection")
    objConn.Provider = "ADsDSOObject"
    objConn.Open "Active Directory Provider"

    Set objCmd = CreateObject("ADODB.Command")
    Set objCmd.ActiveConnection = objConn
    objCmd.Properties("Chase referrals") = ADS_CHASE_REFERRALS_ALWAYS
    objCmd.Properties("Page Size") = 500

    objCmd.CommandText = strBase & ";" & strFilter & ";" & strAttrib & ";" & strScope

    intRow = 2

    Set objRs = objCmd.Execute
    Do Until objRs.EOF
        If Not (objRs.Fields("name").Value Like "link*") Then
            Sheets("DFSRoot").Cells(intRow, 1).Value = objRs.Fields("name").Value
            intRow = intRow + 1
        End If
    objRs.MoveNext
    Loop
    objRs.Close
    objConn.Close
End Sub

Sub AD6_TrustedDomains()
    Dim objRootDSE, objConn, objCmd, objRs As Object
    Dim strUname As String
    Dim strBase As String
    Dim strFilter As String
    Dim strAttrib As String
    Dim strScope As String
    Dim intRow As Long
    Const ADS_CHASE_REFERRALS_ALWAYS = &H60

    If targetDC <> "" Then
        strBase = "<LDAP://" & targetDC & ">"
    Else
        Set rootDSE = GetObject("LDAP://RootDSE")
        strBase = "<LDAP://" & rootDSE.Get("defaultNamingContext") & ">"
    End If
    strFilter = "(objectClass=trustedDomain)"
    strAttrib = "name,trustAttributes,trustDirection,trustPartner,trustType"
    strScope = "subtree"

    Set objConn = CreateObject("ADODB.Connection")
    objConn.Provider = "ADsDSOObject"
    objConn.Open "Active Directory Provider"

    Set objCmd = CreateObject("ADODB.Command")
    Set objCmd.ActiveConnection = objConn
    objCmd.Properties("Chase referrals") = ADS_CHASE_REFERRALS_ALWAYS
    objCmd.Properties("Page Size") = 500

    objCmd.CommandText = strBase & ";" & strFilter & ";" & strAttrib & ";" & strScope

    intRow = 2

    Set objRs = objCmd.Execute
    Do Until objRs.EOF
        Sheets("TrustedDomains").Cells(intRow, 1).Value = objRs.Fields("name").Value
        Sheets("TrustedDomains").Cells(intRow, 2).Value = objRs.Fields("trustAttributes").Value
        Sheets("TrustedDomains").Cells(intRow, 3).Value = objRs.Fields("trustDirection").Value
        Sheets("TrustedDomains").Cells(intRow, 4).Value = objRs.Fields("trustPartner").Value
        Sheets("TrustedDomains").Cells(intRow, 5).Value = objRs.Fields("trustType").Value
    objRs.MoveNext
    Loop
    objRs.Close
    objConn.Close
End Sub

Sub Q1_QueryUser()
    Dim objRootDSE, objConn, objCmd, objRs As Object
    Dim strUname As String
    Dim strBase As String
    Dim strFilter As String
    Dim strAttrib As String
    Dim strScope As String
    Dim intRow As Long
    Const ADS_CHASE_REFERRALS_ALWAYS = &H60
    
    strUname = Sheets("QueryUser").Cells(1, 2).Value
    
    If Not IsNull(strUname) Then

	If targetDC <> "" Then
            strBase = "<LDAP://" & targetDC & ">"
	Else
            set rootDSE = GetObject("LDAP://RootDSE")
	    strBase = "<LDAP://" & rootDSE.Get("defaultNamingContext") & ">"
	End If
        strFilter = "(&(objectClass=user)(objectCategory=Person)(samAccountName=" & strUname & "))"
        strAttrib = "name,samAccountName,userPrincipalName,badPwdCount,mail,telephoneNumber,mobile,manager,logonCount,memberOf"
        strScope = "subtree"

        Set objConn = CreateObject("ADODB.Connection")
        objConn.Provider = "ADsDSOObject"
        objConn.Open "Active Directory Provider"

        Set objCmd = CreateObject("ADODB.Command")
        Set objCmd.ActiveConnection = objConn
        objCmd.Properties("Chase referrals") = ADS_CHASE_REFERRALS_ALWAYS
        objCmd.Properties("Page Size") = 500
        objCmd.CommandText = strBase & ";" & strFilter & ";" & strAttrib & ";" & strScope

        Set objRs = objCmd.Execute
        Do Until objRs.EOF
            Sheets("QueryUser").Cells(3, 2).Value = objRs.Fields("name").Value
            Sheets("QueryUser").Cells(4, 2).Value = objRs.Fields("samAccountName").Value
            Sheets("QueryUser").Cells(5, 2).Value = objRs.Fields("userPrincipalName").Value
            Sheets("QueryUser").Cells(6, 2).Value = objRs.Fields("badPwdCount").Value
            Sheets("QueryUser").Cells(7, 2).Value = objRs.Fields("mail").Value
            Sheets("QueryUser").Cells(8, 2).Value = objRs.Fields("telephoneNumber").Value
            Sheets("QueryUser").Cells(9, 2).Value = objRs.Fields("mobile").Value
            Sheets("QueryUser").Cells(10, 2).Value = objRs.Fields("manager").Value
            Sheets("QueryUser").Cells(11, 2).Value = objRs.Fields("logonCount").Value
            objGroups = objRs.Fields("memberOf")
    
            If Not IsNull(objGroups) Then
                intRow = 12
                For Each objGroup In objGroups
                    Sheets("QueryUser").Cells(intRow, 2).Value = objGroup
                    intRow = intRow + 1
                Next
            End If
            
        objRs.MoveNext
        Loop
        objRs.Close
        objConn.Close
    End If
End Sub

Sub Q2_QueryGroup()
    Dim objRootDSE, objConn, objCmd, objRs As Object
    Dim strUname As String
    Dim strGroup As String
    Dim strBase As String
    Dim strFilter As String
    Dim strAttrib As String
    Dim strScope As String
    Dim intRow As Long
    Const ADS_CHASE_REFERRALS_ALWAYS = &H60
    
    strGroup = Sheets("QueryGroup").Cells(1, 2).Value
    
    If Not IsNull(strGroup) Then

	If targetDC <> "" Then
            strBase = "<LDAP://" & targetDC & ">"
	Else
            set rootDSE = GetObject("LDAP://RootDSE")
	    strBase = "<LDAP://" & rootDSE.Get("defaultNamingContext") & ">"
	End If
        strFilter = "(&(objectClass=Group)(objectCategory=Group)(sAMAccountName=" & strGroup & "))"
        strAttrib = "samAccountName,member,comment,description,ADsPath"
        strScope = "subtree"

        Set objConn = CreateObject("ADODB.Connection")
        objConn.Provider = "ADsDSOObject"
        objConn.Open "Active Directory Provider"

        Set objCmd = CreateObject("ADODB.Command")
        Set objCmd.ActiveConnection = objConn
        objCmd.Properties("Chase referrals") = ADS_CHASE_REFERRALS_ALWAYS
        objCmd.Properties("Page Size") = 500
        objCmd.CommandText = strBase & ";" & strFilter & ";" & strAttrib & ";" & strScope

        Set objRs = objCmd.Execute
        Do Until objRs.EOF
            Sheets("QueryGroup").Cells(3, 2).Value = objRs.Fields("samAccountName").Value
            Sheets("QueryGroup").Cells(4, 2).Value = objRs.Fields("comment").Value
            Sheets("QueryGroup").Cells(5, 2).Value = objRs.Fields("description").Value
    
            Set Group = GetObject(objRs.Fields("ADsPath"))
        
            gpRow = 6
            For Each member In Group.Members
                Sheets("QueryGroup").Cells(gpRow, 2) = member.samAccountName
                gpRow = gpRow + 1
            Next
            objRs.MoveNext
        Loop
        objRs.Close
        objConn.Close
    
    End If
End Sub

Sub Q3_QueryLAPS()
    Dim objRootDSE, objConn, objCmd, objRs As Object
    Dim strUname As String
    Dim strBase As String
    Dim strFilter As String
    Dim strAttrib As String
    Dim strScope As String
    Dim intRow As Long
    Const ADS_CHASE_REFERRALS_ALWAYS = &H60
    On Error Resume Next

    If targetDC <> "" Then
        strBase = "<LDAP://" & targetDC & ">"
    Else
        Set rootDSE = GetObject("LDAP://RootDSE")
        strBase = "<LDAP://" & rootDSE.Get("defaultNamingContext") & ">"
    End If
    strFilter = "(&(objectClass=Computer)(objectCategory=Computer)(pwdLastSet=*))"
    strAttrib = "samAccountName,userAccountControl,operatingSystem,pwdLastSet,comment,description,ms-mcs-AdmPwd"
    strScope = "subtree"

    Set objConn = CreateObject("ADODB.Connection")
    objConn.Provider = "ADsDSOObject"
    objConn.Open "Active Directory Provider"

    Set objCmd = CreateObject("ADODB.Command")
    Set objCmd.ActiveConnection = objConn
    objCmd.Properties("Chase referrals") = ADS_CHASE_REFERRALS_ALWAYS
    objCmd.Properties("Page Size") = 500
    objCmd.CommandText = strBase & ";" & strFilter & ";" & strAttrib & ";" & strScope

    intRow = 2

    Set objRs = objCmd.Execute
    Do Until objRs.EOF
        Sheets("QueryLAPS").Cells(intRow, 1).Value = objRs.Fields("samAccountName").Value
        Sheets("QueryLAPS").Cells(intRow, 2).Value = objRs.Fields("userAccountControl").Value
        Sheets("QueryLAPS").Cells(intRow, 3).Value = "=MID(BASE(B" & intRow & ",2,32),19,1)"
        Sheets("QueryLAPS").Cells(intRow, 4).Value = "=MID(BASE(B" & intRow & ",2,32),6,1)"
        Sheets("QueryLAPS").Cells(intRow, 5).Value = objRs.Fields("operatingSystem").Value
        Sheets("QueryLAPS").Cells(intRow, 6).Value = objRs.Fields("pwdLastSet").Value.HighPart
        Sheets("QueryLAPS").Cells(intRow, 7).Value = objRs.Fields("pwdLastSet").Value.LowPart
        Sheets("QueryLAPS").Cells(intRow, 8).Value = "=((F" & intRow & " *2^32) + G" & intRow & ")/(8.64*10^11) -109205"
        Sheets("QueryLAPS").Cells(intRow, 9).Value = objRs.Fields("comment").Value
        Sheets("QueryLAPS").Cells(intRow, 10).Value = objRs.Fields("description").Value
        Sheets("QueryLAPS").Cells(intRow, 11).Value = objRs.Fields("ms-mcs-AdmPwd").Value
        intRow = intRow + 1
        objRs.MoveNext
    Loop
    objRs.Close
    objConn.Close
    Sheets("QueryLAPS").Range("H2", "H" & intRow).NumberFormat = "yyyy-mm-dd"
End Sub

Sub Q4_QuerySQL()
    On Error Resume Next
    Dim objRootDSE, objConn, objCmd, objRs As Object
    Dim strUname As String
    Dim strBase As String
    Dim strFilter As String
    Dim strAttrib As String
    Dim strScope As String
    Dim intRow As Long
    Const ADS_CHASE_REFERRALS_ALWAYS = &H60

    If targetDC <> "" Then
        strBase = "<LDAP://" & targetDC & ">"
    Else
        Set rootDSE = GetObject("LDAP://RootDSE")
        strBase = "<LDAP://" & rootDSE.Get("defaultNamingContext") & ">"
    End If
    strFilter = "(&(objectClass=Computer)(objectCategory=Computer)(servicePrincipalName=MSSQLSvc/*))"
    strAttrib = "samAccountName,userAccountControl,operatingSystem,pwdLastSet,comment,description,servicePrincipalName,ADSPath"
    strScope = "subtree"

    Set objConn = CreateObject("ADODB.Connection")
    objConn.Provider = "ADsDSOObject"
    objConn.Open "Active Directory Provider"

    Set objCmd = CreateObject("ADODB.Command")
    Set objCmd.ActiveConnection = objConn
    objCmd.Properties("Chase referrals") = ADS_CHASE_REFERRALS_ALWAYS
    objCmd.Properties("Page Size") = 500
    objCmd.CommandText = strBase & ";" & strFilter & ";" & strAttrib & ";" & strScope

    intRow = 2

    Set objRs = objCmd.Execute
    Do Until objRs.EOF
        
        Set Computer = GetObject(objRs.Fields("ADsPath"))
        SPNValue = ""
        For Each SPN In Computer.servicePrincipalName
            If InStr(SPN, "MSSQLSvc") > 0 Then
                Sheets("QuerySQL").Cells(intRow, 1).Value = objRs.Fields("samAccountName").Value
                Sheets("QuerySQL").Cells(intRow, 2).Value = objRs.Fields("userAccountControl").Value
                Sheets("QuerySQL").Cells(intRow, 3).Value = "=MID(BASE(B" & intRow & ",2,32),19,1)"
                Sheets("QuerySQL").Cells(intRow, 4).Value = "=MID(BASE(B" & intRow & ",2,32),6,1)"
                Sheets("QuerySQL").Cells(intRow, 5).Value = objRs.Fields("operatingSystem").Value
                Sheets("QuerySQL").Cells(intRow, 6).Value = objRs.Fields("pwdLastSet").Value.HighPart
                Sheets("QuerySQL").Cells(intRow, 7).Value = objRs.Fields("pwdLastSet").Value.LowPart
                Sheets("QuerySQL").Cells(intRow, 8).Value = "=((F" & intRow & " *2^32) + G" & intRow & ")/(8.64*10^11) -109205"
                Sheets("QuerySQL").Cells(intRow, 9).Value = objRs.Fields("comment").Value
                Sheets("QuerySQL").Cells(intRow, 10).Value = objRs.Fields("description").Value

                SPNComponents = Split(SPN, "/")
                DataSource = SPNComponents(1)

                Dim conn As New ADODB.Connection
                Dim recs As New ADODB.Recordset
                Dim ConnectionString As String
                Dim StrQuery As String

                ConnectionString = "Provider=SQLOLEDB;Data Source=" & DataSource & ";Initial Catalog=Master;Integrated Security=SSPI;"

                conn.Open ConnectionString
                conn.CommandTimeout = 900
                
                If conn.State = 1 Then
                    Sheets("QuerySQL").Cells(intRow, 11).Value = "Success"
                    
                    StrQuery = "SELECT name FROM master.dbo.sysdatabases"
                    recs.Open StrQuery, conn
                    txtDatabases = ""
                    For i = 1 To recs.RecordCount
                        If i = 1 Then
                            txtTables = rs.Fields("name").Value
                        Else
                            txtTables = txtTables & ";" & rs.Fields("name").Value
                        End If
                    Next
                    Sheets("QuerySQL").Cells(intRow, 12).Value = txtDatabases
                Else
                    Sheets("QuerySQL").Cells(intRow, 11).Value = "Failed"
                    Sheets("QuerySQL").Cells(intRow, 12).Value = "None"
                End If
                conn.Close
            End If
        Next
        
        intRow = intRow + 1
        objRs.MoveNext
    Loop
    objRs.Close
    objConn.Close
End Sub

Sub Q5_QueryGPO()
    Dim oFSO As Object
    Dim oFolder As Object
    Dim XDoc As Object
    Dim root As Object
    Dim oFile As Object, sf
    Dim n As Object
    Dim i As Integer, colFolders As New Collection, ws As Worksheet
    Dim fileExt
    Dim strDnsDomain As String

    strDnsDomain = Environ$("userdnsdomain")
    
    Set ws = Sheets("QueryGPO")
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.getFolder("\\" & strDnsDomain & "\sysvol\" & strDnsDomain & "\Policies")
    
    colFolders.Add oFolder
    
    i = 0
    
    Do While colFolders.Count > 0
        On Error Resume Next
        Set oFolder = colFolders(1)
        colFolders.Remove (1)
        
        If oFolder.Name <> "PolicyDefinitions" Then
                For Each oFile In oFolder.Files
                    fileExt = LCase(Right(oFile.Name, 4))
                    If fileExt = ".bat" Or fileExt = ".ps1" Then
                        ws.Cells(i + 1, 1) = "Script"
                        ws.Cells(i + 1, 2) = oFolder.Path
                        ws.Cells(i + 1, 3) = oFile.Name
                        i = i + 1
                    End If
                    If fileExt = ".xml" Then
                        Select Case LCase(oFile.Name)
                            Case "drives.xml"
                                ws.Cells(i + 1, 1) = "Mapped Drive Policy"
                                ws.Cells(i + 1, 2) = oFolder.Path
                                ws.Cells(i + 1, 3) = oFile.Name
                                i = i + 1
                                
                                Set XDoc = CreateObject("MSXML2.DOMDocument")
                                XDoc.async = False: XDoc.validateOnParse = False
                                XDoc.Load (oFolder.Path & "\" & oFile.Name)
                                Set root = XDoc.DocumentElement
                                Set xmlNodes = XDoc.SelectNodes("/Drives/Drive/Properties")
                                For Each n In xmlNodes
                                    ws.Cells(i + 1, 2) = n.Attributes(4).Text
                                    ws.Cells(i + 1, 3) = n.Attributes(8).Text
                                    i = i + 1
                                Next n
                                
                                Set XDoc = Nothing
                            Case "shortcuts.xml"
                                ws.Cells(i + 1, 1) = "Shortcut Policy"
                                ws.Cells(i + 1, 2) = oFolder.Path
                                ws.Cells(i + 1, 3) = oFile.Name
                                i = i + 1
                                
                                Set XDoc = CreateObject("MSXML2.DOMDocument")
                                XDoc.async = False: XDoc.validateOnParse = False
                                XDoc.Load (oFolder.Path & "\" & oFile.Name)
                                Set root = XDoc.DocumentElement
                                Set xmlNodes = XDoc.SelectNodes("/Shortcut/Shortcuts/Properties")
                                For Each n In xmlNodes
                                    ws.Cells(i + 1, 2) = n.Attributes(8).Text
                                    ws.Cells(i + 1, 3) = n.Attributes(1).Text
                                    i = i + 1
                                Next n
                                
                                Set XDoc = Nothing
                                
                            Case "groups.xml"
                            Case "printers.xml"
                            Case "files.xml"
                            Case "folders.xml"
                            Case "internetsettings.xml"
                            Case "poweroptions.xml"
                            Case "registry.xml"
                            Case "scheduledtasks.xml"
                            Case "services.xml"
                            Case "startmenu.xml"
                            Case "startmenutaskbar.xml"
                            Case Else
                                ws.Cells(i + 1, 1) = "Unknown Policy"
                                ws.Cells(i + 1, 2) = oFolder.Path
                                ws.Cells(i + 1, 3) = oFile.Name
                                i = i + 1
                        End Select
                    End If
                Next oFile
                    
                For Each sf In oFolder.subfolders
                    colFolders.Add sf
                Next sf
        End If
    Loop
End Sub

Sub AD7_ADSites()
    Dim objRootDSE, objConn, objCmd, objRs As Object
    Dim strUname As String
    Dim strBase As String
    Dim strFilter As String
    Dim strAttrib As String
    Dim strScope As String
    Dim intRow As Long
    Const ADS_CHASE_REFERRALS_ALWAYS = &H60

    If targetDC <> "" Then
        strBase = "<LDAP://" & targetDC & ">"
    Else
        set rootDSE = GetObject("LDAP://RootDSE")
	strBase = "<LDAP://" & rootDSE.Get("defaultNamingContext") & ">"
    End If
    strFilter = "(objectClass=Site)"
    strAttrib = "name,siteObjectBL"
    strScope = "subtree"

    Set objConn = CreateObject("ADODB.Connection")
    objConn.Provider = "ADsDSOObject"
    objConn.Open "Active Directory Provider"

    Set objCmd = CreateObject("ADODB.Command")
    Set objCmd.ActiveConnection = objConn
    objCmd.Properties("Chase referrals") = ADS_CHASE_REFERRALS_ALWAYS
    objCmd.Properties("Page Size") = 500
    objCmd.CommandText = strBase & ";" & strFilter & ";" & strAttrib & ";" & strScope

    intRow = 2

    Set objRs = objCmd.Execute
    Do Until objRs.EOF
        Sheets("ADSites").Cells(intRow, 1).Value = objRs.Fields("name").Value
        Sheets("ADSites").Cells(intRow, 2).Value = objRs.Fields("siteObjectBL").Value
        intRow = intRow + 1
        objRs.MoveNext
    Loop
    objRs.Close
    objConn.Close
End Sub

Sub AD8_ADSubnets()
    Dim objRootDSE, objConn, objCmd, objRs As Object
    Dim strUname As String
    Dim strBase As String
    Dim strFilter As String
    Dim strAttrib As String
    Dim strScope As String
    Dim intRow As Long
    Const ADS_CHASE_REFERRALS_ALWAYS = &H60

    If targetDC <> "" Then
        strBase = "<LDAP://" & targetDC & ">"
    Else
        set rootDSE = GetObject("LDAP://RootDSE")
	strBase = "<LDAP://" & rootDSE.Get("defaultNamingContext") & ">"
    End If
    strFilter = "(objectClass=Subnet)"
    strAttrib = "name,siteObject"
    strScope = "subtree"

    Set objConn = CreateObject("ADODB.Connection")
    objConn.Provider = "ADsDSOObject"
    objConn.Open "Active Directory Provider"

    Set objCmd = CreateObject("ADODB.Command")
    Set objCmd.ActiveConnection = objConn
    objCmd.Properties("Chase referrals") = ADS_CHASE_REFERRALS_ALWAYS
    objCmd.Properties("Page Size") = 500
    objCmd.CommandText = strBase & ";" & strFilter & ";" & strAttrib & ";" & strScope

    intRow = 2

    Set objRs = objCmd.Execute
    Do Until objRs.EOF
        Sheets("ADSubnets").Cells(intRow, 1).Value = objRs.Fields("name").Value
        Sheets("ADSubnets").Cells(intRow, 2).Value = objRs.Fields("siteObject").Value
        intRow = intRow + 1
        objRs.MoveNext
    Loop
    objRs.Close
    objConn.Close
End Sub

Sub F1_ApplyConditionalFormatting()
    Dim ws As Worksheet
    Dim oDisabledUsersFormat As FormatCondition
    Dim conditionStringArray
    
    For Each ws In Worksheets
        If ws.Name = "ADUsers" Then
            ws.Activate
    
                On Error Resume Next
                    ws.Cells.FormatConditions.Delete
                On Error GoTo 0
    
                With Sheets("ADUsers").Range("A:N")
                    .FormatConditions.Add Type:=xlExpression, Formula1:="=$C1=""1"""
                    .FormatConditions(1).Interior.Color = RGB(255, 80, 80)
                End With
                
                conditionStringArray = Array("pass", "pw", "p:", "!")
                For myLoop = LBound(conditionStringArray) To UBound(conditionStringArray)
                    Set oFormatCondition = Sheets("ADUsers").Range("J:J").FormatConditions.Add(Type:=Excel.XlFormatConditionType.xlTextString, TextOperator:=Excel.XlContainsOperator.xlContains, String:=conditionStringArray(myLoop))
                    With oFormatCondition
                        .Interior.Color = RGB(255, 255, 153)
                    End With
                    Set oFormatCondition = Sheets("ADUsers").Range("K:K").FormatConditions.Add(Type:=Excel.XlFormatConditionType.xlTextString, TextOperator:=Excel.XlContainsOperator.xlContains, String:=conditionStringArray(myLoop))
                    With oFormatCondition
                        .Interior.Color = RGB(255, 255, 153)
                    End With
                    Set oFormatCondition = Sheets("ADUsers").Range("L:L").FormatConditions.Add(Type:=Excel.XlFormatConditionType.xlTextString, TextOperator:=Excel.XlContainsOperator.xlContains, String:=conditionStringArray(myLoop))
                    With oFormatCondition
                        .Interior.Color = RGB(255, 255, 153)
                    End With
                Next
        End If
        If ws.Name = "ADUserUACInfo" Then
            ws.Activate
    
                On Error Resume Next
                    ws.Cells.FormatConditions.Delete
                On Error GoTo 0
                
                ' Highlight rows where accounts are disabled
                With Sheets("ADUserUACInfo").Range("A:K")
                    .FormatConditions.Add Type:=xlExpression, Formula1:="=$C1=""1"""
                    .FormatConditions(1).Interior.Color = RGB(255, 80, 80)
                End With
                
                ' Highlight accounts with password not required
                Set oFormatCondition = Sheets("ADUserUACInfo").Range("D:D").FormatConditions.Add(Type:=Excel.XlFormatConditionType.xlTextString, TextOperator:=Excel.XlContainsOperator.xlContains, String:="1")
    
                With oFormatCondition
                    .Interior.Color = RGB(255, 255, 153)
                    .Font.Bold = True
                End With
                
                ' Highlight accounts with preauth not required
                Set oFormatCondition = Sheets("ADUserUACInfo").Range("E:E").FormatConditions.Add(Type:=Excel.XlFormatConditionType.xlTextString, TextOperator:=Excel.XlContainsOperator.xlContains, String:="1")
    
                With oFormatCondition
                    .Interior.Color = RGB(255, 255, 153)
                    .Font.Bold = True
                End With
                
                ' Highlight accounts with reversible encryption
                Set oFormatCondition = Sheets("ADUserUACInfo").Range("F:F").FormatConditions.Add(Type:=Excel.XlFormatConditionType.xlTextString, TextOperator:=Excel.XlContainsOperator.xlContains, String:="1")
    
                With oFormatCondition
                    .Interior.Color = RGB(255, 255, 153)
                    .Font.Bold = True
                End With
                
                
        End If
        If ws.Name = "ADComputers" Then
            ws.Activate
    
                On Error Resume Next
                    Sheets("ADComputers").Range("A:A").FormatConditions.Delete
                On Error GoTo 0
                
                conditionStringArray = Array("BITBUCKET", "CONFLUENCE", "DB", "DC", "DMZ", "ESX", "EXCH", "GIT", "JIRA", "JMP", "JUMP", "MAIL", "PASSWORD", "PKI", "SCCM", "SECRET", "SQL", "SVN", "TANK", "TFS", "THY", "TS", "VC", "VMW", "VSP", "WDS")
                For myLoop = LBound(conditionStringArray) To UBound(conditionStringArray)
                    Sheets("ADComputers").Range("A:A").Select
                    Set oFormatCondition = Sheets("ADComputers").Range("A:A").FormatConditions.Add(Type:=Excel.XlFormatConditionType.xlTextString, TextOperator:=Excel.XlContainsOperator.xlContains, String:=conditionStringArray(myLoop))
                    With oFormatCondition
                        .Interior.Color = RGB(102, 255, 255)
                    End With
                Next
                
                On Error Resume Next
                    Set oFormatCondition = Sheets("ADComputers").Range("C:C").FormatConditions(1)
                    oFormatCondition.Delete
                On Error GoTo 0
    
                Sheets("ADComputers").Range("C:C").Select
                Set oFormatCondition = Sheets("ADComputers").Range("C:C").FormatConditions.Add(Type:=Excel.XlFormatConditionType.xlTextString, TextOperator:=Excel.XlContainsOperator.xlContains, String:="1")
    
                With oFormatCondition
                    .Interior.Color = RGB(255, 255, 153)
                End With
                
                On Error Resume Next
                    Set oFormatCondition = Sheets("ADComputers").Range("D:D").FormatConditions(1)
                    oFormatCondition.Delete
                On Error GoTo 0
    
                Sheets("ADComputers").Range("D:D").Select
                Set oFormatCondition = Sheets("ADComputers").Range("D:D").FormatConditions.Add(Type:=Excel.XlFormatConditionType.xlTextString, TextOperator:=Excel.XlContainsOperator.xlContains, String:="1")
    
                With oFormatCondition
                    .Interior.Color = RGB(255, 255, 153)
                End With
        End If
    Next ws
    
End Sub

