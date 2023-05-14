Attribute VB_Name = "Module2"
Public Const l8 = 0.155125
Public Const l5 = 0.98485
Public Const l2 = 0.049475
Public Const x238 = 0.9928
Public Const k = 137.88
Public Const x235 = 0.0072
Dim pass As Boolean
Dim presedence As Boolean
Public stop_calculation As Boolean
Dim enable207 As Boolean
Dim enable208 As Boolean
Private sx As Double
Private sxt As Double
Private sy As Double
Private syt As Double
Private sxr As Double
Private syr As Double
Dim t2 As Double

Dim ctest As Boolean
Public ndisc As Integer

'***********************************************************************************************
'ComPbCorr#3.0
'***********************************************************************************************
'Version info and development history:
' Version 2.01: Instability problem fixed, comments added. Sent to GEMOC 18.06.02
'--------------------
'Version 2.02 has an improved concordance test which seems to agree with Isoplot.
'A bug in the t1 newton routine has been fixed, but does not seem to have any effect.
'--------------------
'Version 2.03 (Sept. 10, 2002)  Bugfix:
'
'Bug in printout routine (207/206 ratio in uncorrected analyses) found and fixed
'Correct analytical expressions for errors in the Pb/U and Pb/Th ages introduced. Results now matches Glitter,
'Error in the formula for correlation of errors in corrected data found and
'fixed (see Eq. 14 in Chem. Geol paper).
'Code has been tidied up somewhat -  unused lines and lines originally used for testing have been removed.
' One unnecessary column removed from output
'--------------------------
'Version 2.1 (Sept. 17, 2002)  Major modification (208-correction) + bugfix
'
'Introduced optional 208-correction, to be used for analyses where everything else fails
'
'Corrected misprint in t_netwon rutine ("xt" instead of "x" in first test)
'Relaxed conversion test in same routine from 1E-14 to 1E-13 to avoid crashes for old zircons
'Removed t1 and st1 from output of 207- or 208 corrected points
'Found and corrected error in initial discordance test
'Found inconsistency in second discordance test - fixed the problem
'----------------------------
'Version 3.01 (Oct. 3 ,2002)
'Modified input to read U/Th ratios and errors from worksheet.
'Removed raw counts from input
'Shifted output columns accordingly
'Introduced three-way options for error correlations, allowing observed rho for individual samples
'-----------------------------
'Version 3.12, Nov. 21st
'Yet another exception in the concordance test detected
'Second discordance test modified to check both against calculated discordance and against isotope ratios
'-----------------------------
'02.Jan. 03
'Still a bug in discordance test - versions up to 3.12 report negative discordances for quite
'heavily inversely discordant points. Seems to be a problem only for quite strong corrections, and the
'points ought to be kicked out anyway. Cause: Faulty expression for discordance.
'Version 3.13, Jan. 2nd, 2003 - this bug fixed.

'------------------------------------
'March 2003: Report from Javier Suárez that grains which should be OK revert to 207-Pb correction. The reason
'was found to be that grains plotting ever so slightly off the discordia from t2 to the wrong side may recalculate with
'negative common Pb. Which then invokes a 207 correction. In this
'case, t1> raw 207/206 age, and no correction should be applied. This combination  causes an exception in the "crazyness" test
'from October 2002, which has been modified by an extra statement which resets pass to true, and thereby preventing an
'unnecessary 207 correction.
'
'Version 3.14 March 27th 2003
'
'========================================

' Version 3.15 -
' Work started: May 27th, 2003, Sydney
' Changed default for discordance test to 2
' Changed default fractionation factor for GEMOC to something between observed 213 and 266 fractionation
' Revised concordance test implemented: Now the rim of the error ellipse is calcualted and checked against t
' the concordia at 40 points between extremes.
' 207 correction made optional....
' Restructure error calculation....
' Calculate U and Th ppm for GEMOC users  ?
' Biased 207 age ?....
' Tested on Javier's data: .......
' Completed for prerelease:
' Released to users:
'********************************************

Sub PbCorr2()
    '***********************************************************************************************
   'Start of interactive part. Modified from the first version
   '*******************************************************************************************
'**************
't2 and commonlead age default at zero
'************************
tc = 0
t2 = 0
presedence = False
''''''' testing
sxr = 0

ndisc = 2
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Set and display first form
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


UserForm1.CheckBox4.Value = presedence
UserForm1.TextBox2.Value = tc
UserForm1.TextBox3.Value = t2
UserForm1.CheckBox3.Value = False
UserForm1.TextBox7.Value = 18.7
UserForm1.TextBox8.Value = 15.628
UserForm1.TextBox9.Value = 38.63
UserForm1.TextBox10.Value = 0.9
UserForm1.OptionButton1.Value = False
UserForm1.OptionButton2.Value = False
UserForm1.OptionButton3.Value = True
UserForm1.TextBox11.Value = ndisc
UserForm1.CheckBox5.Value = False

'315 - additional control to turn 207 correction on / off
UserForm1.CheckBox6.Value = False

' Alternatives required by GEMOC input, plus stop crash on default options if "RHOT" is empty

If gemoc_on = True Or [L12] = "" Then
UserForm1.OptionButton3.Visible = False
UserForm1.OptionButton2.Value = True
Else
UserForm1.OptionButton3.Visible = True
UserForm1.OptionButton3.Value = True
End If


UserForm1.Show

tc = UserForm1.TextBox2.Value

t2 = UserForm1.TextBox3.Value
presedence = UserForm1.CheckBox4.Value

ndisc = UserForm1.TextBox11.Value
[af3] = ndisc

enable208 = UserForm1.CheckBox5.Value
' new in 315
enable207 = UserForm1.CheckBox6.Value
'
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'Set and display second form
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
gemoc_on = False


UserForm2.TextBox1.Text = "Analyses will be corrected to a 3D discordia line to the specified lower intercept."
    If enable208 = True And enable207 = True Then
      UserForm2.TextBox3.Text = "Pb-207 correction will be used for grains which correct to inversely discordant compositions by the discordance pattern method, grains which still correct to negative 208Pb/232Th ratios will be corrected by the Pb-208 method"
ElseIf enable207 = True Then
        UserForm2.TextBox3.Text = "Pb-207 correction will be used for grains which  correct to inversely discordant compositions by the discordance pattern method"
    ElseIf enable208 = True And enable207 = False Then
     UserForm2.TextBox3.Text = "Pb-208 correction will be used for grains which  correct to inversely discordant compositions by the discordance pattern method"
   Else: UserForm2.TextBox3.Text = ""
    End If
    
    If enable208 = True Or enable207 = True Then
    UserForm2.TextBox4.Text = "Pb-207 or Pb-208 corrections are unlikely to yield interpretable results"
    Else
     UserForm2.TextBox4.Text = ""
     End If
UserForm2.TextBox2.Value = t2
UserForm2.Show

'Get ready for data input

    If UserForm1.CheckBox3.Value = True Then
    common64 = UserForm1.TextBox7.Value
    common74 = UserForm1.TextBox8.Value
    common84 = UserForm1.TextBox9.Value
    common6 = common64 / (1 + common64 + common74 + common84)
    common7 = common74 / (1 + common64 + common74 + common84)
    common8 = common84 / (1 + common64 + common74 + common84)

    
    
    
    Else
    'calculate composition of common-lead at preferred age
    common6 = commont(tc, 6)
    common7 = commont(tc, 7)
    common8 = commont(tc, 8)
    c7 = common7 / common6
    c8 = common8 / common6
    [AP1] = Format(c7, "#.####")
    [AR1] = Format(c8, "#.####")
'
'calculate apparent age of common lead



    End If
  
    [BC1] = Format(t2 * 1000)
    
    '*************************************************************************************************
    'Ready to start main loop
   '************************************************************************************************
    Application.ScreenUpdating = False
    Worksheets("Random").Visible = False
    Worksheets("data").Activate
    ActiveSheet.Select
    
      '  If UserForm1.OptionButton2.Value = True Then Cells(4, 12) = "Assumed"
      '  If UserForm1.OptionButton1.Value = True Then Cells(4, 12) = "Inferred"
      '  If UserForm1.OptionButton3.Value = True Then Cells(4, 12) = "Observed"
       
        If UserForm1.OptionButton2.Value = True Then
            Cells(4, 15) = "Assumed"
            ElseIf UserForm1.OptionButton3 = True Then Cells(4, 13) = "Observed"
            Else
            Cells(4, 15) = "Inferred"
        End If
        
    j = 7


    Do Until Cells(j, 1) = 0
    'set initial values to prevent crash - these will be overridden
    t1 = 0
    st1 = 0
    fc = 0
    sfc = 0
    rhot = 0.9
    pass = True
    'Comment = "OK1"
    correction_type = "None"
    Range(Cells(j, 15), Cells(j, 56)).Clear
   
    '---------------
    'Read data from worksheet
    '---------------
    xt = Cells(j, 4)
    yt = Cells(j, 6)
    zt = Cells(j, 8)
    xyt = Cells(j, 2)
    'U = Cells(j, 14)
    'Th = Cells(j, 13)
  
    'ut = U / Th
    ut = Cells(j, 10)
    'Read errors
    '
    sxytobs = Cells(j, 3)
    '
    syt = Cells(j, 7)
    sxt = Cells(j, 5)
    szt = Cells(j, 9)
    sxyt = Cells(j, 3)
    sc7 = 0.01 * c7
    sc8 = 0.01 * c8
    '
    sut = Cells(j, 11)
    If Cells(j, 11) = "" Then sut = 0.01 * ut
    If sut < 0.01 Then sut = 0.01
   
    'sut = 0.01 * ut
    
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    'Esimate error correlation from  data given i worksheet or set it to a specified value
    '
    If UserForm1.OptionButton2.Value = True Then
        rhot = UserForm1.TextBox10.Text
        ElseIf UserForm1.OptionButton3 = True Then rhot = Cells(j, 12)
        Else
        rhot = -1 / 2 * (-sxt ^ 2 * yt ^ 2 * xyt ^ 2 - syt ^ 2 * xt ^ 2 * xyt ^ 2 + sxyt ^ 2 * xt ^ 2 * yt ^ 2) / xt / yt / xyt ^ 2 / sxt / syt
    End If
    
    
    sxyt = xyt * ((sxt / xt) ^ 2 + (syt / yt) ^ 2 - 2 * rhot * (sxt / xt) * (syt / yt)) ^ 0.5
    '---------------
    'Readin finished
    '---------------
        '============================
        'calculate  raw ages
        '============================
        'first 76 age of uncorrected data
        age76 = leadage(xyt, t2)
        
        'Concordant lead at raw 76 age, for initial discordance only
 '315: Next to can be removed ?
            x0 = Exp(l5 * age76) - 1
            y0 = Exp(l8 * age76) - 1
            'rough initial discordance, disregarding thorogenic lead
  ' 315: Remove old disc calculation
           ' initial_disc = (Sqr(xt ^ 2 + yt ^ 2)) / Sqr(x0 ^ 2 + y0 ^ 2) - 1
          
            initial_disc = discord(xt, yt, age76)

 '315: Insert minimum rim discordance calculation here
  initrimdisc = rimdisc(xt, sxt, yt, syt, rhot, ndisc)
   

   
   
            'first discordance test
 'Cells(j, 17) = intitial_disc * 1000
'Cells(j, 18) = pass
  'CORRECTED 170902
  ' 315 modified
  
  '315: Obsolete concordance test replaced by:
          'If Abs(initrimdisc) < 0.01 Then
          If initrimdisc = 0 Then
                pass = False
                Comment = "Concordant"
                correction_type = "None"
                ElseIf initrimdisc > 0.01 Then
                pass = False
                Comment = "Initially inv. disc."
                        
                      If t2 > 0 Then Comment = "t2 too high"
                        
                Else: End If
                
  ' Error expressions corrected , version 2.03
              
                
                 erra76 = leadageerr(xyt, sxyt, t2)
                 age7 = Log(xt + 1) / l5
                 erra7 = sxt / (xt + 1) / l5
                 age6 = Log(yt + 1) / l8
                  erra6 = syt / (yt + 1) / l8
                 age8 = Log(zt + 1) / l2
                erra8 = szt / (zt + 1) / l2
                
                

                  

    


'End If


'315: remove old concordance test
'Cells(j, 18) = ctest
                   ' If ctest = True Then
                    'pass = False
                    'Comment = "Concordant"
                    'initial_disc = 0
                    'End If
                    
'Cells(j, 18) = pass
 
                '============================
                'raw ages done
                '============================
        
'''''''''
' Start main if loop ????
''''''''''''''
                '^^^^^^^^^^^^^^^^^^^^
                'Start parameter calculation
                '^^^^^^^^^^^^^^^^^^^^
                ' first check if correction makes sense, then
                ' call routine for numeric calculation of t1
                
                
                
          
                

If pass = True Then

' Calculate t1 and corrected parameters based on discordance model

                    t1 = t_newton(xt, yt, zt, ut, t2, c7, c8)

                    x1 = Exp(l5 * t1) - 1
                    y1 = Exp(l8 * t1) - 1
                    z1 = Exp(l2 * t1) - 1
                    x2 = Exp(l5 * t2) - 1
                    y2 = Exp(l8 * t2) - 1
                    z2 = Exp(l2 * t2) - 1

                    fc = commonlead(xt, yt, zt, t1, t2, c7)
                       
   
                    yr = yt * (1 - fc)
                    xr = xt - k * c7 * fc * yt
                    zr = zt - ut * c8 * fc * yt
                  
                    xyr = xr / yr / k
    
    'new lines 02 Jan 03
    'calculation of 207/206 age of corrected data, for testing purposes
   
                txyr = leadage(xyr, t2)
                
                  '  xconc = Exp(l5 * txyr) - 1
                   ' yconc = Exp(l8 * txyr) - 1
                    
' Do the error propagation using Monte Carlo routine for errors in corrected ratios and error correlation
 
Call errors(xt, yt, zt, sxt, syt, szt, ut, sut, c7, c8, sc7, sc8, xr, yr, zr, xyr, rhot, "Disc", t1, t2, fc, st1, sfc, syr, sxr, szr, rhor, sxyr)

' Second disconcordance test - checks if the correction has made sense, or gone beyond the concordia

 
'discordance = discord(xr, yr, txyr)
discordance = discord(xr, yr, t1)

 rimdiscordance = rimdisc(xr, sxr, yr, syr, rhor, ndisc)
 
 

 
 
  ' as default, reject discordance correction if it goes inversely discordant
  
            If rimdiscordance > 0# Then
                    pass = False
                                     
                   Comment = "Disc. corr: Inversely discordant"
            
            Else
                    pass = True
            
                    correction_type = "Disc"
                    Comment = "OK"
            
            End If
    '
    '
    '
End If


'If pass = False and comment <> "Concordant" Then
'but keep it as it is, if the user has asked for it
                   
            If presedence = True Then
                   pass = True
                   Comment = "Discordance test overridden"
                        
            End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

 ' Here follows diverse tests for oddballs and exceptions.
 ' The  routines are inherited from #3.14 and earlier, and most are ad-hoc solutions to users' problems.
 ' Relevance of thes tests with current (3.15) algoritm is unknown (as of May 29th, 2003)
                   
                   If pass = True And zr < z2 Then
                   pass = False
                   Comment = "208Pb too low for disc. corr."
                   End If
                 'DISCORDANCE CORRECTION DOES NOT WORK WITH VERY OLD ZIRCONS WHICH HAVE HAD THEIR U/TH RATIIO MESSED UP
                 'THESE WOULD CORRECT TO NEGATIVE AGES, SO MAKE A TEST TO PREVENT THE PROGRAM FROM CRASHING. DO NOT NECESSARILY
                 'TRUST THE PB-207 CORRECTED AGE DERIVED FROM SUCH POINTS
                 '
 
                  If pass = True And t1 <= 0 And Comment <> "Concordant" Then
                  pass = False
                  Comment = "Disc. corr. failed to give t1>0"
                  End If
   ' ----------------------------
   '"Crazyness" test
   '-----------------------------
    ' test if correction routine may have gone crazy
    ' modified 11.10.02 to account for Ayesha's exception with negative output ages
                  
    ' Comment May 2003: This test to be regarded with due sceptisism
                  'txyr = leadage(xyr, t2)
                  
      


                ' If txyr < 0 Or txyr > t1 * 1.01 Then
                ' pass = False
               ' Comment = "Disc. corr. failed"
                 '
                 '
    'The following line added 27th  March 2003 to prevent a 207-correction in cases where fc <0
    
                ' If t1 > age76 * 0.99 Then pass = True
                
                 ' End If

     ' End of crazyness test ------------------------
                  
'End of oddball testing
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                
 '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
 ' Here follows routines for Pb-207 and Pb-208 corrections, to be implemented only if the discordance
 ' correction does not work, AND THESE TWO METHODS HAVE BEEN AUTHORIZED BY THE USER IN DIALOG 1
 ' When implemented, a warning will be issued.
     If pass = False And Comment <> "Concordant" Then
      
            If pass = False And enable207 = True And Comment = "Disc. corr: Inversely discordant" Then
     
                        t1 = t207_corr(xt, yt, c7)
                        yr = Exp(l8 * t1) - 1
                
                        xr = Exp(l5 * t1) - 1
                        fc = (xt - xr) / k / c7 / yt
        
                        zr = zt - ut * c8 * fc * yt
                           
                        discordance = 0
                        xyr = xr / yr / k
                        correction_type = "Pb-207"

                        pass = True
                            
                        'modified 11.10.02
                        txyr = leadage(xyr, t2)
                                    
                        Else
   
                            
                            'correction_type = "Disc"
            End If
                        
    ' If the 207 correction also fails, and 208/232 goes negative, invoke a 208 correction (if allowed by user)
            
            
            If ((zr < 0 And enable208 = True) _
            Or (pass = False And enable208 = True And Comment = "Disc. corr: Inversely discordant" And enable207 = False)) Then
                    t1 = t208_corr(yt, zt, ut, c8)
                                    yr = Exp(l8 * t1) - 1
                
                                    zr = Exp(l2 * t1) - 1
                                    fc = 1 - yr / yt
        
                                  xr = xt - yt * c7 * k * fc
                           
                                     
                                    discordance = Sqr((yr - y2) ^ 2 + (xr - x2) ^ 2) / Sqr((y1 - y2) ^ 2 + (x1 - x2) ^ 2) - 1
                   
                                    xyr = xr / yr / k
                                    correction_type = "Pb-208"
'Cells(j, 27) = t1
                                    pass = True
                                    'modified 11.10.02
                             
   
                             txyr = leadage(xyr, t2)
                                    
                End If
   
    'calculate  discordant 7/6  age w/errors,
    '
                        'modified 11.10.02
                        'txyr = leadage(xyr, t2)
    
      
    '315: Error propagation in separate routine
    'If correction_type = "Pb-207" Or correction_type = "Pb-208" Then
    
   ' Call errors(xt, yt, zt, sxt, syt, szt, ut, sut, c7, c8, sc7, sc8, xr, yr, zr, xyr, rhot, correction_type, t1, t2, fc)
   ' Else: End If
    '..............
    'end of big loop
    '..............
   End If
    '.............
    
        
''' Do the errors for 207 and 208 corrections (Monte Carlo routine)


If correction_type = "Pb-207" Or correction_type = "Pb-208" Then
Call errors(xt, yt, zt, sxt, syt, szt, ut, sut, c7, c8, sc7, sc8, xr, yr, zr, xyr, rhot, correction_type, t1, t2, fc, st1, sfc, syr, sxr, szr, rhor, sxyr)
Else
End If

  '''''' Calculate final ages and age errors
   
   
               
                stxyr = leadageerr(xyr, sxyr, t2)


        
'errors in U/Th/Pb age schanged in version 2.03 to get consistency with error in uncorrected age:

                tx = Log(1 + xr) / l5
                stx = sxr / (1 + xr) / l5
                
                ty = Log(1 + yr) / l8
                 sty = syr / (1 + yr) / l8
  '170902: zr may go negative in 207-corrected, old zircons
  If zr <= -1 Then zr = 0
                 
                tz = Log(1 + zr) / l2
           
                
                
                 stz = szr / (1 + zr) / l2









''' Start new output routine here ?

'reformat t1 and error to integers, to avoid problems in output
t1 = Int(1000 * t1)
st1 = Int(1000 * st1)


' Comment if one of the ratios goes negative in 207 correction:
'
' 315 - following lines removed
'Set discordance=0 for grains which are concordant within error


'
'

'If conctest(xr, yr, syr, rhor) = True Then discordance = 0
'
'--------
' 315, new:
'315: preliminary discordance calc (testing purposes, could be removed)


'discordance = discord(xr, yr, txyr)

' Cells(j, 32) = "!"
'rimdiscordance = rimdisc(xr, sxr, yr, syr, rhor, ndisc)
''''' 315:
'Move all testing and comment generation for final output here !


'If Abs(initrimdisc) < 0.001 Then
If initrimdisc = 0 Then
correction_type = "None"
Comment = "Concordant"
pass = False
End If

If fc < ndisc * sfc And Comment <> "Concordant" Then
'correction_type = "None"
Comment = "Common Pb < det. lim."
pass = False
End If

If rimdiscordance > 0 Then
    pass = False
    'Comment = "Disc. corr.: Inverse"
            If enable207 = True Then
                correction_type = "207Pb"
                pass = True
            ElseIf presedence = True Then
                Comment = "Discordance test overridden"
                
                pass = True
                ElseIf enable208 = True Then
                Comment = "Disc. & Pb-208 corr.: Inversely discordant"
                
                pass = False
                
                
            Else
                
            End If
      
        ' Else
        ' Comment = "OK2"
End If

    If rimdiscordance <= 0 And Comment = "Discordance test overridden" Then Comment = "OK"

'If correction_type = "None" Then pass = False

If fc < 0 Then pass = False


'--------
'RESET INPUT DATA if no correction can be made
    If pass = False Then
    
        correction_type = "None"
        If Comment <> "Common Pb < det. lim." Then fc = 0
        sfc = 0
        discordance = initial_disc
        rimdiscordance = initrimdisc
        xyr = xyt
        sxyr = sxytobs
        xr = xt
        sxr = sxt
        yr = yt
        syr = syt
        zr = zt
        szr = szt
        txyr = age76
        stxyr = erra76
        tx = age7
        stx = erra7
        ty = age6
        sty = erra6
        tz = age8
        stz = erra8
        rhor = rhot
        t1 = ""
        st1 = ""
        ElseIf correction_type = "Pb-207" Or correction_type = "Pb-208" Then
            t1 = ""
            st1 = ""
    End If
If xr * yr * zr < 0 Or correction_type = "Pb-208" Or correction_type = "Pb-207" Or rhor <= 0 Then Cells(j, 27) = "WARNING !"
'315: Final calculation of discordance
'discordance = discord(xr, yr, txyr)
'rimdiscordance = rimdisc(xr, sxr, yr, syr, rhor, ndisc)


 '***********************
 'Write results to worksheet
 '***********************
        Cells(j, 15) = Format(rhot, "#.##")
        Cells(j, 17) = Format(initial_disc * 100, "###.#")
        Cells(j, 18) = Format(initrimdisc * 100, "###.#")
        Cells(j, 19) = Format(age76 * 1000, "#.")
        Cells(j, 20) = Format(erra76 * 1000, "#.")
        Cells(j, 21) = Format(age7 * 1000, "#.")
        Cells(j, 22) = Format(erra7 * 1000, "#.")
        Cells(j, 23) = Format(age6 * 1000, "#.")
        Cells(j, 24) = Format(erra6 * 1000, "#.")
        Cells(j, 25) = Format(age8 * 1000, "#.")
        Cells(j, 26) = Format(erra8 * 1000, "#.")
        Cells(j, 28) = correction_type
        Cells(j, 29) = Comment
        Cells(j, 30) = Format(100 * fc, "##.##")
        Cells(j, 31) = Format(100 * sfc, "#.##")
        'Cells(j, 32) = Format(100 * loss, "###")
        Cells(j, 33) = Format(100 * discordance, "###.#")
        Cells(j, 34) = Format(100 * rimdiscordance, "###.#")
        Cells(j, 35) = Format(xyr, "#.#####")
        Cells(j, 36) = Format(sxyr, "#.#####")
        Cells(j, 37) = Format(xr, "##.#####")
        Cells(j, 38) = Format(sxr, "#.#####")
        Cells(j, 39) = Format(yr, "#.#####")
        Cells(j, 40) = Format(syr, "#.#####")
        Cells(j, 41) = Format(zr, "#.#####")
        Cells(j, 42) = Format(szr, "#.#####")
        Cells(j, 43) = Format(rhor, "#.##")
        Cells(j, 44) = Format(ut, "##.##")
        Cells(j, 45) = Format(sut, "##.##")
        Cells(j, 46) = t1
        Cells(j, 47) = st1
        
        Cells(j, 49) = Format(1000 * txyr, "####")
        Cells(j, 50) = Format(1000 * stxyr, "###")
        Cells(j, 51) = Format(1000 * tx, "####")
        Cells(j, 52) = Format(1000 * stx, "####")
        Cells(j, 53) = Format(1000 * ty, "####")
        Cells(j, 54) = Format(1000 * sty, "####")
        Cells(j, 55) = Format(1000 * tz, "####")
        Cells(j, 56) = Format(1000 * stz, "####")
        
' tidy up output if 208 has been overcorrected
    
    If zr = 0 Then
        Cells(j, 41) = "<0"
        Cells(j, 42) = ""
        Cells(j, 55) = ""
        Cells(j, 56) = ""
    End If
    
' And redo the age output if errors are less than 1 Ma
If stx < 0.001 Then
        Cells(j, 51) = Format(1000 * tx, "####.#")
        Cells(j, 52) = Format(1000 * stx, "####.#")
End If
If sty < 0.001 Then
        Cells(j, 53) = Format(1000 * ty, "####.#")
        Cells(j, 54) = Format(1000 * sty, "####.#")
End If
If stz < 0.001 Then
        Cells(j, 55) = Format(1000 * tz, "####.#")
        Cells(j, 56) = Format(1000 * stz, "####.#")
End If
If stxyr < 0.001 Then
        Cells(j, 49) = Format(1000 * txyr, "####")
        Cells(j, 50) = Format(1000 * stxyr, "###")
End If

If erra7 < 0.001 Then
        Cells(j, 21) = Format(age7 * 1000, "#.#")
        Cells(j, 22) = Format(erra7 * 1000, "#.#")
End If
If erra6 < 0.001 Then
        Cells(j, 23) = Format(age6 * 1000, "#.#")
        Cells(j, 24) = Format(erra6 * 1000, "#.#")
End If
If erra8 < 0.001 Then
        Cells(j, 25) = Format(age8 * 1000, "#.#")
        Cells(j, 26) = Format(erra8 * 1000, "#.#")
End If

If rhor = 0 Then Cells(j, 43) = ""
 
 '****************
 'Initialize and to next line of data
 ''''''''''''''''''''''''''''''''
 
    j = j + 1
    
 
 
 

    
    Loop
gemoc_on = False
Application.ScreenUpdating = True
End Sub




Function leadage(x, t2)
t1 = 5

Do Until t1 < 0
V = x - (Exp(l5 * t1) - Exp(l5 * t2)) / (Exp(l8 * t1) - Exp(l8 * t2)) / k
DV = -l5 * Exp(l5 * t1) / (Exp(l8 * t1) - Exp(l8 * t2)) / k + (Exp(l5 * t1) - Exp(l5 * t2)) / (Exp(l8 * t1) - Exp(l8 * t2)) ^ 2 / k * l8 * Exp(l8 * t1)
If Abs(V) < 0.000000001 Then Exit Do
t1 = t1 - V / DV
Loop
leadage = t1
End Function




Function t_newton(xt, yt, zt, ut, t2, c7, c8)
'Setting approximate t1

'OBS: correction 17.09.02
If xt < 5 Then t1 = 3 Else t1 = 15



Do Until t1 < 0

  
   

   
    'Expression for A1

    a1 = (yt * (Exp(l5 * t1) - 1) - yt * (Exp(l5 * t2) - 1) - (Exp(l8 * t2) - 1) * (Exp(l5 * t1) - 1) - xt * (Exp(l8 * t1) - 1) _
    + xt * (Exp(l8 * t2) - 1) + (Exp(l5 * t2) - 1) * (Exp(l8 * t1) - 1)) * yt * (-c8 * ut * (Exp(l8 * t1) - 1) + c8 * ut * (Exp(l8 * t2) - 1) _
    + Exp(l2 * t1) - Exp(l2 * t2)) - (-zt * (Exp(l8 * t1) - 1) + zt * (Exp(l8 * t2) - 1) + (Exp(l2 * t2) - 1) * (Exp(l8 * t1) - 1) _
    + yt * (Exp(l2 * t1) - 1) - yt * (Exp(l2 * t2) - 1) - (Exp(l8 * t2) - 1) * (Exp(l2 * t1) - 1)) * yt * (Exp(l5 * t1) - Exp(l5 * t2) - c7 * k * (Exp(l8 * t1) - 1) _
    + c7 * k * (Exp(l8 * t2) - 1))
    

    ' Expression for dA1/dt1

    da1dt = (yt * l5 * Exp(l5 * t1) - (Exp(l8 * t2) - 1) * l5 * Exp(l5 * t1) - xt * l8 * Exp(l8 * t1) + _
    (Exp(l5 * t2) - 1) * l8 * Exp(l8 * t1)) * yt * (-c8 * ut * (Exp(l8 * t1) - 1) + c8 * ut * (Exp(l8 * t2) - 1) _
    + Exp(l2 * t1) - Exp(l2 * t2)) + (yt * (Exp(l5 * t1) - 1) - yt * (Exp(l5 * t2) - 1) - (Exp(l8 * t2) - 1) * (Exp(l5 * t1) - 1) - xt * (Exp(l8 * t1) - 1) _
    + xt * (Exp(l8 * t2) - 1) + (Exp(l5 * t2) - 1) * (Exp(l8 * t1) - 1)) * yt * (-c8 * ut * l8 * Exp(l8 * t1) + l2 * Exp(l2 * t1)) _
    - (-zt * l8 * Exp(l8 * t1) + (Exp(l2 * t2) - 1) * l8 * Exp(l8 * t1) + yt * l2 * Exp(l2 * t1) - (Exp(l8 * t2) - 1) * l2 * Exp(l2 * t1)) * yt * (Exp(l5 * t1) _
    - Exp(l5 * t2) - c7 * k * (Exp(l8 * t1) - 1) + c7 * k * (Exp(l8 * t2) - 1)) - (-zt * (Exp(l8 * t1) - 1) + zt * (Exp(l8 * t2) - 1) _
    + (Exp(l2 * t2) - 1) * (Exp(l8 * t1) - 1) + yt * (Exp(l2 * t1) - 1) - yt * (Exp(l2 * t2) - 1) - (Exp(l8 * t2) - 1) * (Exp(l2 * t1) - 1)) * yt * (l5 * Exp(l5 * t1) _
    - c7 * k * l8 * Exp(l8 * t1))
      
'17.09.02: Removed unnecessary test line and relaxed convergence test to 1E-13 to avoid crash in old zircons


   'If Abs(a1) < 0.00000000000001 Then Exit Do
   
   If Abs(a1) < 0.0000000000001 Then Exit Do
   
   t1 = t1 - a1 / da1dt
    't1 = t1 + (t - t1) / 2
'a2 = a1

  ' i = i + 1

Loop
t_newton = t1
End Function
Function commonlead(xt, yt, zt, t1, t2, c7)
x1 = Exp(l5 * t1) - 1
y1 = Exp(l8 * t1) - 1
z1 = Exp(l2 * t1) - 1
x2 = Exp(l5 * t2) - 1
y2 = Exp(l8 * t2) - 1
z2 = Exp(l2 * t2) - 1

commonlead = (yt * x1 - yt * x2 - y2 * x1 - xt * y1 + xt * y2 + x2 * y1) / (yt * x1 - yt * x2 - yt * c7 * k * y1 + yt * c7 * k * y2)
End Function
Function t207_corr(xt, yt, c7)
'guestimate first t
tI = -2 * (l5 * yt - xt * l8) / yt / l5 / (l5 - l8)
'set scene for newton

F2 = 100
Do Until t1 < 0

F1 = Exp(l5 * t1) - 1 - xt + yt * ((-Exp(l8 * t1) + 1) / yt + 1) * c7 * k

DF1 = l5 * Exp(l5 * t1) - l8 * Exp(l8 * t1) * c7 * k

If Abs(F1) < 0.000001 Then Exit Do
   
t1 = t1 - F1 / DF1
    
F2 = F1

  

Loop
t207_corr = t1

End Function

Function leadageerr(xy, sxy, t2)
tmean = leadage(xy, t2)
' Monte Carlo routine for error in 76 age
    '
    '
    '   Open table of random numbers
    Worksheets("Random").Activate

    txy_var = 0
    For n = 1 To 250
        
        xye = xy + sxy * Cells(n, 8)
        txye = leadage(xye, t2)

    txy_var = txy_var + (txye - tmean) ^ 2
    Next n
    leadageerr = Sqr(txy_var / (n - 1))
Worksheets("data").Activate
ActiveSheet.Select
End Function
Function commont(t, i)
commonp64 = 18.7 - 9.74 * (Exp(l8 * t) - 1)
commonp74 = 15.628 - 9.74 * (Exp(l5 * t) - 1) / 137.88
commonp84 = 38.63 - 36.84 * (Exp(l2 * t) - 1)
common4p = 1 / (1 + commonp64 + commonp74 + commonp84)
common6p = common4p * commonp64
common7p = common4p * commonp74
common8p = common4p * commonp84
If i = 6 Then
commont = common6p
ElseIf i = 7 Then
commont = common7p
ElseIf i = 8 Then
commont = common8p
End If
[aj1] = Format(commonp64, "#.###")
[al1] = Format(commonp74, "#.###")
[an1] = Format(commonp84, "#.###")
End Function

Function conctest(x, y, sy, r)
conctest = False
t_x = Log(1 + x) / l5
y_conc = Exp(t_x * l8) - 1

If Abs(y_conc - y) < Abs(ndisc * (sy * Sqr(1 - r ^ 2))) Then conctest = True

End Function
Function t208_corr(yt, zt, ut, c8)
'initial t1:
t1 = Log(1 + yt) / l8

't1 = 5


Do Until t1 < 0

F1 = Exp(l2 * t1) - 1 - zt + yt * ((-Exp(l8 * t1) + 1) / yt + 1) * c8 * ut

DF1 = l2 * Exp(l2 * t1) - l8 * Exp(l8 * t1) * c8 * ut

If Abs(F1) < 0.000001 Then Exit Do
   
t1 = t1 - F1 / DF1
    


  

    
Loop


t208_corr = t1
End Function

'from here onwards new in 3.15
Function discord(x, y, t)

xconc = Exp(l5 * t) - 1
yconc = Exp(l8 * t) - 1

discord = Sqr((x ^ 2 + y ^ 2) / (xconc ^ 2 + yconc ^ 2)) - 1
End Function
Function rim(x, xm, sx, ym, sy, rho, k)
' Necessary to use an absolute value here, otherwise there may be small, negative values near extreme points
' of ellipsa, causing crash.
Root = Sqr(Abs(rho ^ 2 * sy ^ 2 * xm ^ 2 - 2 * rho ^ 2 * sy ^ 2 * xm * x + rho ^ 2 * sy ^ 2 * x ^ 2 _
- sy ^ 2 * x ^ 2 + 2 * sy ^ 2 * xm * x - sy ^ 2 * xm ^ 2 - rho ^ 2 * sx ^ 2 * sy ^ 2 + sx ^ 2 * sy ^ 2))

rim = (sx * ym - rho * sy * xm + rho * sy * x + k * Root) / sx

End Function

Function rimdisc(xm, sx, ym, sy, rho, ndisc)

' Improved discordance test, which takes the shape of the entire error ellipsa into account, not only its
' extreme points



' Had to define these, otherwise sxt in main program were also multiplied, despite variables being declared !

sxx = ndisc * sx
syy = ndisc * sy

x = xm - 1.1 * sx * ndisc
dx = sx / 20

' Calculation of minimum rim discordance
For i = 1 To 42

x = x + dx
y1 = rim(x, xm, sxx, ym, syy, rho, 1)
y2 = rim(x, xm, sxx, ym, syy, rho, -1)


xy1 = x / y1 / k
xy2 = x / y2 / k

txy1 = leadage(xy1, 0)
txy2 = leadage(xy2, 0)

disc1 = discord(x, y1, txy1)

disc2 = discord(x, y2, txy2)


If disc1 * disc2 < 0.0001 Then

mindisc = 0
 
 
 Else: End If

If Abs(disc1) < Abs(disc2) Then midisc = disc1 Else midisc = disc2
   
    If i = 1 Then
            mindisc = midisc
            
            Else
            
                If Abs(midisc) < Abs(mindisc) Then
               
                mindisc = midisc
             
            Else
            
            End If
    End If
 
Next i


rimdisc = mindisc


End Function

Sub errors(xt, yt, zt, sxt, syt, szt, ut, sut, c7, c8, sc7, sc8, xr, yr, zr, xyr, rhot, correction_type, t1, t2, fc, st1, sfc, syr, sxr, szr, rhor, sxyr)

    '*********************************
    'Error propagation routine follows:
    '*********************************
    
   ' 315 - move to separate routine, to get rid of trouble with preliminary rimdiscordance calc ?
    
    

    ' Monte Carlo routine for error in t1 and fc
    '
    '
    '   Open table of random numbers
                    Worksheets("Random").Activate

                    t1e_var = 0
                    fce_var = 0
                    yre_var = 0
                    xre_var = 0
                    zre_var = 0
                    yxe_cov = 0

                        For n = 10 To 250
        ' set random values
                        xe = xt + sxt * (rhot * Cells(n, 2) + Sqr(1 - rhot ^ 2) * Cells(n, 1))
                        ye = yt + syt * Cells(n, 2)
                        ze = zt + szt * Cells(n, 3)
                        ute = ut * (1 + Cells(n, 4) / 100)
                        c7e = c7 * (1 + Cells(n, 5) / 100)
                        c8e = c8 * (1 + Cells(n, 6) / 100)
        'calculate
        
        
                        If correction_type = "Disc" Then
        
                            t1e = t_newton(xe, ye, ze, ute, t2, c7e, c8e)
                            fce = commonlead(xe, ye, ze, t1e, t2, c7)
            
                                ElseIf correction_type = "Pb-207" Then
            
                                t1e = t207_corr(xe, ye, c7)
                               
                                fce = (ye - (Exp(l8 * t1e) - 1)) / ye
                                
                                 ElseIf correction_type = "Pb-208" Then
                                t1e = t208_corr(xe, ye, ut, c8)
                               
                                fce = (ye - (Exp(l8 * t1e) - 1)) / ye
                        End If
            

                            yre = ye * (1 - fce)
                            xre = xe - k * c7e * fce * ye
                            zre = ze - ute * c8e * fce * ye

       
                         t1e_var = t1e_var + (t1e - t1) ^ 2
                        fce_var = fce_var + (fce - fc) ^ 2

                            yre_var = yre_var + (yre - yr) ^ 2
                            xre_var = xre_var + (xre - xr) ^ 2
                            zre_var = zre_var + (zre - zr) ^ 2
                            yxe_cov = yxe_cov + (yre - yr) * (xre - xr)


                        Next n


                     



    'do the average and standard deviation calculation
                st1 = Sqr(t1e_var / (n - 1))
                sfc = Sqr(fce_var / (n - 1))
                sxr = Sqr(xre_var / (n - 1))
                syr = Sqr(yre_var / (n - 1))
                szr = Sqr(zre_var / (n - 1))
               yxe_cov = yxe_cov / n
                rhor = yxe_cov / sxr / syr

                sxyr = xyr * (sxr ^ 2 / xr ^ 2 + syr ^ 2 / yr ^ 2 - 2 * A * rhor * (syr / yr) * (sxr / xr)) ^ 0.5
    
                    If correction_type = "Pb-207" Then
     
                        sfc = (1 / yt ^ 2 / c7 ^ 2 / k ^ 2 * sxt ^ 2 + (xt - xr) ^ 2 / yt ^ 4 / c7 ^ 2 / k ^ 2 * syt ^ 2 + (xt - xr) ^ 2 / yt ^ 2 / c7 ^ 4 / k ^ 2 * sc7 ^ 2) ^ 0.5
                    ElseIf correction_type = "Pb-208" Then
                    sfc = (1 / yt ^ 2 / c8 ^ 2 / ut ^ 2 * szt ^ 2 + (zt - zr) ^ 2 / yt ^ 4 / c8 ^ 2 / ut ^ 2 * syt ^ 2 + (zt - zr) ^ 2 / yt ^ 2 / c8 ^ 4 / ut ^ 2 * sc8 ^ 2) ^ 0.5
                    
                    End If
    
' get back to data worksheet

                Worksheets("data").Activate
                ActiveSheet.Select



    
    
    
    
   
    
End Sub
