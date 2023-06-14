DECLARE SUB inputS (S#, C#, mK#, mNa#, mCa#, mBa#)
DECLARE SUB inputS2 (S#, mK#, mNa#, mCa#)
DECLARE SUB graph (fOstart#, fOend#)
DECLARE SUB water (aH2O#, ion#, T#)
DECLARE SUB debye (T#, Nion%, charge#(), iion#, alpha#(), gamma#())
DECLARE SUB neutral (X#, ion#, T#)

DEFINT I-N
DEF FNLOG# (X#) = LOG(X#) / LOG(10#): DEF FNEXP# (X#) = EXP(X# * LOG(10#))
CONST M = 25: CONST N = 255: A$ = "#####.########": B$ = "#####.##"
DIM IonName$(M), alpha#(M),charge#(M), LogK$(M), A0#(M), A1#(M), A2#(M), A3#(M), A4#(M)
DIM gamma#(M), LfO2#(N), fO2#(N), pH1(N), pH2(N)
COLOR 14, 9: CLS : LOCATE 5, 1

'1. Ion parameters for Debye-Huckel equation (a & z)
OPEN "data.1" FOR INPUT AS #1
   LINE INPUT #1, Temp1$: LINE INPUT #1, Temp2$
FOR I = 1 TO M
 INPUT #1, IonName$(I), alpha#(I), charge#(I)
 IF EOF(1) THEN EXIT FOR
NEXT I
CLOSE #1: Nion% = I

'2. coefficients for Log K
OPEN "data.2" FOR INPUT AS #1
   LINE INPUT #1, Temp1$
FOR I = 1 TO M
 INPUT #1, LogK$(I), A0#(I), A1#(I), A2#(I), A3#(I), A4#(I)
 IF EOF(1) THEN EXIT FOR
NEXT I
CLOSE #1

10 CLS : LOCATE 10, 1
PRINT "You are about to input some parameters': PRINT
INPUT " Please press RETURN to continue"; aa$: PRINT : PRINT
INPUT "INPUT the temperature = = > "; T#: PRINT
CALL inputS2(TotalS#, mK#, mNa#, mCa#)
ion# = .5# * (mK# + mNa# + 4# * (mCa# + mBa#)): iion# = SQR(ion#)
T2# = T# * T#: T3# = T# * T2#: T4# = T# * T3#
D3# = .001# * T#: D5# = .00001# * T2#: D7# = .0000001# * T3#: D9# = .000000001# * T4#

CALL debye(T#, Nion%, charge#(), iion#, alpha#(), gamma#())
gammaHS# = gamma#(1): gammaHSO4# = gamma#(2): gammaSO4# = gamma#(3)
gammaKSO4# = gamma#(4): gammaK# = gamma#(5)
gammaNaSO4# = gamma#(6): gammaNa# = gamma#(7)
gxmmaCa = gamma#(8): gammaCaSO4# = 1#
CALL neutral(gammaH2S#, ion#, T#)
kH2S# = FNEXP#(A0#(1) + A1#(1) * D3# + A2#(1) * D5# + A3#(1) * D7# + A4#(1) * D9#)
kHSO4# = FNEXP#(A0#(2) + Al#(2) * D3# + A2#(2) * D5# + A3#(2) * D7# + A4#(2) * D9#)
kSO4# = FNEXP#(A0#(3) + Al#(3) * D3# + A2#(3) * D5# + A3#(3) * D7# + A4#(3) * D9#)
kKSO4# = FNEXP#(A0#(4) + A1#(4) * D3# + A2#(4) * D5# + A3#(4) * D7# + A4#(4) * D9#)
kNaSO4# = FNEXP#(A0#(5) + A1#(5) * D3# + A2#(5) * D5# + A3#(5) * D7# + A4#(5) * D9#)
kCaSO4# = FNEXP#(A0#(6) + A1#(6) * D3# + A2#(6) * D5# + A3#(6) * D7# + A4#(6) * D9#)

range% = 250
IF T# > 239# THEN
   fOstart# = 25#: fOend# = 50#
  ELSEIF T# > 179 THEN
   fOstart# = 30#: fOend# = 55#
   ELSE
   fOstart# = 35#: fOend# = 60#
   END IF

30 CLS : LOCATE 10, 1
PRINT "You can choose one of the following four sulfur-bearing species"
PRINT "to calculate its Delta-S34 distribution in log fO2-pH plane': PRINT
PRINT " 1. H2S(aqueous)"
PRINT " 2. Pyrite"
PRINT "3. Sphalerite"
PRINT " 4. Galena"
PRINT " 5. Sulfate (Barite, Anhydrite etc.)": PRINT
INPUT "Choose one (enter the code of the species) = = > "; ans%: PRINT
INPUT "INPUT: Delta $34 of this species = = > "; S34#
CLS : LOCATE 12, 1
INPUT "INPUT: Delta S34 of the Total S = = > "; TotS34#

TK# = T# + 273.15#: TK2# = TK# * TK#
DelHS# = -60000# / TK2# - .6#: DelSO4# = 5260000# / TK2# + 6#
DelPy# = 400000# / TK2#: DelSph = 100000# / TK2#: DelGa# = -630000# / TK2#

SELECT CASE ans%
  CASE 1: S34H2S# = S34#
  CASE 2: S34H2S# = S34# - Delpy#
  CASE 3: S34H2S# = S34# - DelSph#
  CASE 4: S34H2S# = S34# - DelGa#
  CASE 5: S34H2S# = S34# - DelSO4#
END SELECT

DeltaS# = S34H2S# - TotS34#
IF (DeltaS# > -DelHS# OR DeltaS# < -DelSO4#) THEN
  CLS : LOCATE 12, 1: PRINT " Your input data:': PRINT
  PRINT " DeltaS34 (mineral) - DeltaS34 (total Sulfur)": PRINT
  PRINT " is impossible.": PRINT
  INPUT " Please RE-ENTER your data. Press RETURN to continue."; dum$
  GOTO 30
END IF
IF DeltaS# = 0# THEN DeltaS# = -.001# ELSE DeltaS# = DeltaS#
CLS : LOCATE 12, 1
PRINT "PROGRAM is in progress, please wait."

FOR J = 0 TO range%
   zz# = CDBL(J) / 10#
 LfO2#(J) = -(zz# + fOstart#)
 fO2#(J) = FNEXP#(LfO2#(J)): fO2#(J) = fO2#(J) * fO2#(J)
   aa# = kH2S# * gammaH2S# / gammaHS#
   cc# = kSO4# * gammaH2S# * fO2#(J) / gammaSO4#
   dd# = gammaSO4# / kHSO4# / gammaHSO4#
   ee# = mK# * gammaK# * gammaSO4# / kKSO4# / gammaKSO4#
   ff# = mNa# * gammaNa# * gammaSO4# / kNaSO4# / gammaNaSO4#
   gg# = mCa# * gammaCa# * gammaSO4# / kCaSO4# / gammaCaSO4#
     Temp1# = aa# + cc# * dd#
     Temp2# = cc# * (1# + ee# + ff# + gg#)
     Temp11# = DelHS# * aa# + DelSO4# * cc# * dd#
     Temp22# = DelSO4# * cc# * (1# + ee# + ff# + gg#)
 aTemp# = DeltaS#
 bTemp# = DeltaS# * Temp1# + Temp11#
 cTemp# = DeltaS# * Temp2# + Temp22#
    Test# = bTemp# * bTemp# - 4# * aTemp# * cTemp#
 IF Test# < 0# THEN
       pH1(J) = -100: pH2(J) = -100
  ELSE
     aH1# = (-bTemp# + SQR(Test#)) / 2# / aTemp#
     aH2# = (-bTemp# - SQR(Test#)) / 2# / aTemp#
 END IF
  IF aH1# < 0# THEN pH1(J) = -I00 ELSE pH1(J) = -FNLOG#(aH1#)
  IF aH2# < 0# THEN pH2(J) = -I00 ELSE pH2(J) = -FNLOG#(aH2#)
NEXT J
FOR J = range% TO 1 STEP -1
  IF pH1(J) < 0 THEN pH1(J - 1) = pH1(J)
NEXT J

CALL graph(fOstart#, fOend#)
DO UNTIL INKEY$ <> ""
FOR J = range% TO 0 STEP -1
 IF aTemp# > 0 AND (pH1(J) > 0 AND pH1(J) < 14) THEN
     LINE (pH1(J) * 10, LfO2#(J))-(pH1(J) * 10, LfO2#(J))
 END IF
NEXT J
FOR J = 0 TO range%
 IF (pH2(J) > 0 AND pH2(J) < 14) THEN
      LINE (pH2(J) * 10, LfO2#(J))-(pH2(J) * 10, LfO2#(J))
 END IF
NEXT J
LOOP

SCREEN 0: COLOR 14, 9: CLS : LOCATE 15, 1
PRINT "If you need output files, press (Y), then RETURN;': PRINT
INPUT "IF you do not need output files, press (N), then RETURN = = > "; ans$
IF ans$ = "N" OR ans$ = "n" THEN GOTO 1000 ELSE GOTO 300
300 PRINT : INPUT "Output File Name --- > "; name$
OPEN name$ FOR OUTPUT AS #1

FOR J = range% TO 0 STEP -1
 IF aTemp# > 0 AND (pH1(J) > 0 AND pH1(J) < 14) THEN
     PRINT #1, USING A$; pH1(J); LfO2#(J)
 END IF
NEXT J
FOR J = 0 TO range%
 IF (pH2(J) > 0 AND pH2(J) < 14) THEN
    PRINT #1, USING A$; pH2(J); LfO2#(J)
 END IF
NEXT J
CLOSE #1

1000 PRINT : INPUT "More calculation ? (Y) or (N) "; answer$
IF answer$ = "Y" OR answer$ = "y" THEN GOTO 10 ELSE GOTO 10000

10000 END

DEFSNG I-N
SUB debye (T#, Nion%, charge#(), iion#, alpha#(), gamma#())

T2# = T# * T#: T3# = T# * T2#: T4# = T# * T3#
D3# = .001# * T#: D5# = .00001# * T2#: D7# = .0000001# * T3#: D9# = .000000001# * T4#

'********* Debye-Huckel coefficient A, B & b ********

A# = .512144# - 1.13612# * D3# + 3.4718# * D5# - 1.75238 * D7# + .333889# * D9#
B# = .325337# + .0793621# * D3# + .143473# * D5# - .0644803# * D7# + .0114498 * D9#
beta# = .0428049# - .0962915# * D3# + .199187# * D5# - .072107# * D7#

FOR I = 1 TO Nion%
  gamma#(I) = -A# * charge#(I) * charge#(I) * iion#
  gamma#(I) = gamma#(I) / (1# + alpha#(I) * B# * iion#) + beta# * ion%
  gamma#(I) = FNEXP#(gamma#(I))
NEXT I
END SUB

SUB graph (fOstart#, fOend#)

' ********************* Graphic ***********************

SCREEN 9
fOstart = CINT(fOstart#): fOend = CINT(fOend#)
ran% = (fOend# - fOstart#) / 5 - 1
VIEW (I00, 10)-(550, 250), 12, 1
WINDOW (0, -fOend#)-(140, -fOstart#)
FOR I= 1 TO 13
  LINE (10 * I, -fOend#)-(10 * I, -fOend# + .5#)
NEXT I
FOR I = 1 TO ran%
  LINE (0, -fOend# + 5 * I)-(2, -fOend# + 5 * I)
NEXT I

LOCATE 20, 40: PRINT "pH': LOCATE 19, 41: PRINT "7"
LOCATE 19, 13: PRINT "0": LOCATE 19, 69: PRINT "14"
LOCATE 10, 1: PRINT "log fO2"
LOCATE 1, 9: PRINT -fOstart#: LOCATE 18, 9: PRINT -fOend#
LOCATE 22, 1: PRINT "press any key to continue "

END SUB

SUB inputS2 (S#, mK#, mNa#, mCa#)
CLS : LOCATE 10, 1
PRINT "Input chemical data, please USING molal unit"
PRINT
INPUT "INPUT: total sulfur in the system = = > "; TotalS#: PRINT
CLS : LOCATE 10, 1
INPUT "INPUT: mK+ == > "; mK#: PRINT
INPUT "INPUT: mNa+ = = > "; mNa#: PRINT
INPUT "INPUT: mCa++ = = > "; mCa#: PRINT
END SUB

SUB neutral (X#, ion#, T#)

T2# = T# * T#: T3# = T# * T2#: T4# = T# * T3#
 Temp1# = -.7944830201999999# + 3.598014883# * ion# - .0000009433# * T#
 Temp2# = Temp1# - 1.8395350711# * ion# * ion# - .0000136098# * T2#
 Temp3# = Temp2# + .0001133749# * ion# * T# + .3175818606# * ion# * ion# * ion#
 Temp4# = Temp3# + .000000039# * T3# - .0004620397# * ion# * ion# * T#
X# = Temp4# + .0000058337# * ion# * T2#
END SUB