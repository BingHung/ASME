Module FlangDesign

    '|Input|
    Dim π As Single
    'Dim objWs, objWsTbl As Object, intRow As Integer, intClm As Integer
    Dim P, Ca, S, B, g0, Fs, Sf, m, y, tg, w, N, N1, Sb, Sa, nb, Db As Double
    Dim GSTP, GSSK, BOTP As String
    Dim bo As Double
    Dim intTblRow, intTblClm As Integer
    Dim g1, h, A, tf As Double
    Dim intCaseCounter As Integer
    Dim aResult(11) As Double
    Dim Ar, Ba, Rh, Eb, C As Double
    Dim Dog, Di, bg, G As Double
    Dim Wm1, Wm2, Ab, Am As Double

    Sub Main()

        π = 4 * 0.785398163
        P = 10
        'T = 65
        Ca = 3
        S = 14.07
        B = 540
        g0 = 9
        Fs = 0.9
        Sf = 14.07
        GSTP = "1a"
        GSSK = "II"
        m = 3
        y = 7.03
        tg = 4.5
        w = 0
        N = 0
        N1 = 0
        Sa = 17.54
        Sb = 17.54
        BOTP = "mm"
        Db = 24

        '| Gasket Diameter|

        If N = 0 Then
            N = GasketWeight(B)
        End If

        bo = GasketRequiredWidth(GSTP, GSSK, N, w, tg)

        '| Bolt minimum Diameter|

        If Db = 0 Then
            If BOTP = "mm" Then
                Db = 20
            Else
                Db = 0.5
            End If
        End If

        If BOTP = "mm" Then
            intTblRow = 1
            intTblClm = 9
        Else
            intTblRow = 1
            intTblClm = 2
        End If

        '| Flange Calculation Loop|

        aResult(0) = 10 ^ 100
        B = B + 2 * Ca
        For g1 = (g0 + 1) To Carry(3 * g0) Step 1
            For h = Carry(1.5 * g0) To Carry(5 * g0) Step 1
                If h >= 3 * (g1 - g0) Then
                    For tf = Max(30, 1.5 * g0) To (15 * g0) Step 1
                        PROC_BCDCheck()
                        A = Carry(C + 2 * Eb)
                        g0 = g0 - Ca
                        g1 = g1 - Ca
                        PROC_Flange()
                        g0 = g0 + Ca
                        g1 = g1 + Ca
                    Next tf
                End If
            Next h
        Next g1
        B = B - 2 * Ca

    End Sub


    Sub PROC_BCDCheck()

        Dim ASME_Bolt_Dimension_Flag As Double = 0

        Do
            '-------------------------------------------------------------------------------------------
            ' Bolt Diameter
            '-------------------------------------------------------------------------------------------

            'With objWsTbl
            '    .Cells(intTblRow, intTblClm).Value = Db
            '    Ar = .Cells(intTblRow, intTblClm + 1).Value
            '    Ba = .Cells(intTblRow, intTblClm + 2).Value
            '    Rh = .Cells(intTblRow, intTblClm + 3).Value
            '    Eb = .Cells(intTblRow, intTblClm + 4).Value
            'End With

            Dim ASME_Bolt_Dimension(,) As Double = {
            {12, 72.4, 32, 21, 16},
            {16, 138.3, 44, 29, 21},
            {20, 217.1, 52, 32, 24},
            {22, 272.4, 54, 33, 25},
            {24, 312.7, 59, 37, 29},
            {27, 413.9, 64, 39, 29},
            {30, 562.1, 73, 46, 33},
            {33, 675.1, 78, 48, 35},
            {36, 842.5, 84, 54, 40},
            {42, 1179.4, 100, 62, 49},
            {48, 1572.9, 113, 68, 56},
            {56, 2185.6, 127, 76, 64},
            {64, 2898.8, 140, 84, 67},
            {72, 3712.5, 156, 89, 70},
            {80, 4626.7, 167, 94, 75},
            {90, 5910.8, 189, 108, 85},
            {100, 7352.1, 208, 120, 94}
            }



            Select Case Db
                Case Is = ASME_Bolt_Dimension(0, 0)
                    ASME_Bolt_Dimension_Flag = 0
                    Ar = ASME_Bolt_Dimension(0, 1)
                    Ba = ASME_Bolt_Dimension(0, 2)
                    Rh = ASME_Bolt_Dimension(0, 3)
                    Eb = ASME_Bolt_Dimension(0, 4)
                Case Is = ASME_Bolt_Dimension(1, 0)
                    ASME_Bolt_Dimension_Flag = 1
                    Ar = ASME_Bolt_Dimension(1, 1)
                    Ba = ASME_Bolt_Dimension(1, 2)
                    Rh = ASME_Bolt_Dimension(1, 3)
                    Eb = ASME_Bolt_Dimension(1, 4)
                Case Is = ASME_Bolt_Dimension(2, 0)
                    ASME_Bolt_Dimension_Flag = 2
                    Ar = ASME_Bolt_Dimension(2, 1)
                    Ba = ASME_Bolt_Dimension(2, 2)
                    Rh = ASME_Bolt_Dimension(2, 3)
                    Eb = ASME_Bolt_Dimension(2, 4)
                Case Is = ASME_Bolt_Dimension(3, 0)
                    ASME_Bolt_Dimension_Flag = 3
                    Ar = ASME_Bolt_Dimension(3, 1)
                    Ba = ASME_Bolt_Dimension(3, 2)
                    Rh = ASME_Bolt_Dimension(3, 3)
                    Eb = ASME_Bolt_Dimension(3, 4)
                Case Is = ASME_Bolt_Dimension(4, 0)
                    ASME_Bolt_Dimension_Flag = 4
                    Ar = ASME_Bolt_Dimension(4, 1)
                    Ba = ASME_Bolt_Dimension(4, 2)
                    Rh = ASME_Bolt_Dimension(4, 3)
                    Eb = ASME_Bolt_Dimension(4, 4)
                Case Is = ASME_Bolt_Dimension(5, 0)
                    ASME_Bolt_Dimension_Flag = 5
                    Ar = ASME_Bolt_Dimension(5, 1)
                    Ba = ASME_Bolt_Dimension(5, 2)
                    Rh = ASME_Bolt_Dimension(5, 3)
                    Eb = ASME_Bolt_Dimension(5, 4)
                Case Is = ASME_Bolt_Dimension(6, 0)
                    ASME_Bolt_Dimension_Flag = 6
                    Ar = ASME_Bolt_Dimension(6, 1)
                    Ba = ASME_Bolt_Dimension(6, 2)
                    Rh = ASME_Bolt_Dimension(6, 3)
                    Eb = ASME_Bolt_Dimension(6, 4)
                Case Is = ASME_Bolt_Dimension(7, 0)
                    ASME_Bolt_Dimension_Flag = 7
                    Ar = ASME_Bolt_Dimension(7, 1)
                    Ba = ASME_Bolt_Dimension(7, 2)
                    Rh = ASME_Bolt_Dimension(7, 3)
                    Eb = ASME_Bolt_Dimension(7, 4)
                Case Is = ASME_Bolt_Dimension(8, 0)
                    ASME_Bolt_Dimension_Flag = 8
                    Ar = ASME_Bolt_Dimension(8, 1)
                    Ba = ASME_Bolt_Dimension(8, 2)
                    Rh = ASME_Bolt_Dimension(8, 2)
                    Eb = ASME_Bolt_Dimension(8, 4)
                Case Is = ASME_Bolt_Dimension(9, 0)
                    ASME_Bolt_Dimension_Flag = 9
                    Ar = ASME_Bolt_Dimension(9, 1)
                    Ba = ASME_Bolt_Dimension(9, 2)
                    Rh = ASME_Bolt_Dimension(9, 3)
                    Eb = ASME_Bolt_Dimension(9, 4)
                Case Is = ASME_Bolt_Dimension(10, 0)
                    ASME_Bolt_Dimension_Flag = 10
                    Ar = ASME_Bolt_Dimension(10, 1)
                    Ba = ASME_Bolt_Dimension(10, 2)
                    Rh = ASME_Bolt_Dimension(10, 3)
                    Eb = ASME_Bolt_Dimension(10, 4)
                Case Is = ASME_Bolt_Dimension(11, 0)
                    ASME_Bolt_Dimension_Flag = 11
                    Ar = ASME_Bolt_Dimension(11, 1)
                    Ba = ASME_Bolt_Dimension(11, 2)
                    Rh = ASME_Bolt_Dimension(11, 3)
                    Eb = ASME_Bolt_Dimension(11, 4)
                Case Is = ASME_Bolt_Dimension(12, 0)
                    ASME_Bolt_Dimension_Flag = 12
                    Ar = ASME_Bolt_Dimension(12, 1)
                    Ba = ASME_Bolt_Dimension(12, 2)
                    Rh = ASME_Bolt_Dimension(12, 3)
                    Eb = ASME_Bolt_Dimension(12, 4)
                Case Is = ASME_Bolt_Dimension(13, 0)
                    ASME_Bolt_Dimension_Flag = 13
                    Ar = ASME_Bolt_Dimension(13, 1)
                    Ba = ASME_Bolt_Dimension(13, 2)
                    Rh = ASME_Bolt_Dimension(13, 3)
                    Eb = ASME_Bolt_Dimension(13, 4)
                Case Is = ASME_Bolt_Dimension(14, 0)
                    ASME_Bolt_Dimension_Flag = 14
                    Ar = ASME_Bolt_Dimension(14, 1)
                    Ba = ASME_Bolt_Dimension(14, 2)
                    Rh = ASME_Bolt_Dimension(14, 3)
                    Eb = ASME_Bolt_Dimension(14, 4)
                Case Is = ASME_Bolt_Dimension(15, 0)
                    ASME_Bolt_Dimension_Flag = 15
                    Ar = ASME_Bolt_Dimension(15, 1)
                    Ba = ASME_Bolt_Dimension(15, 2)
                    Rh = ASME_Bolt_Dimension(15, 3)
                    Eb = ASME_Bolt_Dimension(15, 4)
                Case Is = ASME_Bolt_Dimension(16, 0)
                    ASME_Bolt_Dimension_Flag = 16
                    Ar = ASME_Bolt_Dimension(16, 1)
                    Ba = ASME_Bolt_Dimension(16, 2)
                    Rh = ASME_Bolt_Dimension(16, 3)
                    Eb = ASME_Bolt_Dimension(16, 4)
                Case Else
                    MsgBox("Bolt Diameter Out of Range!")
            End Select


            C = Carry(B + 2 * (Max(g1 + Rh, N + N1 + 8 + Db / 2) - Ca))
            Dog = C - Db - 16
            Di = Dog - 2 * N

            If bo > 6.35 Then
                bg = 2.5 * (bo) ^ 0.5
                G = Dog - 2 * bg
            Else
                bg = bo
                G = Dog - N
            End If

            '-------------------------------------------------------------------------------------------
            ' Bolt Area and number Check
            '-------------------------------------------------------------------------------------------

            Dim Bar As Integer, Bmax As Integer

            Wm1 = 0.7854 * G ^ 2 * P / 100 + 6.28 * bg * G * m * P / 100
            Wm2 = 3.14 * bg * G * y
            Am = Max(Wm1 / Sb, Wm2 / Sa) / Fs
            nb = Carry4(Am / (Ar * 4))
            Bar = Int(π * C / nb)
            Bmax = Int(2 * Db + 6 * tf / (m + 0.5))
            If Bar > Bmax Then
                nb = Carry4((π * C) / (Bmax * 4))
            End If
            If Int(π * C / nb) < Ba Then
                ASME_Bolt_Dimension_Flag = ASME_Bolt_Dimension_Flag + 1
                Db = ASME_Bolt_Dimension(ASME_Bolt_Dimension_Flag, 0)
            Else
                Exit Do
            End If
        Loop

    End Sub

    Sub PROC_Flange()
        '---------------------------------------------------------------------------------------------------
        ' Moment of bolting up
        '---------------------------------------------------------------------------------------------------

        Dim Hb, Hp, R, Wg, Mb As Double

        Hb = Math.Round(0.7854 * G ^ 2 * P / 100)
        Hp = Math.Round(2 * π * bg * G * m * P / 100)
        Ab = nb * Ar
        R = (C - B) / 2 - g1
        Wg = Math.Round((Am + Ab) * Sa / 2)
        Mb = Math.Round((C - G) * Wg / 2)

        '---------------------------------------------------------------------------------------------------
        ' Morment of operating
        '---------------------------------------------------------------------------------------------------

        Dim hga, hda, hta, Hd, Hg, Ht, Md, Mg, Mt, Mo, Mmax As Double

        hga = (C - G) / 2
        hda = R + 0.5 * g1
        hta = (R + g1 + hga) / 2
        Hd = 0.785 * B ^ 2 * P / 100
        Hg = Wm1 - Hb
        Ht = Hb - Hd
        Md = Math.Round(Hd * hda)
        Mg = Math.Round(Hg * hga)
        Mt = Math.Round(Ht * hta)
        Mo = Md + Mg + Mt

        Mmax = Max(Mb, Mo)

        '---------------------------------------------------------------------------------------------------
        ' Factor of flange
        '---------------------------------------------------------------------------------------------------

        Dim K, ho, AA, Z, T, U, YY, FC(37), FE(10), FF, V, f, e, D, L As Double

        K = A / B
        ho = (B * g0) ^ 0.5
        AA = (g1 / g0) - 1
        Z = (K ^ 2 + 1) / (K ^ 2 - 1)
        T = (K ^ 2 * (1 + 8.55246 * Math.Log(K) / Math.Log(10.0#)) - 1) / ((1.0472 + 1.9448 * K ^ 2) * (K - 1))
        U = (K ^ 2 * (1 + 8.55246 * Math.Log(K) / Math.Log(10.0#)) - 1) / (1.36136 * (K ^ 2 - 1) * (K - 1))
        YY = (0.66845 + 5.7169 * (K ^ 2 * Math.Log(K) / Math.Log(10.0#)) / (K ^ 2 - 1)) / (K - 1)

        FC(0) = 43.68 * (h / ho) ^ 4
        FC(1) = 1 / 3 + AA / 12
        FC(2) = 5 / 42 + 17 * AA / 336
        FC(3) = 1 / 210 + AA / 360
        FC(4) = 11 / 360 + 59 * AA / 5040 + (1 + 3 * AA) / FC(0)
        FC(5) = 1 / 90 + 5 * AA / 1008 - (1 + AA) ^ 3 / FC(0)
        FC(6) = 1 / 120 + 17 * AA / 5040 + 1 / FC(0)
        FC(7) = 215 / 2772 + 51 * AA / 1232 +
                    (60 / 7 + 225 * AA / 14 + 75 * AA ^ 2 / 7 + 5 * AA ^ 3 / 2) / FC(0)
        FC(8) = 31 / 6930 + 128 * AA / 45045 +
                    (6 / 7 + 15 * AA / 7 + 12 * AA ^ 2 / 7 + 5 * AA ^ 3 / 11) / FC(0)
        FC(9) = 533 / 30240 + 653 * AA / 73920 +
                    (1 / 2 + 33 * AA / 14 + 39 * AA ^ 2 / 28 + 25 * AA ^ 3 / 84) / FC(0)
        FC(10) = 29 / 3780 + 3 * AA / 704 -
                    (1 / 2 + 33 * AA / 14 + 81 * AA ^ 2 / 28 + 13 * AA ^ 3 / 12) / FC(0)
        FC(11) = 31 / 6048 + 1763 * AA / 665280 +
                    (1 / 2 + 6 * AA / 7 + 15 * AA ^ 2 / 28 + 5 * AA ^ 3 / 42) / FC(0)
        FC(12) = 1 / 2925 + 71 * AA / 300300 +
                    (8 / 35 + 18 * AA / 35 + 156 * AA ^ 2 / 385 + 6 * AA ^ 3 / 55) / FC(0)
        FC(13) = 761 / 831600 + 937 * AA / 1663200 +
                    (1 / 35 + 6 * AA / 35 + 11 * AA ^ 2 / 70 + 3 * AA ^ 3 / 70) / FC(0)
        FC(14) = 197 / 415800 + 103 * AA / 332640 -
                    (1 / 35 + 6 * AA / 35 + 17 * AA ^ 2 / 70 + AA ^ 3 / 10) / FC(0)
        FC(15) = 233 / 831600 + 97 * AA / 554400 +
                    (1 / 35 + 3 * AA / 35 + AA ^ 2 / 14 + 2 * AA ^ 3 / 105) / FC(0)
        FC(16) = FC(1) * FC(7) * FC(12) + FC(2) * FC(8) * FC(3) + FC(3) * FC(8) * FC(2) -
                    (FC(3) ^ 2 * FC(7) + FC(8) ^ 2 * FC(1) + FC(2) ^ 2 * FC(12))
        FC(17) = (FC(4) * FC(7) * FC(12) + FC(2) * FC(8) * FC(13) + FC(3) * FC(8) * FC(9) -
                    (FC(13) * FC(7) * FC(3) + FC(8) ^ 2 * FC(4) + FC(12) * FC(2) * FC(9))) / FC(16)
        FC(18) = (FC(5) * FC(7) * FC(12) + FC(2) * FC(8) * FC(14) + FC(3) * FC(8) * FC(10) -
                    (FC(14) * FC(7) * FC(3) + FC(8) ^ 2 * FC(5) + FC(12) * FC(2) * FC(10))) / FC(16)
        FC(19) = (FC(6) * FC(7) * FC(12) + FC(2) * FC(8) * FC(15) + FC(3) * FC(8) * FC(11) -
                    (FC(15) * FC(7) * FC(3) + FC(8) ^ 2 * FC(6) + FC(12) * FC(2) * FC(11))) / FC(16)
        FC(20) = (FC(1) * FC(9) * FC(12) + FC(4) * FC(8) * FC(3) + FC(3) * FC(13) * FC(2) -
                    (FC(3) ^ 2 * FC(9) + FC(13) * FC(8) * FC(1) + FC(12) * FC(4) * FC(2))) / FC(16)
        FC(21) = (FC(1) * FC(10) * FC(12) + FC(5) * FC(8) * FC(3) + FC(3) * FC(14) * FC(2) -
                    (FC(3) ^ 2 * FC(10) + FC(14) * FC(8) * FC(1) + FC(12) * FC(5) * FC(2))) / FC(16)
        FC(22) = (FC(1) * FC(11) * FC(12) + FC(6) * FC(8) * FC(3) + FC(3) * FC(15) * FC(2) -
                    (FC(3) ^ 2 * FC(11) + FC(15) * FC(8) * FC(1) + FC(12) * FC(6) * FC(2))) / FC(16)
        FC(23) = (FC(1) * FC(7) * FC(13) + FC(2) * FC(9) * FC(3) + FC(4) * FC(8) * FC(2) -
                    (FC(3) * FC(7) * FC(4) + FC(8) * FC(9) * FC(1) + FC(2) ^ 2 * FC(13))) / FC(16)

        FC(24) = (FC(1) * FC(7) * FC(14) + FC(2) * FC(10) * FC(3) + FC(5) * FC(8) * FC(2) -
                    (FC(3) * FC(7) * FC(5) + FC(8) * FC(10) * FC(1) + FC(2) ^ 2 * FC(14))) / FC(16)
        FC(25) = (FC(1) * FC(7) * FC(15) + FC(2) * FC(11) * FC(3) + FC(6) * FC(8) * FC(2) -
                    (FC(3) * FC(7) * FC(6) + FC(8) * FC(11) * FC(1) + FC(2) ^ 2 * FC(15))) / FC(16)
        FC(26) = -(FC(0) / 4) ^ 0.25
        FC(27) = FC(20) - FC(17) - 5 / 12 + FC(17) * FC(26)
        FC(28) = FC(22) - FC(19) - 1 / 12 + FC(19) * FC(26)
        FC(29) = -(FC(0) / 4) ^ 0.5
        FC(30) = -(FC(0) / 4) ^ 0.75
        FC(31) = 3 * AA / 2 - FC(17) * FC(30)
        FC(32) = 1 / 2 - FC(19) * FC(30)
        FC(33) = 0.5 * FC(26) * FC(32) + FC(28) * FC(31) * FC(29) -
                    (0.5 * FC(30) * FC(28) + FC(32) * FC(27) * FC(29))
        FC(34) = 1 / 12 + FC(18) - FC(21) - FC(18) * FC(26)
        FC(35) = -FC(18) * (FC(0) / 4) ^ 0.75
        FC(36) = (FC(28) * FC(35) * FC(29) - FC(32) * FC(34) * FC(29)) / FC(33)
        FC(37) = (0.5 * FC(26) * FC(35) + FC(34) * FC(31) * FC(29) -
                    (0.5 * FC(30) * FC(34) + FC(35) * FC(27) * FC(29))) / FC(33)
        FE(1) = FC(17) * FC(36) + FC(18) + FC(19) * FC(37)
        FE(2) = FC(20) * FC(36) + FC(21) + FC(22) * FC(37)
        FE(3) = FC(23) * FC(36) + FC(24) + FC(25) * FC(37)
        FE(4) = 1 / 4 + FC(37) / 12 + FC(36) / 4 - FE(3) / 5 - 3 * FE(2) / 2 - FE(1)
        FE(5) = FE(1) * (1 / 2 + AA / 6) + FE(2) * (1 / 4 + 11 * AA / 84) + FE(3) * (1 / 70 + AA / 105)
        FE(6) = FE(5) - FC(36) * (7 / 120 + AA / 36 + 3 * AA / FC(0)) - 1 / 40 - AA / 72 - FC(37) *
                    (1 / 60 + AA / 120 + 1 / FC(0))

        If g1 = g0 Then
            FF = 0.90892
            V = 0.550103
            f = 1
        Else
            FF = -FE(6) / ((FC(0) / 2.73) ^ 0.25 * (1 + AA) ^ 3 / FC(0))
            V = FE(4) / ((2.73 / FC(0)) ^ (1 / 4) * (1 + AA) ^ 3)
            f = Max(FC(36) / (1 + AA), 1)
        End If

        e = FF / ho
        D = U * ho * g0 ^ 2 / V
        L = (tf * e + 1) / T + tf ^ 3 / D

        '---------------------------------------------------------------------------------------------------
        ' Stress of flange
        '---------------------------------------------------------------------------------------------------

        Dim Sh, Sr, St, ShSr, ShSt, Vf As Double

        Sh = f * Mmax / (L * g1 ^ 2 * B)
        Sr = ((1.33 * tf * e + 1) * Mmax) / (L * tf ^ 2 * B)
        St = YY * Mmax / (tf ^ 2 * B) - Z * Sr

        ShSr = (Sh + Sr) / 2
        ShSt = (Sh + St) / 2

        Vf = (tf + h) * (A ^ 2 - (B - 2 * Ca) ^ 2) * (π / 4)

        '---------------------------------------------------------------------------------------------------
        ' Stress check
        '---------------------------------------------------------------------------------------------------

        If (Sh <= 1.5 * Fs * Sf) And (Sr <= Fs * Sf) Then
            If (St <= Fs * Sf) And (ShSr <= Fs * Sf) Then
                If (ShSt <= Fs * Sf) And Vf <= aResult(0) Then
                    aResult(0) = Vf
                    aResult(1) = A '//Outisde Diamter of Flange
                    aResult(2) = (B - 2 * Ca) '//Inside Diameter of Flange
                    aResult(3) = C '//Bolt-Circle Diameter
                    aResult(4) = tf '//Nominal thickness of Flange
                    aResult(5) = h '//Length of Hub or Welding Leg of Flange
                    aResult(6) = (g0 + Ca) '//Thickness of Hub at Small End
                    aResult(7) = (g1 + Ca) '//Thickness of Hub at Back of Flange
                    aResult(8) = Dog '//Outside Diamter of Gasket or Contact Face
                    aResult(9) = Di '//Inside Diameter of Gasket(without inner ring)
                    aResult(10) = nb '//Number of Bolt
                    aResult(11) = Db '//Nominal Diameter of Bolt

                    Console.WriteLine("Outisde Diamter of Flange-------------------- = {0}", aResult(1))
                    Console.WriteLine("Inside Diameter of Flange-------------------- = {0}", aResult(2))
                    Console.WriteLine("Bolt-Circle Diameter------------------------- = {0}", aResult(3))
                    Console.WriteLine("Nominal thickness of Flange------------------ = {0}", aResult(4))
                    Console.WriteLine("Length of Hub or Welding Leg of Flange------- = {0}", aResult(5))
                    Console.WriteLine("Thickness of Hub at Small End---------------- = {0}", aResult(6))
                    Console.WriteLine("Thickness of Hub at Back of Flange----------- = {0}", aResult(7))
                    Console.WriteLine("Outside Diamter of Gasket or Contact Face---- = {0}", aResult(8))
                    Console.WriteLine("Inside Diameter of Gasket(without inner ring) = {0}", aResult(9))
                    Console.WriteLine("Number of Bolt------------------------------- = {0}", aResult(10))
                    Console.WriteLine("Nominal Diameter of Bolt--------------------- = {0}", aResult(11))
                    Console.ReadLine()

                End If
            End If
        End If
    End Sub

    Function Min(ByVal x As Double, ByVal y As Double) As Double
        If y <= x Then
            Min = y
        Else
            Min = x
        End If
    End Function


    Function Max(ByVal x As Double, ByVal y As Double) As Double

        If y >= x Then
            Max = y
        Else
            Max = x
        End If
    End Function

    Function Carry(ByVal x As Double) As Integer

        If Int(x) < x Then
            Carry = Int(x) + 1
        Else
            Carry = Int(x)
        End If
    End Function


    Function Carry4(ByVal x As Double) As Integer

        If Int(x) < x Then
            Carry4 = Int(x) * 4 + 4
        Else
            Carry4 = Int(x) * 4
        End If

    End Function


    Function GasketWeight(ByVal B As Double) As Integer

        Dim N As Integer

        Select Case B
            Case Is < 1000
                N = 15
            Case Is < 2000
                N = 20
            Case Is < 3000
                N = 25
            Case Is < 3500
                N = 30
        End Select
        Return N
    End Function


    Public Function GasketRequiredWidth(ByVal GSTP As String, ByVal GSSK As String, ByVal N As Double, ByVal w As Double, ByVal tg As Double) As Double

        Dim bo As Double
        Select Case GSTP
            Case Is = "1a"
                bo = N / 2
            Case Is = "1b"
                bo = N / 2
            Case Is = "1c"
                bo = Min((w + tg) / 2, (w + N) / 4)
            Case Is = "1d"
                bo = Min((w + tg) / 2, (w + N) / 4)
            Case Is = 2
                If GSSK = "I" Then
                    bo = (w + N) / 4
                Else
                    bo = (w + 3 * N) / 8
                End If
            Case Is = 3
                If GSSK = "I" Then
                    bo = N / 4
                Else
                    bo = 3 * N / 8
                End If
            Case Is = 4
                If GSSK = "I" Then
                    bo = 3 * N / 8
                Else
                    bo = 7 * N / 16
                End If
            Case Is = 5
                If GSSK = "I" Then
                    bo = N / 4
                Else
                    bo = 3 * N / 8
                End If
            Case Is = 6
                bo = w / 8
        End Select
        Return bo
    End Function

End Module

