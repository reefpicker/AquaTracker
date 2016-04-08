VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMersenneTwister64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Visual Basic Mersenne-Twister
' Author: Carmine Arturo Sangiovanni
'         carmine @ daygo.com.br
'
'         Aug 13,2004
'
'         based on C++ code
'

'Example Usage:
'        'randomize seed itself using timer
'        Mersenne_Twister.Randomize Timer
'        r = Mersenne_Twister.Generate_Number(NumberOfGroups) + 1

Option Explicit

Const N = 624
Const m = 397

Dim mt(0 To N) As Currency
Dim mti As Currency

Dim MATRIX_A As Currency
Dim UPPER_MASK As Currency
Dim LOWER_MASK As Currency
Dim FULL_MASK As Currency
Dim TEMPERING_MASK_B As Currency
Dim TEMPERING_MASK_C As Currency

Private Function tempering_shift_u(ty As Currency)
    tempering_shift_u = f_and(Int(ty / 2048@), FULL_MASK)
End Function

Private Function tempering_shift_s(ty As Currency)
    tempering_shift_s = and_ffffffff(ty * 128@)
End Function

Private Function tempering_shift_t(ty As Currency)
    tempering_shift_t = and_ffffffff(ty * 32768@)
End Function

Private Function tempering_shift_l(ty As Currency)
    tempering_shift_l = f_and(Int(ty / 262144@), FULL_MASK)
End Function

Private Function f_and(p1 As Currency, p2 As Currency)
    Dim v As Currency
    Dim i As Integer
    
    If (p1 < UPPER_MASK) And (p2 < UPPER_MASK) Then
        f_and = p1 And p2
    End If
    
    If (p1 < UPPER_MASK) And (p2 >= UPPER_MASK) Then
        f_and = p1 And (p2 - UPPER_MASK)
    End If
    
    If (p1 >= UPPER_MASK) And (p2 < UPPER_MASK) Then
        f_and = (p1 - UPPER_MASK) And p2
    End If
    
    If (p1 >= UPPER_MASK) And (p2 >= UPPER_MASK) Then
        f_and = (p1 - UPPER_MASK) And (p2 - UPPER_MASK)
        f_and = f_and + UPPER_MASK
    End If
End Function

Private Function f_or(p1 As Currency, p2 As Currency)
    Dim v As Currency
    Dim i As Integer
    Dim f As Boolean
    
    If (p1 < UPPER_MASK) And (p2 < UPPER_MASK) Then
        f_or = p1 Or p2
    End If
    If (p1 < UPPER_MASK) And (p2 >= UPPER_MASK) Then
        f_or = p1 Or (p2 - UPPER_MASK)
        f_or = f_or + UPPER_MASK
    End If
    If (p1 >= UPPER_MASK) And (p2 < UPPER_MASK) Then
        f_or = (p1 - UPPER_MASK) And p2
        f_or = f_or + UPPER_MASK
    End If
    If (p1 >= UPPER_MASK) And (p2 >= UPPER_MASK) Then
        f_or = (p1 - UPPER_MASK) And (p2 - UPPER_MASK)
        f_or = f_or + UPPER_MASK
    End If
End Function

Private Function f_xor(p1 As Currency, p2 As Currency)
    Dim v As Currency
    Dim i As Integer
    Dim f1 As Boolean, f2 As Boolean
    
    If (p1 < UPPER_MASK) And (p2 < UPPER_MASK) Then
        f_xor = p1 Xor p2
    End If
    If (p1 < UPPER_MASK) And (p2 >= UPPER_MASK) Then
        f_xor = p1 Xor (p2 - UPPER_MASK)
        f_xor = f_xor + UPPER_MASK
    End If
    If (p1 >= UPPER_MASK) And (p2 < UPPER_MASK) Then
        f_xor = (p1 - UPPER_MASK) Xor p2
        f_xor = f_xor + UPPER_MASK
    End If
    If (p1 >= UPPER_MASK) And (p2 >= UPPER_MASK) Then
        f_xor = (p1 - UPPER_MASK) Xor (p2 - UPPER_MASK)
    End If
End Function

Private Function f_lower(p1 As Currency)
    Do
        If p1 < UPPER_MASK Then
            f_lower = p1
            Exit Do
        Else
            p1 = p1 - UPPER_MASK
        End If
    Loop
End Function

Private Function f_upper(p1 As Currency)
    If p1 > LOWER_MASK Then
        f_upper = UPPER_MASK
    Else
        f_upper = 0
    End If
End Function

Private Function f_xor3(p1 As Currency, p2 As Currency, p3 As Currency)
    Dim v As Currency
    Dim tmp As Currency
    Dim i As Integer
    Dim f As Integer
    
    
    If (p1 < UPPER_MASK) And (p2 < UPPER_MASK) Then
        tmp = p1 Xor p2
    End If
    If (p1 < UPPER_MASK) And (p2 >= UPPER_MASK) Then
        tmp = p1 Xor (p2 - UPPER_MASK)
        tmp = tmp + UPPER_MASK
    End If
    If (p1 >= UPPER_MASK) And (p2 < UPPER_MASK) Then
        tmp = (p1 - UPPER_MASK) Xor p2
        tmp = tmp + UPPER_MASK
    End If
    If (p1 >= UPPER_MASK) And (p2 >= UPPER_MASK) Then
        tmp = (p1 - UPPER_MASK) Xor (p2 - UPPER_MASK)
    End If
    
    If (tmp < UPPER_MASK) And (p3 < UPPER_MASK) Then
        f_xor3 = tmp Xor p3
    End If
    If (tmp < UPPER_MASK) And (p3 >= UPPER_MASK) Then
        f_xor3 = tmp Xor (p3 - UPPER_MASK)
        f_xor3 = f_xor3 + UPPER_MASK
    End If
    If (tmp >= UPPER_MASK) And (p3 < UPPER_MASK) Then
        f_xor3 = (tmp - UPPER_MASK) Xor p3
        f_xor3 = f_xor3 + UPPER_MASK
    End If
    If (tmp >= UPPER_MASK) And (p3 >= UPPER_MASK) Then
        f_xor3 = (tmp - UPPER_MASK) Xor (p3 - UPPER_MASK)
    End If
End Function

Private Function and_ffffffff(c As Currency)
    Dim e As Currency
    Dim i As Integer
    
    i = 32
    Do
        e = 2 ^ (i + 16)
        Do While c >= e
            c = c - e
        Loop
        i = i - 1
    Loop While i > 15
    and_ffffffff = c
End Function

Public Sub Randomize(seed As Currency)
    mt(0) = and_ffffffff(seed)
    For mti = 1 To N - 1
        mt(mti) = and_ffffffff(69069 * mt(mti - 1))
    Next mti
End Sub


Public Function Generate_Number(max As Integer) As Currency

    Dim kk As Integer
    
    Dim ty1 As Currency
    Dim ty2 As Currency
    Dim y As Currency
    
    Dim mag01(0 To 1) As Currency
    
    MATRIX_A = 2567483615@              '&H9908b0df
    UPPER_MASK = 2147483648@            '&H80000000
    LOWER_MASK = 2147483647@            '&H7fffffff
    FULL_MASK = LOWER_MASK + UPPER_MASK '&Hffffffff
    TEMPERING_MASK_B = 2636928640@      '&H9d2c5680
    TEMPERING_MASK_C = 4022730752@      '&Hefc60000
    
    mag01(0) = 0@
    mag01(1) = MATRIX_A
    
    If mti >= N Then
        If mti = N + 1 Then
            Randomize 4537
        End If
        
        For kk = 0 To (N - m) - 1
            y = f_or(f_upper(mt(kk)), f_lower(mt(kk + 1)))
            mt(kk) = f_xor3(mt(kk + m), Int(y / 2@), mag01(f_and(y, 1)))
        Next kk
        
        For kk = kk To (N - 1) - 1
            y = f_or(f_upper(mt(kk)), f_lower(mt(kk + 1)))
            mt(kk) = f_xor3(mt(kk + (m - N)), Int(y / 2@), mag01(f_and(y, 1)))
        Next kk
        
        y = f_or(f_upper(mt(N - 1)), f_lower(mt(0)))
        mt(N - 1) = f_xor3(mt(m - 1), Int(y / 2@), mag01(f_and(y, 1)))
        mti = 0
    End If
    
    '---------------------------------------------------
    y = mt(mti): mti = mti + 1
    
    '---------------------------------------------------
    y = f_xor(y, tempering_shift_u(y))
    
    ty1 = f_and(tempering_shift_s(y), TEMPERING_MASK_B)
    y = f_xor(y, ty1)
    
    ty1 = f_and(tempering_shift_t(y), TEMPERING_MASK_C)
    y = f_xor(y, ty1)
    
    y = f_xor(y, tempering_shift_l(y))
    
    '---------------------------------------------------
    If max = 0 Then
        Generate_Number = 0
    Else
        Generate_Number = Int(y / 32) Mod max
    End If
End Function





