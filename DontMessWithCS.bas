
'--- --- --- --- --- --- --- --- --- --- --- ---
'Quiz 3.2
'--- --- --- --- --- --- --- --- --- --- --- ---
Function Q321I(A, n, r, disp)
'F = A*(1-(1+r)^-n)/r

    If disp Then
      Q321I = "F = A*(1-(1+r)^-n)/r "
    Else
      Q321I = A * ((1 - (1 + r) ^ -n) / r)
    End If

End Function
'--- --- --- --- --- --- --- --- --- --- --- ---
Function Q322F(A, n, r, disp)
'F = A*(1-(1+r)^-n)/r

    If disp Then
      Q322F = "F = A*(1-(1+r)^-n)/r "
    Else
      Q322F = A * (((1 + r) ^ n - 1) / r)
    End If

End Function
'--- --- --- --- --- --- --- --- --- --- --- ---
Function Q323A(F, r, disp)
'F = A*(1-(1+r)^-n)/r
    If disp Then
      Q323A = "F = A*(1-(1+r)^-n)/r , n = inf"
    Else
      Q323A = F * r
    End If

End Function

'--- --- --- --- --- --- --- --- --- --- --- ---
'Quiz 3.3
'--- --- --- --- --- --- --- --- --- --- --- ---
Function Q331A(m, n, P, r, disp)
'A = P * r / (1 - (1+r)^-n)

    If disp Then
      Q331A = "A = P * r / (1 - (1+r)^-n)"
    Else
      Q331A = P * r / (1 - (1 + r) ^ -n)
    End If

End Function
'--- --- --- --- --- --- --- --- --- --- --- ---
Function Q331P(A, m, n, r, disp)
'P = (1+r)^-m * A * ((1-(1+r)^-n)/r)

    If disp Then
      Q331P = "P = (1+r)^-m * A * ((1-(1+r)^-n)/r)"
    Else
      Q331P = (A * (1 - (1 + r) ^ -n)) / r
    End If

End Function

'--- --- --- --- --- --- --- --- --- --- --- ---
Function Q332A(r, x1, x2, x3, disp)
'P1 = x1 + x2(1+r)^-1 + x3(1+r)^-2
'P2 = A(1 + (1+r)^-1 + (1+r)^-2)

    If disp Then
      Q332A = "P1 = x1 + x2(1+r)^-1 + x3(1+r)^-2 and P2 = A(1 + (1+r)^-1 + (1+r)^-2)"
    Else
      Q332A = (x1 + (x2 * (1 + r) ^ -1) + (x3 * (1 + r) ^ -2)) / (1 + (1 + r) ^ -1 + (1 + r) ^ -2)
    End If

End Function
'--- --- --- --- --- --- --- --- --- --- --- ---
Function Q333(initial, n, period, r, disp)
  'A(1-(1+r)^(-n+k-1))/r

    If disp Then
      Q333 = "Portion = A(1+r)^(-n+k-1)"
    Else
      Q333 = initial * (1 - (1 + r) ^ (-(n - (period - 1)))) / r
    End If

End Function
'--- --- --- --- --- --- --- --- --- --- --- ---
'Michelle's Functions that conform to Cory's dictator-like naming constraints
'--- --- --- --- --- --- --- --- --- --- --- ---

Function P1F0i0n(F, i, n, disp)
'(P/F, i, n)
'P = F*(1+i)^-n

  If disp Then
    P1F0i0n = "P = F*(1+i)^-n"
  Else
    P1F0i0n = F*(1+i)^-n
  End If

End Function

Function F1A0i0n(A, i, n, disp)
'(F/A, i, n)
'F = A * [(1 + i) ^ n - 1) / i]

  If disp Then
    F1A0i0n = "F = A * [(1 + i) ^ n - 1) / i]"
  Else
    F1A0i0n = A * [(1 + i) ^ n - 1) / i]
  End If

End Function

Function A1F0i0n(F, i, n, disp)
'(A/F, i, n)
'A = F * [i / ((1 + i)^n - 1)]

  If disp Then
    A1F0i0n = "A = F * [i / ((1 + i)^n - 1)]"
  Else
    A1F0i0n = F * [i / ((1 + i)^n - 1)]
  End If

End Function

Function A1P0i0n(P, i, n, disp)
'(A/P, i, n)
'A = P * [(i * (1 + i)^n) / ((1 + i)^n - 1)]

  If disp Then
    A1P0i0n = "A = P * [(i * (1 + i)^n) / ((1 + i)^n - 1)]"
  Else
    A1P0i0n = P * [(i * (1 + i)^n) / ((1 + i)^n - 1)]
  End If

End Function

Function P1A0i0n(A, i, n, disp)
'(P/A, i, n)
'P = A * [((1 + i)^n - 1) / (i * (1 + i)^n)]

  If disp Then
    P1A0i0n = "P = A * [((1 + i)^n - 1) / (i * (1 + i)^n)]"
  Else
    P1A0i0n = A * [((1 + i)^n - 1) / (i * (1 + i)^n)]
  End If

End Function

'--- --- --- --- --- --- --- --- --- --- --- ---
'Cory's Functions
'--- --- --- --- --- --- --- --- --- --- --- ---
Function F1P0i0n(P, i, n, disp)

    '(F/P,i,n)
    'F = P(1+r)^n

    If disp Then
        F1P0i0n = "F = P(1+i)^n"
    Else
        F1P0i0n = P * (1 + i) ^ n
    End If

End Function
