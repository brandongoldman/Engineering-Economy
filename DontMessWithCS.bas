
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

'--- --- --- --- --- --- --- --- --- --- --- ---
'Single Payment
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

'--- --- --- --- --- --- --- --- --- --- --- ---
'Uniform Series
'--- --- --- --- --- --- --- --- --- --- --- ---
Function F1A0i0n(A, i, n, disp)
'(F/A, i, n)
'F = A * [(1 + i) ^ n - 1) / i]

  If disp Then
    F1A0i0n = "F = A * [(1 + i) ^ n - 1) / i]"
  Else
    F1A0i0n = A * [(1 + i) ^ n - 1) / i]
  End If

End Function
'--- --- --- --- --- --- --- --- --- --- --- ---
Function A1F0i0n(F, i, n, disp)
'(A/F, i, n)
'A = F * [i / ((1 + i)^n - 1)]

  If disp Then
    A1F0i0n = "A = F * [i / ((1 + i)^n - 1)]"
  Else
    A1F0i0n = F * [i / ((1 + i)^n - 1)]
  End If

End Function
'--- --- --- --- --- --- --- --- --- --- --- ---
Function A1P0i0n(P, i, n, disp)
'(A/P, i, n)
'A = P * [(i * (1 + i)^n) / ((1 + i)^n - 1)]

  If disp Then
    A1P0i0n = "A = P * [(i * (1 + i)^n) / ((1 + i)^n - 1)]"
  Else
    A1P0i0n = P * [(i * (1 + i)^n) / ((1 + i)^n - 1)]
  End If

End Function
'--- --- --- --- --- --- --- --- --- --- --- ---
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
'Continuous Compounding at Nominal Rate r : Single Payment
'--- --- --- --- --- --- --- --- --- --- --- ---
Function SinglePaymentF(P, e, r, n, disp)
'F = P* e^(r*n)

  If disp Then
    SinglePaymentF = "F = P* e^(r*n)"
  Else
    SinglePaymentF = P * exp(r * n)
  End If

End Function
'--- --- --- --- --- --- --- --- --- --- --- ---
Function SinglePaymentP(F, e, r, n, disp)
'P = F*e^(-r*n)

  If disp Then
    SinglePaymentP = "P = F * e^(-r*n)"
  Else
    SinglePaymentP = F * exp(-r*n)
  End If

End Function

'--- --- --- --- --- --- --- --- --- --- --- ---
'Continuous Compounding at Nominal Rate r : Uniform Series
'--- --- --- --- --- --- --- --- --- --- --- ---
Function UniformCompooundA1(F, r, n, disp)
'A = F * [(e^r - 1) / (e^r*n - 1)]

  If disp Then
    UniformCompooundA1 = "F * [(e^r - 1) / (e^r*n - 1)]"
  Else
    UniformCompooundA1 = F * [(exp(r) - 1) / (exp(r*n) - 1)]
  End If

End Function
'--- --- --- --- --- --- --- --- --- --- --- ---
Function UniformCompoundA2(P, r, n, disp)
'A = P * [(e^(r*n) * (e^(r) - 1)) / (e^(r*n) - 1)]

  If disp Then
    UniformCompoundA2 = "A = P * [(e^(r*n) * (e^(r) - 1)) / (e^(r*n) - 1)]"
  Else
    UniformCompoundA2 = P * [(exp(r*n) * (exp(r) - 1)) / (exp(r*n) - 1)]
  End If

End Function
'--- --- --- --- --- --- --- --- --- --- --- ---
Function UniformCompoundF(A, r, n, disp)
'F = A * [(e^(r*n) - 1) /  (e^(r) - 1)]

  If disp Then
    UniformCompoundF = "F = A * [(e^(r*n) - 1) /  (e^(r) - 1)]"
  Else
    UniformCompoundF =A * [(exp(r*n) - 1) /  (exp(r) - 1)]
  End If

End Function
'--- --- --- --- --- --- --- --- --- --- --- ---
Function UniformCompoundP(A, r, n, disp)
'P = A * [(e^(r*n) - 1) / (e^(r*n) * (e^(r) - 1))]

  If disp Then
    UniformCompoundP = "P = A * [(e^(r*n) - 1) / (e^(r*n) * (e^(r) - 1))]"
  Else
    UniformCompoundP =  A * [(exp(r*n) - 1) / (exp(r*n) * (exp(r) - 1))]
  End If

End Function

'--- --- --- --- --- --- --- --- --- --- --- ---
' Continuous Uniform Cash Flow with Continuous Compounding : Present Worth
'--- --- --- --- --- --- --- --- --- --- --- ---
Function P1F0r0n(F, r, n, disp)
'(P/F, r, n)
'P = F * [(e^(r)-1) / (r * e^(r*n))]
  If disp Then
    P1F0r0n = "P = F * [(e^(r)-1) / (r * e^(r*n))]"
  Else
    P1F0r0n = F * [(exp(r)-1) / (r * exp(r*n))]
  End If

End Function
'--- --- --- --- --- --- --- --- --- --- --- ---
' Continuous Uniform Cash Flow with Continuous Compounding : Compound Amount
'--- --- --- --- --- --- --- --- --- --- --- ---
Function F1P0r0n(P, r, n, disp)
'(F/P, r, n)
'F = P * [( (e^(r) - 1) * (e^(r*n)) ) / (r*e^(r))]
  If disp Then
    F1P0r0n = "F = P * [( (e^(r) - 1) * (e^(r*n)) ) / (r*e^(r))]"
  Else
    F1P0r0n = P * [( (exp(r) - 1) * (exp(r*n)) ) / (r*exp(r))]
  End If

End Function