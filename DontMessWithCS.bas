Attribute VB_Name = "DontMessWithCS"


/*Quiz 3.2*/

Function Q321I(A, n, r)

    Q321I = A * ((1 - (1 + r) ^ -n) / r)

End Function


Function Q322F(A, n, r)

    Q322F = A * (((1 + r) ^ n - 1) / r)

End Function


Function Q323A(F, r)

    Q323A = F * r
End Function

/*Quiz 3.3*/

Function Q331A(m, n, P, r)

    Q331A = P * r / (1 - (1 + r) ^ -n)

End Function


Function Q331P(A, m, n, r)

    Q331P = (A * (1 - (1 + r) ^ -n)) / r

End Function


Function Q332A(r, x1, x2, x3)

    Q332A = (x1 + (x2 * (1 + r) ^ -1) + (x3 * (1 + r) ^ -2)) / (1 + (1 + r) ^ -1 + (1 + r) ^ -2)

End Function


Function Q333(initial, n, period, r)

    Q333 = initial * (1 - (1 + r) ^ (-(n - (period - 1)))) / r

End Function

Function F1P0i0n(P, i, n, disp)
    '(F/P,i,n)
    'F = P(1+r)^n

    If disp Then
        F1P0i0n = "F = P(1+i)^n"
    Else
        F1P0i0n = P * (1 + i) ^ n
    End If

End Function
