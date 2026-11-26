Attribute VB_Name = "modEngine"
Option Explicit

' Roster search engine. Port of python/roster/engine.py. The Python _Search
' object becomes the Private module-level variables below. A step is
' (kind, arg): senior and coag steps carry the group, slot steps the slot.

Private Const ST_SENIOR As Long = 0
Private Const ST_COAG As Long = 1
Private Const ST_SLOT As Long = 2

Private mD As Long
Private mRow() As Long
Private mUsed() As Boolean
Private mElig() As Boolean
Private mOrder() As Long
Private mStepKind() As Long
Private mStepArg() As Long
Private mStepCount As Long
Private mScore0 As Long
Private mScore1 As Long
Private mScore2 As Long
Private mBound0 As Long
Private mBound1 As Long
Private mBound2 As Long
Private mHaveBest As Boolean
Private mBestScore0 As Long
Private mBestScore1 As Long
Private mBestScore2 As Long
Private mBest() As Long
Private mBestUnmetCount As Long
Private mBestUnmetGroup(0 To 3) As Long
Private mBestUnmetCode(0 To 3) As Long
Private mUnmetCount As Long
Private mUnmetGroup(0 To 3) As Long
Private mUnmetCode(0 To 3) As Long
Private mNodes As Long
Private mStop As Boolean
Private mCut As Boolean

' Sizes every result array for a fresh run and clears it.
Public Sub NewResult()
    Dim nd As Long
    Dim nk As Long
    Dim ns As Long
    Dim d As Long
    Dim k As Long

    nd = gDateCount
    nk = gSlotCount
    ns = gStaffCount
    If nd > 0 And nk > 0 Then
        ReDim gGrid(0 To nd - 1, 0 To nk - 1)
        ReDim gApproved(0 To nd - 1, 0 To nk - 1)
        For d = 0 To nd - 1
            For k = 0 To nk - 1
                gGrid(d, k) = EMPTY_CELL
                gApproved(d, k) = False
            Next k
        Next d
    Else
        Erase gGrid
        Erase gApproved
    End If
    If nd > 0 Then
        ReDim gCutShort(0 To nd - 1)
        ReDim gNodes(0 To nd - 1)
        ReDim gUnmetCount(0 To nd - 1)
        ReDim gUnmetGroup(0 To nd - 1, 0 To 3)
        ReDim gUnmetCode(0 To nd - 1, 0 To 3)
    Else
        Erase gCutShort
        Erase gNodes
        Erase gUnmetCount
        Erase gUnmetGroup
        Erase gUnmetCode
    End If
    If ns > 0 Then
        ReDim gActiveCount(0 To ns - 1)
        ReDim gStandbyCount(0 To ns - 1)
        ReDim gEligible(0 To ns - 1)
    Else
        Erase gActiveCount
        Erase gStandbyCount
        Erase gEligible
    End If
    gApprCount = 0
    modRules.ComputeEligible
End Sub

' Kuhn's algorithm. Left node i (0 To nLeft - 1) has adjCount(i) right-side
' nodes, adj(i, 0 To adjCount(i) - 1).
Private Function Augment(ByVal u As Long, ByRef adjCount() As Long, ByRef adj() As Long, ByRef seen() As Boolean, ByRef matchRight() As Long) As Boolean
    Dim t As Long
    Dim r As Long

    Augment = False
    For t = 0 To adjCount(u) - 1
        r = adj(u, t)
        If Not seen(r) Then
            seen(r) = True
            If matchRight(r) < 0 Then
                matchRight(r) = u
                Augment = True
                Exit Function
            End If
            If Augment(matchRight(r), adjCount, adj, seen, matchRight) Then
                matchRight(r) = u
                Augment = True
                Exit Function
            End If
        End If
    Next t
End Function

Public Function MaxMatching(ByVal nLeft As Long, ByRef adjCount() As Long, ByRef adj() As Long) As Long
    Dim nRight As Long
    Dim i As Long
    Dim t As Long
    Dim r As Long
    Dim total As Long
    Dim matchRight() As Long
    Dim seen() As Boolean

    MaxMatching = 0
    If nLeft < 1 Then Exit Function
    nRight = 0
    For i = 0 To nLeft - 1
        For t = 0 To adjCount(i) - 1
            r = adj(i, t)
            If r + 1 > nRight Then nRight = r + 1
        Next t
    Next i
    If nRight < 1 Then Exit Function
    ReDim matchRight(0 To nRight - 1)
    ReDim seen(0 To nRight - 1)
    For r = 0 To nRight - 1
        matchRight(r) = -1
    Next r
    total = 0
    For i = 0 To nLeft - 1
        For r = 0 To nRight - 1
            seen(r) = False
        Next r
        If Augment(i, adjCount, adj, seen, matchRight) Then
            total = total + 1
        End If
    Next i
    MaxMatching = total
End Function

' ---- slot helpers ----

Private Function IsActiveOf(ByVal k As Long, ByVal g As Long) As Boolean
    IsActiveOf = False
    If gSlotGroup(k) = g Then
        If gSlotActive(k) Then IsActiveOf = True
    End If
End Function

Private Function IsStandbyOf(ByVal k As Long, ByVal g As Long) As Boolean
    IsStandbyOf = False
    If gSlotGroup(k) = g Then
        If Not gSlotActive(k) Then IsStandbyOf = True
    End If
End Function

Private Function CountActive(ByVal g As Long) As Long
    Dim k As Long
    Dim n As Long

    n = 0
    For k = 0 To gSlotCount - 1
        If IsActiveOf(k, g) Then n = n + 1
    Next k
    CountActive = n
End Function

Private Function HasSenior(ByVal g As Long) As Boolean
    Dim k As Long

    HasSenior = False
    For k = 0 To gSlotCount - 1
        If IsActiveOf(k, g) Then
            If mRow(k) >= 0 Then
                If gSenior(mRow(k)) Then
                    HasSenior = True
                    Exit Function
                End If
            End If
        End If
    Next k
End Function

Private Function HasCoag(ByVal g As Long) As Boolean
    Dim k As Long

    HasCoag = False
    For k = 0 To gSlotCount - 1
        If IsActiveOf(k, g) Then
            If mRow(k) >= 0 Then
                If gCoag(mRow(k)) Then
                    HasCoag = True
                    Exit Function
                End If
            End If
        End If
    Next k
    For k = 0 To gSlotCount - 1
        If IsStandbyOf(k, g) Then
            If mRow(k) >= 0 Then
                If gApproved(mD, k) Then
                    If gCoag(mRow(k)) Then
                        HasCoag = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next k
End Function

Private Function IsFree(ByVal s As Long, ByVal k As Long) As Boolean
    IsFree = False
    If Not mUsed(s) Then
        If mElig(s, k) Then IsFree = True
    End If
End Function

' Fills adj for the slots in slotList(0 To n - 1): staff in candidate order that are free.
Private Sub FillAdj(ByVal n As Long, ByRef slotList() As Long, ByRef adjCount() As Long, ByRef adj() As Long)
    Dim i As Long
    Dim t As Long
    Dim s As Long
    Dim cnt As Long

    For i = 0 To n - 1
        cnt = 0
        For t = 0 To gStaffCount - 1
            s = mOrder(t)
            If IsFree(s, slotList(i)) Then
                adj(i, cnt) = s
                cnt = cnt + 1
            End If
        Next t
        adjCount(i) = cnt
    Next i
End Sub

Private Function MaxOf(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then
        MaxOf = a
    Else
        MaxOf = b
    End If
End Function

' ---- lower bound ----

Private Sub LowerBound()
    Dim nk As Long
    Dim emptyAll() As Long
    Dim emptyAct() As Long
    Dim adjCount() As Long
    Dim adj() As Long
    Dim nAll As Long
    Dim nAct As Long
    Dim gapsA As Long
    Dim activeA As Long
    Dim reqG As Long
    Dim reqA As Long
    Dim pen As Long
    Dim i As Long
    Dim k As Long
    Dim t As Long
    Dim s As Long
    Dim g As Long
    Dim found As Boolean
    Dim k1 As Long
    Dim ok As Boolean

    nk = gSlotCount
    mBound0 = 0
    mBound1 = 0
    mBound2 = 0
    If nk < 1 Then Exit Sub
    ReDim emptyAll(0 To nk - 1)
    ReDim emptyAct(0 To nk - 1)
    ReDim adjCount(0 To nk - 1)
    ReDim adj(0 To nk - 1, 0 To gStaffCount)

    nAll = 0
    nAct = 0
    For k = 0 To nk - 1
        If mRow(k) = EMPTY_CELL Then
            emptyAll(nAll) = k
            nAll = nAll + 1
            If gSlotActive(k) Then
                emptyAct(nAct) = k
                nAct = nAct + 1
            End If
        End If
    Next k
    FillAdj nAll, emptyAll, adjCount, adj
    gapsA = nAll - MaxMatching(nAll, adjCount, adj)
    FillAdj nAct, emptyAct, adjCount, adj
    activeA = nAct - MaxMatching(nAct, adjCount, adj)

    reqG = 0
    reqA = 0
    pen = 0
    For i = 0 To mStepCount - 1
        If mStepKind(i) = ST_SENIOR Then
            g = mStepArg(i)
            If Not HasSenior(g) Then
                found = False
                For t = 0 To gStaffCount - 1
                    s = mOrder(t)
                    If gSenior(s) Then
                        For k = 0 To nk - 1
                            If IsActiveOf(k, g) Then
                                If mRow(k) = EMPTY_CELL Then
                                    If IsFree(s, k) Then found = True
                                End If
                            End If
                        Next k
                    End If
                Next t
                If Not found Then
                    reqG = reqG + 1
                    reqA = reqA + 1
                    pen = pen + 2
                End If
            End If
        ElseIf mStepKind(i) = ST_COAG Then
            g = mStepArg(i)
            If Not HasCoag(g) Then
                found = False
                For t = 0 To gStaffCount - 1
                    s = mOrder(t)
                    If gCoag(s) Then
                        For k = 0 To nk - 1
                            If IsActiveOf(k, g) Then
                                If mRow(k) = EMPTY_CELL Then
                                    If IsFree(s, k) Then found = True
                                End If
                            End If
                        Next k
                    End If
                Next t
                If Not found Then
                    k1 = -1
                    For k = 0 To nk - 1
                        If IsStandbyOf(k, g) Then
                            If mRow(k) = EMPTY_CELL Then
                                k1 = k
                                Exit For
                            End If
                        End If
                    Next k
                    ok = False
                    If k1 >= 0 Then
                        For t = 0 To gStaffCount - 1
                            s = mOrder(t)
                            If gCoag(s) Then
                                If IsFree(s, k1) Then ok = True
                            End If
                        Next t
                    End If
                    If ok Then
                        pen = pen + 1
                    Else
                        pen = pen + 2
                        reqG = reqG + 1
                        If k1 < 0 Then reqA = reqA + 1
                    End If
                End If
            End If
        End If
    Next i
    mBound0 = MaxOf(gapsA, reqG)
    mBound1 = MaxOf(activeA, reqA)
    mBound2 = pen
End Sub

' ---- search ----

' Python tuple comparison: is (n0, n1, n2) >= (b0, b1, b2).
Private Function ScoreGE(ByVal n0 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal b0 As Long, ByVal b1 As Long, ByVal b2 As Long) As Boolean
    If n0 <> b0 Then
        ScoreGE = (n0 > b0)
        Exit Function
    End If
    If n1 <> b1 Then
        ScoreGE = (n1 > b1)
        Exit Function
    End If
    ScoreGE = (n2 >= b2)
End Function

Private Sub Rec(ByVal i As Long)
    Dim x As Long
    Dim g As Long
    Dim k As Long
    Dim kk As Long
    Dim t As Long
    Dim s As Long
    Dim nk As Long

    nk = gSlotCount
    If i = mStepCount Then
        mHaveBest = True
        For k = 0 To nk - 1
            mBest(k) = mRow(k)
        Next k
        mBestUnmetCount = mUnmetCount
        For k = 0 To mUnmetCount - 1
            mBestUnmetGroup(k) = mUnmetGroup(k)
            mBestUnmetCode(k) = mUnmetCode(k)
        Next k
        mBestScore0 = mScore0
        mBestScore1 = mScore1
        mBestScore2 = mScore2
        If mBestScore0 = mBound0 And mBestScore1 = mBound1 And mBestScore2 = mBound2 Then
            mStop = True
        End If
        Exit Sub
    End If
    x = mStepArg(i)
    If mStepKind(i) = ST_SLOT Then
        If mRow(x) <> EMPTY_CELL Then
            Rec i + 1
            Exit Sub
        End If
        For t = 0 To gStaffCount - 1
            s = mOrder(t)
            If Not mUsed(s) Then
                If mElig(s, x) Then
                    TryPlace i, s, x, 0, 0, 0
                    If mStop Then Exit Sub
                End If
            End If
        Next t
        TryGap i, x, GAP, -1
        Exit Sub
    End If
    If mStepKind(i) = ST_SENIOR Then
        g = x
        If HasSenior(g) Then
            Rec i + 1
            Exit Sub
        End If
        For t = 0 To gStaffCount - 1
            s = mOrder(t)
            If gSenior(s) And Not mUsed(s) Then
                For k = 0 To nk - 1
                    If IsActiveOf(k, g) Then
                        If mRow(k) = EMPTY_CELL Then
                            If mElig(s, k) Then
                                TryPlace i, s, k, 0, 0, 0
                                If mStop Then Exit Sub
                            End If
                        End If
                    End If
                Next k
            End If
        Next t
        k = -1
        For kk = 0 To nk - 1
            If IsActiveOf(kk, g) Then
                If mRow(kk) = EMPTY_CELL Then k = kk
            End If
        Next kk
        TryGap i, k, GAP_SENIOR, g
        Exit Sub
    End If
    ' kind = ST_COAG
    g = x
    If HasCoag(g) Then
        Rec i + 1
        Exit Sub
    End If
    ' (a) active
    For t = 0 To gStaffCount - 1
        s = mOrder(t)
        If gCoag(s) And Not mUsed(s) Then
            For k = 0 To nk - 1
                If IsActiveOf(k, g) Then
                    If mRow(k) = EMPTY_CELL Then
                        If mElig(s, k) Then
                            TryPlace i, s, k, 0, 0, 0
                            If mStop Then Exit Sub
                        End If
                    End If
                End If
            Next k
        End If
    Next t
    k = -1
    For kk = 0 To nk - 1
        If IsStandbyOf(kk, g) Then
            If mRow(kk) = EMPTY_CELL Then
                k = kk
                Exit For
            End If
        End If
    Next kk
    ' (b) fallback
    If k >= 0 Then
        For t = 0 To gStaffCount - 1
            s = mOrder(t)
            If gCoag(s) And Not mUsed(s) Then
                If mElig(s, k) Then
                    TryPlace i, s, k, 0, 0, 1
                    If mStop Then Exit Sub
                End If
            End If
        Next t
    End If
    ' (c) unmet
    If k < 0 Then
        For kk = 0 To nk - 1
            If IsActiveOf(kk, g) Then
                If mRow(kk) = EMPTY_CELL Then k = kk
            End If
        Next kk
    End If
    TryGap i, k, GAP_COAG, g
End Sub

Private Sub TryPlace(ByVal i As Long, ByVal s As Long, ByVal k As Long, ByVal dl0 As Long, ByVal dl1 As Long, ByVal dl2 As Long)
    Dim n0 As Long
    Dim n1 As Long
    Dim n2 As Long
    Dim o0 As Long
    Dim o1 As Long
    Dim o2 As Long

    n0 = mScore0 + dl0
    n1 = mScore1 + dl1
    n2 = mScore2 + dl2
    If mHaveBest Then
        If ScoreGE(n0, n1, n2, mBestScore0, mBestScore1, mBestScore2) Then Exit Sub
    End If
    mNodes = mNodes + 1
    If mNodes > NODE_CAP Then
        mStop = True
        mCut = True
        Exit Sub
    End If
    mRow(k) = s
    mUsed(s) = True
    o0 = mScore0
    o1 = mScore1
    o2 = mScore2
    mScore0 = n0
    mScore1 = n1
    mScore2 = n2
    Rec i + 1
    mRow(k) = EMPTY_CELL
    mUsed(s) = False
    mScore0 = o0
    mScore1 = o1
    mScore2 = o2
End Sub

' k = -1 means no slot (an unmet requirement); g is the group then, else -1.
Private Sub TryGap(ByVal i As Long, ByVal k As Long, ByVal code As Long, ByVal g As Long)
    Dim pen As Long
    Dim dl1 As Long
    Dim n0 As Long
    Dim n1 As Long
    Dim n2 As Long
    Dim o0 As Long
    Dim o1 As Long
    Dim o2 As Long

    If code <> GAP Then
        pen = 2
    Else
        pen = 0
    End If
    If k < 0 Then
        dl1 = 1
    ElseIf gSlotActive(k) Then
        dl1 = 1
    Else
        dl1 = 0
    End If
    n0 = mScore0 + 1
    n1 = mScore1 + dl1
    n2 = mScore2 + pen
    If mHaveBest Then
        If ScoreGE(n0, n1, n2, mBestScore0, mBestScore1, mBestScore2) Then Exit Sub
    End If
    mNodes = mNodes + 1
    If mNodes > NODE_CAP Then
        mStop = True
        mCut = True
        Exit Sub
    End If
    If k < 0 Then
        mUnmetGroup(mUnmetCount) = g
        mUnmetCode(mUnmetCount) = code
        mUnmetCount = mUnmetCount + 1
    Else
        mRow(k) = code
    End If
    o0 = mScore0
    o1 = mScore1
    o2 = mScore2
    mScore0 = n0
    mScore1 = n1
    mScore2 = n2
    Rec i + 1
    If k < 0 Then
        mUnmetCount = mUnmetCount - 1
    Else
        mRow(k) = EMPTY_CELL
    End If
    mScore0 = o0
    mScore1 = o1
    mScore2 = o2
End Sub

' Solves date index d: fills gGrid(d, *), the unmet list, cut flag and node count.
Public Sub SolveDate(ByVal d As Long)
    Dim nk As Long
    Dim ns As Long
    Dim k As Long
    Dim s As Long
    Dim g As Long
    Dim u As Long

    nk = gSlotCount
    ns = gStaffCount
    mD = d
    If nk > 0 Then
        ReDim mRow(0 To nk - 1)
        ReDim mBest(0 To nk - 1)
        For k = 0 To nk - 1
            mRow(k) = gGrid(d, k)
            mBest(k) = EMPTY_CELL
        Next k
    End If
    modRules.CandidateOrder mOrder
    If ns > 0 Then
        ReDim mUsed(0 To ns - 1)
        For s = 0 To ns - 1
            mUsed(s) = False
        Next s
        If nk > 0 Then
            For k = 0 To nk - 1
                If mRow(k) >= 0 Then mUsed(mRow(k)) = True
            Next k
            ReDim mElig(0 To ns - 1, 0 To nk - 1)
            For s = 0 To ns - 1
                For k = 0 To nk - 1
                    If Len(modRules.WithoutRule(modRules.BlockingRules(s, d, k), 3)) = 0 Then
                        mElig(s, k) = Not modRules.ForwardNightBlock(s, d, k)
                    Else
                        mElig(s, k) = False
                    End If
                Next k
            Next s
        End If
    End If

    ' steps: at most 4 senior/coag steps plus one per slot
    ReDim mStepKind(0 To nk + 4)
    ReDim mStepArg(0 To nk + 4)
    mStepCount = 0
    For g = 0 To 1
        If CountActive(g) > 0 Then
            If gRuleOn(1) Then
                mStepKind(mStepCount) = ST_SENIOR
                mStepArg(mStepCount) = g
                mStepCount = mStepCount + 1
            End If
            If gRuleOn(2) Then
                mStepKind(mStepCount) = ST_COAG
                mStepArg(mStepCount) = g
                mStepCount = mStepCount + 1
            End If
        End If
    Next g
    For k = 0 To nk - 1
        If mRow(k) = EMPTY_CELL Then
            mStepKind(mStepCount) = ST_SLOT
            mStepArg(mStepCount) = k
            mStepCount = mStepCount + 1
        End If
    Next k

    LowerBound
    mScore0 = 0
    mScore1 = 0
    mScore2 = 0
    mHaveBest = False
    mBestScore0 = 0
    mBestScore1 = 0
    mBestScore2 = 0
    mBestUnmetCount = 0
    mNodes = 0
    mStop = False
    mCut = False
    mUnmetCount = 0

    Rec 0

    If mHaveBest Then
        For k = 0 To nk - 1
            gGrid(d, k) = mBest(k)
        Next k
    End If
    gUnmetCount(d) = mBestUnmetCount
    For u = 0 To mBestUnmetCount - 1
        gUnmetGroup(d, u) = mBestUnmetGroup(u)
        gUnmetCode(d, u) = mBestUnmetCode(u)
    Next u
    gCutShort(d) = mCut
    gNodes(d) = mNodes
End Sub

Private Sub CountDate(ByVal d As Long)
    Dim k As Long
    Dim v As Long

    For k = 0 To gSlotCount - 1
        v = gGrid(d, k)
        If v >= 0 Then
            If gSlotActive(k) Then
                gActiveCount(v) = gActiveCount(v) + 1
            Else
                gStandbyCount(v) = gStandbyCount(v) + 1
            End If
        End If
    Next k
End Sub

' Re-solves every date from index fromD; earlier dates are kept and counted.
Public Sub SolveYear(ByVal fromD As Long)
    Dim d As Long
    Dim k As Long
    Dim s As Long

    For d = fromD To gDateCount - 1
        For k = 0 To gSlotCount - 1
            If Not gApproved(d, k) Then gGrid(d, k) = EMPTY_CELL
        Next k
        gCutShort(d) = False
        gUnmetCount(d) = 0
        gNodes(d) = 0
    Next d
    For s = 0 To gStaffCount - 1
        gActiveCount(s) = 0
        gStandbyCount(s) = 0
    Next s
    For d = 0 To fromD - 1
        CountDate d
    Next d
    For d = fromD To gDateCount - 1
        SolveDate d
        CountDate d
    Next d
End Sub
