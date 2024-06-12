Option Explicit
Option Base 1

Sub zscore_comparison_thresholding()
'A pairwise comparison between two CRISPR/Cas9 loss-of-function screens reporting enrichment by z-score by Schneider et al. and Wei et al. was performed
'A z-score cut-off to define a moderate enrichment of a hit was defined as 0.3

'In the following sections counting variables i, j and k were defined

Dim i As Long
Dim j As Long
Dim k As Long
'Furthermore, two matrices containing the information of the CRISPR screens were defined, containing the gene names in the first column and the
'z-scores in the second column
Dim Schneider_et_al(19364, 2) As Variant
Dim Wei_et_al(26172, 2) As Variant
Dim common(30000, 3) As Variant


'In this for-loop the values from the Schneider et al. screen are read into the matrix
For i = 1 To 19364
    Schneider_et_al(i, 1) = Tabelle1.Cells(i + 2, 1)
    Schneider_et_al(i, 2) = Tabelle1.Cells(i + 2, 2)
Next i

'In this for-loop the values from the Wei et al. screen are read into the matrix
For i = 1 To 21672
    Wei_et_al(i, 1) = Tabelle1.Cells(i + 2, 4)
    Wei_et_al(i, 2) = Tabelle1.Cells(i + 2, 5)
Next i

'In this for-loop factors are identified, that occur in both CRISPR screens with a z-score > 0.3
'In addition, the hits will be printed in the Excel sheet
k = 1
For i = 1 To 19364
    For j = 1 To 21672
        If Schneider_et_al(i, 1) = Wei_et_al(j, 1) Then
            If Schneider_et_al(i, 2) > 0.3 And Wei_et_al(j, 2) > 0.3 Then
                common(k, 1) = Schneider_et_al(i, 1)
                common(k, 2) = Schneider_et_al(i, 2)
                common(k, 3) = Wei_et_al(j, 2)
                Tabelle3.Cells(k + 2, 1) = Schneider_et_al(i, 1)
                Tabelle3.Cells(k + 2, 2) = Schneider_et_al(i, 2)
                Tabelle3.Cells(k + 2, 3) = Wei_et_al(j, 2)
                k = k + 1
            End If
        End If
    Next j
Next i

End Sub

Sub LFC_comparison_thresholding()
'A  comparison between four CRISPR/Cas9 loss-of-function screens reporting enrichment by log2fc
'Two CRISPR screens performed in this study (in A549 ACE2 and Calu-3) as well as two published screens were included (by Daniloski et al. and Wang et al.)
'A log2fc cut-off to define a moderate enrichment of a hit was defined as 0.2

'In the following sections counting variables i, j and flag were defined
Dim i As Long
Dim j As Long
Dim flag As Long

'Furthermore, four matrices containing the information of the CRISPR screens were defined, containing the gene names in the first column and the
'z-scores in the second column
Dim This_study_A549_ACE2(18659, 2) As Variant
Dim This_study_Calu3(18659, 2) As Variant
Dim Daniloski_et_al(19049, 2) As Variant
Dim Wang_et_al(20915, 2) As Variant

'Here, 2 matrices (common1 and common2) were defined to store overlapping factors between the two screens from this study (common1) as well as overlapping factors
'of the two previously CRISPR screens (common2)
'In addition, a third matrix (common_total) was defined to store the hits that are enriched in each of the four screens with a log2fc > 0.2
'In each of the matrices the gene name is stored in column 1 followed by the log2fc in the respective screens in the order listed above
Dim common1(30000, 3) As Variant
Dim common2(30000, 3) As Variant
Dim common_total(3000, 5) As Variant

'In this for-loop the values from the CRISPR screen performed in A549 ACE2 during this study are read into the matrix
For i = 1 To 18659
    This_study_A549_ACE2(i, 1) = Tabelle2.Cells(i + 2, 1)
    This_study_A549_ACE2(i, 2) = Tabelle2.Cells(i + 2, 2)
Next i

'In this for-loop the values from the CRISPR screen performed in Calu-3 during this study are read into the matrix
For i = 1 To 18659
    This_study_Calu3(i, 1) = Tabelle2.Cells(i + 2, 4)
    This_study_Calu3(i, 2) = Tabelle2.Cells(i + 2, 5)
Next i

'In this for-loop the values from the Daniloski et al. screen are read into the matrix
For i = 1 To 19049
    Daniloski_et_al(i, 1) = Tabelle2.Cells(i + 2, 7)
    Daniloski_et_al(i, 2) = Tabelle2.Cells(i + 2, 8)
Next i

'In this for-loop the values from the Wang et al. screen are read into the matrix
For i = 1 To 20915
    Wang_et_al(i, 1) = Tabelle2.Cells(i + 2, 10)
    Wang_et_al(i, 2) = Tabelle2.Cells(i + 2, 11)
Next i

'In this loop, factors enriched in both CRISPR screens performed during this study (in A549 ACE2 and Calu-3) are compared
'Factors that appear in both screens with a log2fc > 0.2 are included and stored in the matrix common1
flag = 1
For i = 1 To 18659
    For j = 1 To 18659
        If This_study_A549_ACE2(i, 1) = This_study_Calu3(j, 1) Then
            If This_study_A549_ACE2(i, 2) > 0.2 And This_study_Calu3(j, 2) > 0.2 Then
                common1(flag, 1) = This_study_A549_ACE2(i, 1)
                common1(flag, 2) = This_study_A549_ACE2(i, 2)
                common1(flag, 3) = This_study_Calu3(j, 2)
                flag = flag + 1
            End If
        End If
    Next j
Next i

'In this loop, factors enriched in both CRISPR screens froom the literature (Daniloski et al. and Wang et al.) are compared
'Factors that appear in both screens with a log2fc > 0.2 are included and stored in the matrix common2
flag = 1
For i = 1 To 19049
    For j = 1 To 20915
        If Daniloski_et_al(i, 1) = Wang_et_al(j, 1) Then
            If Daniloski_et_al(i, 2) > 0.2 And Wang_et_al(j, 2) > 0.2 Then
                common2(flag, 1) = Daniloski_et_al(i, 1)
                common2(flag, 2) = Daniloski_et_al(i, 2)
                common2(flag, 3) = Wang_et_al(j, 2)
                flag = flag + 1
            End If
        End If
    Next j
Next i

'In this for-loop all factors are identified that are enriched in all 4 screens with a log2fc> 0.2
'The identified factors are stored in common_total and printed into the excel sheet
flag = 1
For i = 1 To 29999
    For j = 1 To 29999
        If common1(i, 1) = common2(j, 1) And common1(i, 1) <> "" Then
            common_total(flag, 1) = common1(i, 1)
            common_total(flag, 2) = common1(i, 2)
            common_total(flag, 3) = common1(i, 3)
            common_total(flag, 4) = common2(j, 2)
            common_total(flag, 5) = common2(j, 3)
            Tabelle4.Cells(flag + 2, 1) = common1(i, 1)
            Tabelle4.Cells(flag + 2, 2) = common1(i, 2)
            Tabelle4.Cells(flag + 2, 3) = common1(i, 3)
            Tabelle4.Cells(flag + 2, 4) = common2(j, 2)
            Tabelle4.Cells(flag + 2, 5) = common2(j, 3)
            flag = flag + 1
        End If
    Next j
Next i



End Sub

Sub total_common()
'In this sub, all factors that are enriched with a z-score > 0.3 in the CRISPR screens by Schneider et al. and Wei et al. and a log2fc > 0.2
'in the CRISPR screens in this study (A549 ACE2 and Calu-3) and the screens by Daniloski et al. and Wang et al. were identified

'Here counting variables are defined
Dim i As Integer
Dim j As Integer
Dim flag As Integer

'A matrix to store the overlapping factors from the CRISPR screens by Schneider et al. and Wei et al. were defined
Dim common_zscore(2142, 3) As Variant
'A matrix to store the overlapping factors from the CRISPR screens performed in this study and by Daniloski et al. and Wang et al. was defined
Dim common_LFC(200, 5) As Variant
'This matrix was defined to store the overlapping factors from both analyses
Dim common_total(200, 7) As Variant

'The overlapping factors from the CRISPR screens by Schneider et al. and Wei et al. are read into the matrix common_zscore
For i = 1 To 2142
    common_zscore(i, 1) = Tabelle3.Cells(i + 2, 1)
    common_zscore(i, 2) = Tabelle3.Cells(i + 2, 2)
    common_zscore(i, 3) = Tabelle3.Cells(i + 2, 3)
Next i

'The overlapping factors from the CRISPR screens performed in this study and by Daniloski et al. and Wang et al. are read into the matrix common_LFC
For i = 1 To 200
    common_LFC(i, 1) = Tabelle4.Cells(i + 2, 1)
    common_LFC(i, 2) = Tabelle4.Cells(i + 2, 2)
    common_LFC(i, 3) = Tabelle4.Cells(i + 2, 3)
    common_LFC(i, 4) = Tabelle4.Cells(i + 2, 4)
    common_LFC(i, 5) = Tabelle4.Cells(i + 2, 5)
Next i

'In this for loop, overlapping factors are identified and printed into the excel sheet
flag = 1
For i = 1 To 2142
    For j = 1 To 200
        If common_zscore(i, 1) = common_LFC(j, 1) Then
            common_total(flag, 1) = common_zscore(i, 1)
            common_total(flag, 2) = common_zscore(i, 2)
            common_total(flag, 3) = common_zscore(i, 3)
            common_total(flag, 4) = common_LFC(j, 2)
            common_total(flag, 5) = common_LFC(j, 3)
            common_total(flag, 6) = common_LFC(j, 4)
            common_total(flag, 7) = common_LFC(j, 5)
            Tabelle5.Cells(flag + 2, 1) = common_total(flag, 1)
            Tabelle5.Cells(flag + 2, 2) = common_total(flag, 2)
            Tabelle5.Cells(flag + 2, 3) = common_total(flag, 3)
            Tabelle5.Cells(flag + 2, 6) = common_total(flag, 1)
            Tabelle5.Cells(flag + 2, 7) = common_total(flag, 4)
            Tabelle5.Cells(flag + 2, 8) = common_total(flag, 5)
            Tabelle5.Cells(flag + 2, 9) = common_total(flag, 6)
            Tabelle5.Cells(flag + 2, 10) = common_total(flag, 7)
            flag = flag + 1
        End If
    Next j
Next i

End Sub


