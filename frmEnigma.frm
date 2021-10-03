VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmEnigma 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Enigma Machine"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNumWheels 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   6180
      Width           =   375
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   495
      Left            =   180
      TabIndex        =   9
      Top             =   6180
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   873
      _Version        =   327681
      Value           =   1
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtOutputBox 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   4620
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   540
      Width           =   4215
   End
   Begin ComctlLib.ProgressBar EnigmaProgress 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   6960
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   556
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Height          =   195
      Left            =   4800
      TabIndex        =   3
      Top             =   6720
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "Decrypt"
      Height          =   315
      Left            =   7620
      TabIndex        =   2
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "Encrypt"
      Height          =   315
      Left            =   7620
      TabIndex        =   1
      Top             =   6420
      Width           =   1215
   End
   Begin VB.TextBox txtEnigma 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   540
      Width           =   4275
   End
   Begin VB.Label Label4 
      Caption         =   "Number of encryption wheels"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1020
      TabIndex        =   11
      Top             =   6300
      Width           =   3015
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Encrypted Output"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4680
      TabIndex        =   8
      Top             =   180
      Width           =   1635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Normal, Unencrypted text"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   7
      Top             =   180
      Width           =   2430
   End
   Begin VB.Label Label1 
      Height          =   195
      Left            =   5400
      TabIndex        =   4
      Top             =   6660
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "frmEnigma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim i As Long
Dim j As Long
Dim k As Long

Dim WheelIndex1 As Long
Dim WheelIndex2 As Long
Dim WheelIndex3 As Long
Dim WheelIndex4 As Long
Dim WheelIndex5 As Long
Dim WheelIndex6 As Long
Dim WheelIndex7 As Long
Dim WheelIndex8 As Long
Dim WheelIndex9 As Long
Dim WheelIndex10 As Long

Dim iTotalChar As Long
Dim lProcessCount As Long
Dim lNumTurns As Long
Dim iCurrentWheel As Integer
Dim TotalWheels As Integer

Dim Wheel2String As String
Dim Wheel3String As String
Dim Wheel4String As String
Dim Wheel5String As String
Dim Wheel6String As String
Dim Wheel7String As String
Dim Wheel8String As String
Dim Wheel9String As String
Dim Wheel10String As String
Dim Wheel11String As String
Dim Wheel12String As String
Dim Wheel13String As String
Dim Wheel14String As String
Dim Wheel15String As String
Dim Wheel16String As String
Dim Wheel17String As String
Dim Wheel18String As String
Dim Wheel19String As String
Dim Wheel20String As String

Dim strCurrentChar As String

Private Wheel1() As Variant
Private Wheel2() As Variant
Private Wheel3() As Variant
Private Wheel4() As Variant
Private Wheel5() As Variant
Private Wheel6() As Variant
Private Wheel7() As Variant
Private Wheel8() As Variant
Private Wheel9() As Variant
Private Wheel10() As Variant
Private Wheel11() As Variant
Private Wheel12() As Variant
Private Wheel13() As Variant
Private Wheel14() As Variant
Private Wheel15() As Variant
Private Wheel16() As Variant
Private Wheel17() As Variant
Private Wheel18() As Variant
Private Wheel19() As Variant
Private Wheel20() As Variant

Public MyValue As Long

Private strW1 As String
Private strW2 As String
Private strW3 As String
Private strW4 As String
Private strW5 As String
Private strW6 As String
Private strW7 As String
Private strW8 As String
Private strW9 As String
Private strW10 As String
Private strW11 As String
Private strW12 As String
Private strW13 As String
Private strW14 As String
Private strW15 As String
Private strW16 As String
Private strW17 As String
Private strW18 As String
Private strW19 As String
Private strW20 As String

Private strMessage As String
Private strOutput As String
Private strDecryptOutput As String

Private Sub cmdDecrypt_Click()
    Decrypt
End Sub

Private Sub cmdEncrypt_Click()
    lNumTurns = 0
    iCurrentWheel = 1
    Encrypt
End Sub

Private Function EncryptRotateWheels()
    Dim k As Long
    Dim strTempHold As String

    ' Rotate Left
    If iCurrentWheel = 1 Then
        strTempHold = Wheel1(1)
        For k = 1 To (iTotalChar - 1) Step 1
            Wheel1(k) = Wheel1(k + 1)
        Next
        Wheel1(iTotalChar) = strTempHold
    End If
    
    ' Rotate Right
    If iCurrentWheel = 2 Then
        strTempHold = Wheel2(iTotalChar)
        For k = iTotalChar To 2 Step -1
            Wheel2(k) = Wheel2(k - 1)
        Next
        Wheel2(1) = strTempHold
    End If
    
    ' Rotate Left
    If iCurrentWheel = 3 Then
        strTempHold = Wheel3(1)
        For k = 1 To (iTotalChar - 1) Step 1
            Wheel3(k) = Wheel3(k + 1)
        Next
        Wheel3(iTotalChar) = strTempHold
    End If

    ' Rotate Right
    If iCurrentWheel = 4 Then
        strTempHold = Wheel4(iTotalChar)
        For k = iTotalChar To 2 Step -1
            Wheel4(k) = Wheel4(k - 1)
        Next
        Wheel4(1) = strTempHold
    End If

    ' Rotate Left
    If iCurrentWheel = 5 Then
        strTempHold = Wheel5(1)
        For k = 1 To (iTotalChar - 1) Step 1
            Wheel5(k) = Wheel5(k + 1)
        Next
        Wheel5(iTotalChar) = strTempHold
    End If

    ' Rotate Right
    If iCurrentWheel = 6 Then
        strTempHold = Wheel6(iTotalChar)
        For k = iTotalChar To 2 Step -1
            Wheel6(k) = Wheel6(k - 1)
        Next
        Wheel6(1) = strTempHold
    End If

    ' Rotate Left
    If iCurrentWheel = 7 Then
        strTempHold = Wheel7(1)
        For k = 1 To (iTotalChar - 1) Step 1
            Wheel7(k) = Wheel7(k + 1)
        Next
        Wheel7(iTotalChar) = strTempHold
    End If

    ' Rotate Right
    If iCurrentWheel = 8 Then
        strTempHold = Wheel8(iTotalChar)
        For k = iTotalChar To 2 Step -1
            Wheel8(k) = Wheel8(k - 1)
        Next
        Wheel8(1) = strTempHold
    End If

    ' Rotate Left
    If iCurrentWheel = 9 Then
        strTempHold = Wheel9(1)
        For k = 1 To (iTotalChar - 1) Step 1
            Wheel9(k) = Wheel9(k + 1)
        Next
        Wheel9(iTotalChar) = strTempHold
    End If

    ' Rotate Right
    If iCurrentWheel = 10 Then
        strTempHold = Wheel10(iTotalChar)
        For k = iTotalChar To 2 Step -1
            Wheel10(k) = Wheel10(k - 1)
        Next
        Wheel10(1) = strTempHold
    End If

    ' Rotate Left
    If iCurrentWheel = 11 Then
        strTempHold = Wheel11(1)
        For k = 1 To (iTotalChar - 1) Step 1
            Wheel11(k) = Wheel11(k + 1)
        Next
        Wheel11(iTotalChar) = strTempHold
    End If

    ' Rotate Right
    If iCurrentWheel = 12 Then
        strTempHold = Wheel12(iTotalChar)
        For k = iTotalChar To 2 Step -1
            Wheel12(k) = Wheel12(k - 1)
        Next
        Wheel12(1) = strTempHold
    End If

    ' Rotate Left
    If iCurrentWheel = 13 Then
        strTempHold = Wheel13(1)
        For k = 1 To (iTotalChar - 1) Step 1
            Wheel13(k) = Wheel13(k + 1)
        Next
        Wheel13(iTotalChar) = strTempHold
    End If

    ' Rotate Right
    If iCurrentWheel = 14 Then
        strTempHold = Wheel14(iTotalChar)
        For k = iTotalChar To 2 Step -1
            Wheel14(k) = Wheel14(k - 1)
        Next
        Wheel14(1) = strTempHold
    End If

    ' Rotate Left
    If iCurrentWheel = 15 Then
        strTempHold = Wheel15(1)
        For k = 1 To (iTotalChar - 1) Step 1
            Wheel15(k) = Wheel15(k + 1)
        Next
        Wheel15(iTotalChar) = strTempHold
    End If

    ' Rotate Right
    If iCurrentWheel = 16 Then
        strTempHold = Wheel16(iTotalChar)
        For k = iTotalChar To 2 Step -1
            Wheel16(k) = Wheel16(k - 1)
        Next
        Wheel16(1) = strTempHold
    End If

    ' Rotate Left
    If iCurrentWheel = 17 Then
        strTempHold = Wheel17(1)
        For k = 1 To (iTotalChar - 1) Step 1
            Wheel17(k) = Wheel17(k + 1)
        Next
        Wheel17(iTotalChar) = strTempHold
    End If

    ' Rotate Right
    If iCurrentWheel = 18 Then
        strTempHold = Wheel18(iTotalChar)
        For k = iTotalChar To 2 Step -1
            Wheel18(k) = Wheel18(k - 1)
        Next
        Wheel18(1) = strTempHold
    End If

    ' Rotate Left
    If iCurrentWheel = 19 Then
        strTempHold = Wheel19(1)
        For k = 1 To (iTotalChar - 1) Step 1
            Wheel19(k) = Wheel19(k + 1)
        Next
        Wheel19(iTotalChar) = strTempHold
    End If

    ' Rotate Right
    If iCurrentWheel = 20 Then
        strTempHold = Wheel20(iTotalChar)
        For k = iTotalChar To 2 Step -1
            Wheel20(k) = Wheel20(k - 1)
        Next
        Wheel20(1) = strTempHold
    End If

End Function

Private Function DecryptRotateWheels()
    Dim k As Long
    Dim strTempHold As String
    
    ' Rotate Left
    If iCurrentWheel = 20 Then
        strTempHold = Wheel20(1)
        For k = 1 To (iTotalChar - 1) Step 1
            Wheel20(k) = Wheel20(k + 1)
        Next
        Wheel20(iTotalChar) = strTempHold
    End If

    ' Rotate Right
    If iCurrentWheel = 19 Then
        strTempHold = Wheel19(iTotalChar)
        For k = iTotalChar To 2 Step -1
            Wheel19(k) = Wheel19(k - 1)
        Next
        Wheel19(1) = strTempHold
    End If

    ' Rotate Left
    If iCurrentWheel = 18 Then
        strTempHold = Wheel18(1)
        For k = 1 To (iTotalChar - 1) Step 1
            Wheel18(k) = Wheel18(k + 1)
        Next
        Wheel18(iTotalChar) = strTempHold
    End If

    ' Rotate Right
    If iCurrentWheel = 17 Then
        strTempHold = Wheel17(iTotalChar)
        For k = iTotalChar To 2 Step -1
            Wheel17(k) = Wheel17(k - 1)
        Next
        Wheel17(1) = strTempHold
    End If

    ' Rotate Left
    If iCurrentWheel = 16 Then
        strTempHold = Wheel16(1)
        For k = 1 To (iTotalChar - 1) Step 1
            Wheel16(k) = Wheel16(k + 1)
        Next
        Wheel16(iTotalChar) = strTempHold
    End If

    ' Rotate Right
    If iCurrentWheel = 15 Then
        strTempHold = Wheel15(iTotalChar)
        For k = iTotalChar To 2 Step -1
            Wheel15(k) = Wheel15(k - 1)
        Next
        Wheel15(1) = strTempHold
    End If

    ' Rotate Left
    If iCurrentWheel = 14 Then
        strTempHold = Wheel14(1)
        For k = 1 To (iTotalChar - 1) Step 1
            Wheel14(k) = Wheel14(k + 1)
        Next
        Wheel14(iTotalChar) = strTempHold
    End If

    ' Rotate Right
    If iCurrentWheel = 13 Then
        strTempHold = Wheel13(iTotalChar)
        For k = iTotalChar To 2 Step -1
            Wheel13(k) = Wheel13(k - 1)
        Next
        Wheel13(1) = strTempHold
    End If

    ' Rotate Left
    If iCurrentWheel = 12 Then
        strTempHold = Wheel12(1)
        For k = 1 To (iTotalChar - 1) Step 1
            Wheel12(k) = Wheel12(k + 1)
        Next
        Wheel12(iTotalChar) = strTempHold
    End If

    ' Rotate Right
    If iCurrentWheel = 11 Then
        strTempHold = Wheel11(iTotalChar)
        For k = iTotalChar To 2 Step -1
            Wheel11(k) = Wheel11(k - 1)
        Next
        Wheel11(1) = strTempHold
    End If

    ' Rotate Left
    If iCurrentWheel = 10 Then
        strTempHold = Wheel10(1)
        For k = 1 To (iTotalChar - 1) Step 1
            Wheel10(k) = Wheel10(k + 1)
        Next
        Wheel10(iTotalChar) = strTempHold
    End If

    ' Rotate Right
    If iCurrentWheel = 9 Then
        strTempHold = Wheel9(iTotalChar)
        For k = iTotalChar To 2 Step -1
            Wheel9(k) = Wheel9(k - 1)
        Next
        Wheel9(1) = strTempHold
    End If

    ' Rotate Left
    If iCurrentWheel = 8 Then
        strTempHold = Wheel8(1)
        For k = 1 To (iTotalChar - 1) Step 1
            Wheel8(k) = Wheel8(k + 1)
        Next
        Wheel8(iTotalChar) = strTempHold
    End If

    ' Rotate Right
    If iCurrentWheel = 7 Then
        strTempHold = Wheel7(iTotalChar)
        For k = iTotalChar To 2 Step -1
            Wheel7(k) = Wheel7(k - 1)
        Next
        Wheel7(1) = strTempHold
    End If

    ' Rotate Left
    If iCurrentWheel = 6 Then
        strTempHold = Wheel6(1)
        For k = 1 To (iTotalChar - 1) Step 1
            Wheel6(k) = Wheel6(k + 1)
        Next
        Wheel6(iTotalChar) = strTempHold
    End If

    ' Rotate Right
    If iCurrentWheel = 5 Then
        strTempHold = Wheel5(iTotalChar)
        For k = iTotalChar To 2 Step -1
            Wheel5(k) = Wheel5(k - 1)
        Next
        Wheel5(1) = strTempHold
    End If

    ' Rotate Left
    If iCurrentWheel = 4 Then
        strTempHold = Wheel4(1)
        For k = 1 To (iTotalChar - 1) Step 1
            Wheel4(k) = Wheel4(k + 1)
        Next
        Wheel4(iTotalChar) = strTempHold
    End If

    ' Rotate Right
    If iCurrentWheel = 3 Then
        strTempHold = Wheel3(iTotalChar)
        For k = iTotalChar To 2 Step -1
            Wheel3(k) = Wheel3(k - 1)
        Next
        Wheel3(1) = strTempHold
    End If

    ' Rotate Left
    If iCurrentWheel = 2 Then
        strTempHold = Wheel2(1)
        For k = 1 To (iTotalChar - 1) Step 1
            Wheel2(k) = Wheel2(k + 1)
        Next
        Wheel2(iTotalChar) = strTempHold
    End If

    ' Rotate Right
    If iCurrentWheel = 1 Then
        strTempHold = Wheel1(iTotalChar)
        For k = iTotalChar To 2 Step -1
            Wheel1(k) = Wheel1(k - 1)
        Next
        Wheel1(1) = strTempHold
    End If

End Function

Private Function Encrypt()
    
    EnigmaProgress.Value = 0
    lProcessCount = 0
    
    iTotalChar = Len(strW1)
    
    ReDim Wheel1(1 To iTotalChar) As Variant
    ReDim Wheel2(1 To iTotalChar) As Variant
    ReDim Wheel3(1 To iTotalChar) As Variant
    ReDim Wheel4(1 To iTotalChar) As Variant
    ReDim Wheel5(1 To iTotalChar) As Variant
    ReDim Wheel6(1 To iTotalChar) As Variant
    ReDim Wheel7(1 To iTotalChar) As Variant
    ReDim Wheel8(1 To iTotalChar) As Variant
    ReDim Wheel9(1 To iTotalChar) As Variant
    ReDim Wheel10(1 To iTotalChar) As Variant
    ReDim Wheel11(1 To iTotalChar) As Variant
    ReDim Wheel12(1 To iTotalChar) As Variant
    ReDim Wheel13(1 To iTotalChar) As Variant
    ReDim Wheel14(1 To iTotalChar) As Variant
    ReDim Wheel15(1 To iTotalChar) As Variant
    ReDim Wheel16(1 To iTotalChar) As Variant
    ReDim Wheel17(1 To iTotalChar) As Variant
    ReDim Wheel18(1 To iTotalChar) As Variant
    ReDim Wheel19(1 To iTotalChar) As Variant
    ReDim Wheel20(1 To iTotalChar) As Variant
    
    For i = 1 To iTotalChar Step 1
        Wheel1(i) = Mid(strW1, i, 1)
        Wheel2(i) = Mid(strW2, i, 1)
        Wheel3(i) = Mid(strW3, i, 1)
        Wheel4(i) = Mid(strW4, i, 1)
        Wheel5(i) = Mid(strW5, i, 1)
        Wheel6(i) = Mid(strW6, i, 1)
        Wheel7(i) = Mid(strW7, i, 1)
        Wheel8(i) = Mid(strW8, i, 1)
        Wheel9(i) = Mid(strW9, i, 1)
        Wheel10(i) = Mid(strW10, i, 1)
        Wheel11(i) = Mid(strW11, i, 1)
        Wheel12(i) = Mid(strW12, i, 1)
        Wheel13(i) = Mid(strW13, i, 1)
        Wheel14(i) = Mid(strW14, i, 1)
        Wheel15(i) = Mid(strW15, i, 1)
        Wheel16(i) = Mid(strW16, i, 1)
        Wheel17(i) = Mid(strW17, i, 1)
        Wheel18(i) = Mid(strW18, i, 1)
        Wheel19(i) = Mid(strW19, i, 1)
        Wheel20(i) = Mid(strW20, i, 1)
    Next

    strMessage = txtEnigma.Text
    
    For j = 1 To Len(strMessage) Step 1

        strCurrentChar = Mid(strMessage, j, 1)
    
        If strCurrentChar <> Chr(13) And strCurrentChar <> Chr(10) And strCurrentChar <> Chr(34) Then 'And strCurrentChar <> "|" Then
                    
            If TotalWheels >= 2 Then
            
                For k = 1 To iTotalChar Step 1
                    If strCurrentChar = Wheel1(k) Then
                        WheelIndex1 = k
                    End If
                Next
                strCurrentChar = Wheel2(WheelIndex1)
            End If
            
            If TotalWheels >= 4 Then
                    
                For k = 1 To iTotalChar Step 1
                    If strCurrentChar = Wheel3(k) Then
                        WheelIndex2 = k
                    End If
                Next
                strCurrentChar = Wheel4(WheelIndex2)
            End If

            If TotalWheels >= 6 Then
                    
                For k = 1 To iTotalChar Step 1
                    If strCurrentChar = Wheel5(k) Then
                        WheelIndex3 = k
                    End If
                Next
                strCurrentChar = Wheel6(WheelIndex3)
            End If

            If TotalWheels >= 8 Then
                    
                For k = 1 To iTotalChar Step 1
                    If strCurrentChar = Wheel7(k) Then
                        WheelIndex4 = k
                    End If
                Next
                strCurrentChar = Wheel8(WheelIndex4)
            End If

            If TotalWheels >= 10 Then
                    
                For k = 1 To iTotalChar Step 1
                    If strCurrentChar = Wheel9(k) Then
                        WheelIndex5 = k
                    End If
                Next
                strCurrentChar = Wheel10(WheelIndex5)
            End If

            If TotalWheels >= 12 Then
                    
                For k = 1 To iTotalChar Step 1
                    If strCurrentChar = Wheel11(k) Then
                        WheelIndex6 = k
                    End If
                Next
                strCurrentChar = Wheel12(WheelIndex6)
            End If

            If TotalWheels >= 14 Then
                    
                For k = 1 To iTotalChar Step 1
                    If strCurrentChar = Wheel13(k) Then
                        WheelIndex7 = k
                    End If
                Next
                strCurrentChar = Wheel14(WheelIndex7)
            End If

            If TotalWheels >= 16 Then
                    
                For k = 1 To iTotalChar Step 1
                    If strCurrentChar = Wheel15(k) Then
                        WheelIndex8 = k
                    End If
                Next
                strCurrentChar = Wheel16(WheelIndex8)
            End If

            If TotalWheels >= 18 Then
                    
                For k = 1 To iTotalChar Step 1
                    If strCurrentChar = Wheel17(k) Then
                        WheelIndex9 = k
                    End If
                Next
                strCurrentChar = Wheel18(WheelIndex9)
            End If

            If TotalWheels >= 20 Then
                    
                For k = 1 To iTotalChar Step 1
                    If strCurrentChar = Wheel19(k) Then
                        WheelIndex10 = k
                    End If
                Next
                strCurrentChar = Wheel20(WheelIndex10)
            
            End If
            strOutput = strOutput & strCurrentChar
        
        Else
            
            strOutput = strOutput & strCurrentChar
            
        End If
        
        EncryptRotateWheels
        
        iCurrentWheel = iCurrentWheel + 1
        If iCurrentWheel = TotalWheels + 1 Then
            iCurrentWheel = 1
        End If
        
        lProcessCount = lProcessCount + 1
        
        If CInt(EnigmaProgress.Value) <> CInt((lProcessCount / Len(strMessage)) * 100) Then
            EnigmaProgress.Value = CInt((lProcessCount / Len(strMessage)) * 100)
            frmEnigma.Caption = "%" & CStr(CInt(EnigmaProgress.Value)) & " The Enigma Machine"
        End If
        DoEvents
    Next

    txtOutputBox.Text = strOutput
    txtEnigma.Text = ""
    strMessage = ""
    strOutput = ""
    
End Function


Private Function Decrypt()
    Dim d As Long
    Dim t As Long
    Dim s As Long
    Dim u As Long
    
    EnigmaProgress.Value = 0
    lProcessCount = 0
    
    iTotalChar = Len(strW1)
    
    ReDim Wheel1(1 To iTotalChar) As Variant
    ReDim Wheel2(1 To iTotalChar) As Variant
    ReDim Wheel3(1 To iTotalChar) As Variant
    ReDim Wheel4(1 To iTotalChar) As Variant
    ReDim Wheel5(1 To iTotalChar) As Variant
    ReDim Wheel6(1 To iTotalChar) As Variant
    ReDim Wheel7(1 To iTotalChar) As Variant
    ReDim Wheel8(1 To iTotalChar) As Variant
    ReDim Wheel9(1 To iTotalChar) As Variant
    ReDim Wheel10(1 To iTotalChar) As Variant
    ReDim Wheel11(1 To iTotalChar) As Variant
    ReDim Wheel12(1 To iTotalChar) As Variant
    ReDim Wheel13(1 To iTotalChar) As Variant
    ReDim Wheel14(1 To iTotalChar) As Variant
    ReDim Wheel15(1 To iTotalChar) As Variant
    ReDim Wheel16(1 To iTotalChar) As Variant
    ReDim Wheel17(1 To iTotalChar) As Variant
    ReDim Wheel18(1 To iTotalChar) As Variant
    ReDim Wheel19(1 To iTotalChar) As Variant
    ReDim Wheel20(1 To iTotalChar) As Variant
    
    For i = 1 To iTotalChar Step 1
        Wheel1(i) = Mid(strW1, i, 1)
        Wheel2(i) = Mid(strW2, i, 1)
        Wheel3(i) = Mid(strW3, i, 1)
        Wheel4(i) = Mid(strW4, i, 1)
        Wheel5(i) = Mid(strW5, i, 1)
        Wheel6(i) = Mid(strW6, i, 1)
        Wheel7(i) = Mid(strW7, i, 1)
        Wheel8(i) = Mid(strW8, i, 1)
        Wheel9(i) = Mid(strW9, i, 1)
        Wheel10(i) = Mid(strW10, i, 1)
        Wheel11(i) = Mid(strW11, i, 1)
        Wheel12(i) = Mid(strW12, i, 1)
        Wheel13(i) = Mid(strW13, i, 1)
        Wheel14(i) = Mid(strW14, i, 1)
        Wheel15(i) = Mid(strW15, i, 1)
        Wheel16(i) = Mid(strW16, i, 1)
        Wheel17(i) = Mid(strW17, i, 1)
        Wheel18(i) = Mid(strW18, i, 1)
        Wheel19(i) = Mid(strW19, i, 1)
        Wheel20(i) = Mid(strW20, i, 1)
    Next
    
    strOutput = txtOutputBox.Text
    iCurrentWheel = 1
    
    ' This If Structure reads in the encrypted text, and determines the wheel positions that would have been
    ' present at the point where the encryption was finished. This way, reading in the text backwards through
    ' a reversed process reveals the message in it's true form.
    If Len(strOutput) <= iTotalChar Then
        For d = 1 To Len(strOutput) Step 1
            EncryptRotateWheels
            iCurrentWheel = iCurrentWheel + 1
            If iCurrentWheel > TotalWheels Then
                iCurrentWheel = 1
            End If
        Next
    Else
        t = CLng(((Len(strOutput)) / TotalWheels))
        If Len(strOutput) < (TotalWheels * t) Then
            t = t - 1
        End If
        s = CLng(t) Mod iTotalChar
        
        For iCurrentWheel = 1 To TotalWheels Step 1
            For d = 1 To t Step 1
                EncryptRotateWheels
            Next
        Next
        
        iCurrentWheel = 1
        
        For d = 1 To Len(strOutput) Mod TotalWheels Step 1
            EncryptRotateWheels
            iCurrentWheel = iCurrentWheel + 1
            If iCurrentWheel > TotalWheels Then
                iCurrentWheel = 1
            End If
        Next
    End If
    
    txtEnigma.Text = ""
    
    EnigmaProgress.Value = 0
    lProcessCount = 0
    
    ' Decrypt
    For j = Len(strOutput) To 1 Step -1
        
        iCurrentWheel = iCurrentWheel - 1
        If iCurrentWheel < 1 Then
            iCurrentWheel = TotalWheels
        End If

        DecryptRotateWheels
        
        strCurrentChar = Mid(strOutput, j, 1)
                
                
        If strCurrentChar <> Chr(13) And strCurrentChar <> Chr(10) And strCurrentChar <> Chr(34) Then ' And strCurrentChar <> "|" Then

            If TotalWheels >= 20 Then
                    
                For k = 1 To iTotalChar Step 1
                    If strCurrentChar = Wheel20(k) Then
                        WheelIndex10 = k
                    End If
                Next
                strCurrentChar = Wheel19(WheelIndex10)
            End If

            If TotalWheels >= 18 Then
                    
                For k = 1 To iTotalChar Step 1
                    If strCurrentChar = Wheel18(k) Then
                        WheelIndex9 = k
                    End If
                Next
                strCurrentChar = Wheel17(WheelIndex9)
            End If
            
            If TotalWheels >= 16 Then
                    
                For k = 1 To iTotalChar Step 1
                    If strCurrentChar = Wheel16(k) Then
                        WheelIndex8 = k
                    End If
                Next
                strCurrentChar = Wheel15(WheelIndex8)
            End If
    
            If TotalWheels >= 14 Then
                    
                For k = 1 To iTotalChar Step 1
                    If strCurrentChar = Wheel14(k) Then
                        WheelIndex7 = k
                    End If
                Next
                strCurrentChar = Wheel13(WheelIndex7)
            End If
    
            If TotalWheels >= 12 Then
                    
                For k = 1 To iTotalChar Step 1
                    If strCurrentChar = Wheel12(k) Then
                        WheelIndex6 = k
                    End If
                Next
                strCurrentChar = Wheel11(WheelIndex6)
            End If
    
            If TotalWheels >= 10 Then
                    
                For k = 1 To iTotalChar Step 1
                    If strCurrentChar = Wheel10(k) Then
                        WheelIndex5 = k
                    End If
                Next
                strCurrentChar = Wheel9(WheelIndex5)
            End If
    
            If TotalWheels >= 8 Then
                    
                For k = 1 To iTotalChar Step 1
                    If strCurrentChar = Wheel8(k) Then
                        WheelIndex4 = k
                    End If
                Next
                strCurrentChar = Wheel7(WheelIndex4)
            End If
    
            If TotalWheels >= 6 Then
                    
                For k = 1 To iTotalChar Step 1
                    If strCurrentChar = Wheel6(k) Then
                        WheelIndex3 = k
                    End If
                Next
                strCurrentChar = Wheel5(WheelIndex3)
            End If
    
            If TotalWheels >= 4 Then
                    
                For k = 1 To iTotalChar Step 1
                    If strCurrentChar = Wheel4(k) Then
                        WheelIndex2 = k
                    End If
                Next
                strCurrentChar = Wheel3(WheelIndex2)
            End If
                        
            If TotalWheels >= 2 Then
                    
                For k = 1 To iTotalChar Step 1
                    If strCurrentChar = Wheel2(k) Then
                        WheelIndex1 = k
                    End If
                Next
                strCurrentChar = Wheel1(WheelIndex1)
            End If
            
            strDecryptOutput = strCurrentChar & strDecryptOutput

        Else
            
            strDecryptOutput = strCurrentChar & strDecryptOutput
        
        End If

        lProcessCount = lProcessCount + 1
        If CInt(EnigmaProgress.Value) <> CInt((lProcessCount / Len(strOutput)) * 100) Then
            EnigmaProgress.Value = CInt((lProcessCount / Len(strOutput)) * 100)
            frmEnigma.Caption = "%" & CStr(CInt(EnigmaProgress.Value)) & " The Enigma Machine"
        End If
        DoEvents

    Next
    
    txtEnigma.Text = strDecryptOutput
    txtOutputBox.Text = ""
    strMessage = ""
    strOutput = ""
    strDecryptOutput = ""

End Function

Private Sub GenerateRandomWheels()
    Dim i As Long
    Dim j As Long

    Dim bDupe As Boolean

    Randomize

    For i = 1 To iTotalChar Step 1
        bDupe = True
        While bDupe = True
            MyValue = Int((iTotalChar * Rnd) + 1)    ' Generate random value between 1 and iTotalChar.
            For j = 1 To i Step 1
                If Wheel2(j) = MyValue Then
                    bDupe = True
                    Exit For
                Else
                    If Wheel2(j) = "" Then
                        bDupe = False
                        Wheel2(j) = MyValue
                    End If
                End If
            Next
        Wend
    Next

    For i = 1 To iTotalChar Step 1
        bDupe = True
        While bDupe = True
            MyValue = Int((iTotalChar * Rnd) + 1)    ' Generate random value between 1 and iTotalChar.
            For j = 1 To i Step 1
                If Wheel3(j) = MyValue Then
                    bDupe = True
                    Exit For
                Else
                    If Wheel3(j) = "" Then
                        bDupe = False
                        Wheel3(j) = MyValue
                    End If
                End If
            Next
        Wend
    Next

    For i = 1 To iTotalChar Step 1
        bDupe = True
        While bDupe = True
            MyValue = Int((iTotalChar * Rnd) + 1)    ' Generate random value between 1 and iTotalChar.
            For j = 1 To i Step 1
                If Wheel4(j) = MyValue Then
                    bDupe = True
                    Exit For
                Else
                    If Wheel4(j) = "" Then
                        bDupe = False
                        Wheel4(j) = MyValue
                    End If
                End If
            Next
        Wend
    Next

    For i = 1 To iTotalChar Step 1
        bDupe = True
        While bDupe = True
            MyValue = Int((iTotalChar * Rnd) + 1)    ' Generate random value between 1 and iTotalChar.
            For j = 1 To i Step 1
                If Wheel5(j) = MyValue Then
                    bDupe = True
                    Exit For
                Else
                    If Wheel5(j) = "" Then
                        bDupe = False
                        Wheel5(j) = MyValue
                    End If
                End If
            Next
        Wend
    Next

    For i = 1 To iTotalChar Step 1
        bDupe = True
        While bDupe = True
            MyValue = Int((iTotalChar * Rnd) + 1)    ' Generate random value between 1 and iTotalChar.
            For j = 1 To i Step 1
                If Wheel6(j) = MyValue Then
                    bDupe = True
                    Exit For
                Else
                    If Wheel6(j) = "" Then
                        bDupe = False
                        Wheel6(j) = MyValue
                    End If
                End If
            Next
        Wend
    Next
    For i = 1 To iTotalChar Step 1
        bDupe = True
        While bDupe = True
            MyValue = Int((iTotalChar * Rnd) + 1)    ' Generate random value between 1 and iTotalChar.
            For j = 1 To i Step 1
                If Wheel7(j) = MyValue Then
                    bDupe = True
                    Exit For
                Else
                    If Wheel7(j) = "" Then
                        bDupe = False
                        Wheel7(j) = MyValue
                    End If
                End If
            Next
        Wend
    Next
    For i = 1 To iTotalChar Step 1
        bDupe = True
        While bDupe = True
            MyValue = Int((iTotalChar * Rnd) + 1)    ' Generate random value between 1 and iTotalChar.
            For j = 1 To i Step 1
                If Wheel8(j) = MyValue Then
                    bDupe = True
                    Exit For
                Else
                    If Wheel8(j) = "" Then
                        bDupe = False
                        Wheel8(j) = MyValue
                    End If
                End If
            Next
        Wend
    Next
    For i = 1 To iTotalChar Step 1
        bDupe = True
        While bDupe = True
            MyValue = Int((iTotalChar * Rnd) + 1)    ' Generate random value between 1 and iTotalChar.
            For j = 1 To i Step 1
                If Wheel9(j) = MyValue Then
                    bDupe = True
                    Exit For
                Else
                    If Wheel9(j) = "" Then
                        bDupe = False
                        Wheel9(j) = MyValue
                    End If
                End If
            Next
        Wend
    Next
    For i = 1 To iTotalChar Step 1
        bDupe = True
        While bDupe = True
            MyValue = Int((iTotalChar * Rnd) + 1)    ' Generate random value between 1 and iTotalChar.
            For j = 1 To i Step 1
                If Wheel10(j) = MyValue Then
                    bDupe = True
                    Exit For
                Else
                    If Wheel10(j) = "" Then
                        bDupe = False
                        Wheel10(j) = MyValue
                    End If
                End If
            Next
        Wend
    Next


    For i = 1 To iTotalChar Step 1
        Wheel2String = Wheel2String & CStr(Wheel1(Wheel2(i)))
    Next
    For i = 1 To iTotalChar Step 1
        Wheel3String = Wheel3String & CStr(Wheel1(Wheel3(i)))
    Next
    For i = 1 To iTotalChar Step 1
        Wheel4String = Wheel4String & CStr(Wheel1(Wheel4(i)))
    Next
    For i = 1 To iTotalChar Step 1
        Wheel5String = Wheel5String & CStr(Wheel1(Wheel5(i)))
    Next
    For i = 1 To iTotalChar Step 1
        Wheel6String = Wheel6String & CStr(Wheel1(Wheel6(i)))
    Next
    For i = 1 To iTotalChar Step 1
        Wheel7String = Wheel7String & CStr(Wheel1(Wheel7(i)))
    Next
    For i = 1 To iTotalChar Step 1
        Wheel8String = Wheel8String & CStr(Wheel1(Wheel8(i)))
    Next
    For i = 1 To iTotalChar Step 1
        Wheel9String = Wheel9String & CStr(Wheel1(Wheel9(i)))
    Next
    For i = 1 To iTotalChar Step 1
        Wheel10String = Wheel10String & CStr(Wheel1(Wheel10(i)))
    Next
    
    Debug.Print Wheel2String
    Debug.Print Wheel3String
    Debug.Print Wheel4String
    Debug.Print Wheel5String
    Debug.Print Wheel6String
    Debug.Print Wheel7String
    Debug.Print Wheel8String
    Debug.Print Wheel9String
    Debug.Print Wheel10String

End Sub

Private Sub Command1_Click()
    Dim testtext As String
    Dim l As Long
    
    testtext = txtEnigma.Text
    
    Encrypt
    
    While txtEnigma.Text <> testtext
        DoEvents
        Encrypt
        l = l + 1
    Wend
    
    Label1.Caption = CStr(l)
End Sub

Private Sub Form_Load()

    strW1 = ">Ebvp#-osBj9R7cFS3Oa M'&G@uN5]Xq_|i8kdlY%D*JeZ/+}!;^{Hh=.I4`\fT?x[6ntCAV2gmzQL1$,yW0~<(UrK:w)P"
    strW2 = ";y8'iT@L7Q~^s->Bw(n*h3N?PAY5po#`.fjSG{+_}:UMz4 D%|t1H/K\2=<[b!cxJlIFXeWamC0)&9,guEdR]Zv6$krOVq"
    strW3 = "%lzet$.v|0si,@SZx;(!p#L'*JH~j}3bd=72Da[C-qk>5fy]?gRW/MEm6 u8FB\4oInQ)wYXUG1A&NKTc:+{_`O^9VrhP<"
    strW4 = "YKq$4iX&3ECt;r2ed8,c]J^>k.Fn{MR_u)0sD#a\L'T%/?IPjv9Zw=`f5U*OgGNz<-x}A(mB:@7[1o lb!h|VS6WHp+Qy~"
    strW5 = "{&A[IoNp *`4B@X-Di7b#;MKRVPv'wt)Gu\5gWQ]0_q}Z:8s(Uk$rYdE9a3fjnCF=!lL,e/mJ|1SzH<y6h?Ox~^cT2.%>+"
    strW6 = ";u>i1$8EOa:o/&t0~Yd%5cBL)4 gI=,vTJCFxN3RShVmX_9@2p*-zbfHn+Q|qKUD6^M[(\r?{']Z}!#jAWy.`sPkw7lGe<"
    strW7 = "W*V-:kv^MB_~FUCO?+pSEn6Go/P#Rz'D|H8N]1LJTd$;txX!@95 }f)yj3(r%ic,b[{>eAQmsYq\IwK7hau&`0Z<4.=l2g"
    strW8 = "$r@^wIPcDF!}yE6'jNAYOVRK;g`~o>{_ 48+0W/7d|GmUHeCBftJp=z%[,qSv]la-L#*()Msb<&h5:3?ix\9.XQnZ12ukT"
    strW9 = "76W|XG9,n{Pis_]H-@t[Tp+lJU48SbV=OE\yv0}C'a%1zAo2j(wf#h :QmZDN~gI&Lqe.<$>k*;5K)cd!R`B3Yx/ur?^FM"
    strW10 = "Ya0L_kq1Ax=t|CUbBRTDI&S[@>NOc`Wo*Pp;hE-)Kj#f8wV}lQ J29$7nsdG\~:!F,%gimz'3H6e]({.+?<54M/uvryZX^"
    strW11 = "G[gXIs3Azx1?ScB-8H$qjp{2#&*KJ=o)95L+|aNvm\F>E67<0RiZ(M~_n, C:Wr`u}.%fk]OVyU/bDlh@Q!P'^;w4TdtYe"
    strW12 = "n)7Blafizdh]s`$u^,#LRGN6Djv'U@yT(=:~EF0[8J3\AWt9Hrg|I eCS.1V4pb%/5O-kmcQwXx>;Zqo?+!K&}<PYM{2_*"
    strW13 = "J5)[>:cupBSm@qjwV^s*~_%o-7viR`,9.0x<41dIl|?XfZPzbhATtYGkNM&]!;2+($r \KyQ{F6DO/8'CUEan#eHgW3=L}"
    strW14 = "Knj[pZ7F|NE9r?^{eIW!f=C,d$kP*;Sz8%Lq 4V>()51O~63&0+X#Qa/<H@i\2}msUTx'-JYbw_vuAthDcyglGoB]M`:.R"
    strW15 = "MQxc-YjkO$%f)dq;n@s4+]l>G[e~3HwW:'_v8y EhU1}L5mi`VztS0/B{,?^=69F7uCa#DJ|&oAZg.XPI<\rK(RNb2*Tp!"
    strW16 = "14tdb7(?]%/[Dx6hK&AJjswzyV3Yalm_r5WRi9Z{HkBnX>2)ouC!f^*O8~+MP\G;}pe '#<$@N0E`gUQTvL|SIq,.c-:=F"
    strW17 = "0q%buT[Q=7V1t<o?jB@N>GAfPY 6c^a!{]5H('4sU~RS*-/ClLdiKM.;8$\`x|vekhZrX:D2#&m+I)O9_nyW}E3gzwpF,J"
    strW18 = "M(;Qt?Oo9e+Rfz-w50'mUEJkHj][P|NqKs8>Ynivl1%b _,4}VW2C!6Bhgprud*)Z=\.&S^FAG`<y{#c/aI73~XLTD$:x@"
    strW19 = "/aH7=<bkA]hf_3w;jo|BK9Vm.[NEg!:#M-1rv $WSp?0Z(J4Q}e~Id2@G&)z5unRyX8l*D+i^xC%UYs6cO',`q\TF{t>PL"
    strW20 = "5|H$9h/~Cn+=`3qFs)@]^xO!zo>*{1T8d:yDeib#0ER(%ZVpvJ?W-<}Q[SLU;w aAc.\NGf2,turlK6m_gP7Xk&Bj4Y'MI"



    iTotalChar = Len(strW1)
    TotalWheels = 10
    txtNumWheels.Text = CStr(TotalWheels)
'
'    ReDim Wheel1(1 To iTotalChar) As Variant
'    ReDim Wheel2(1 To iTotalChar) As Variant
'    ReDim Wheel3(1 To iTotalChar) As Variant
'    ReDim Wheel4(1 To iTotalChar) As Variant
'    ReDim Wheel5(1 To iTotalChar) As Variant
'    ReDim Wheel6(1 To iTotalChar) As Variant
'    ReDim Wheel7(1 To iTotalChar) As Variant
'    ReDim Wheel8(1 To iTotalChar) As Variant
'    ReDim Wheel9(1 To iTotalChar) As Variant
'    ReDim Wheel10(1 To iTotalChar) As Variant
'    ReDim Wheel11(1 To iTotalChar) As Variant
'    ReDim Wheel12(1 To iTotalChar) As Variant
'    ReDim Wheel13(1 To iTotalChar) As Variant
'    ReDim Wheel14(1 To iTotalChar) As Variant
'    ReDim Wheel15(1 To iTotalChar) As Variant
'    ReDim Wheel16(1 To iTotalChar) As Variant
'    ReDim Wheel17(1 To iTotalChar) As Variant
'    ReDim Wheel18(1 To iTotalChar) As Variant
'    ReDim Wheel19(1 To iTotalChar) As Variant
'    ReDim Wheel20(1 To iTotalChar) As Variant
'
'    For i = 1 To iTotalChar Step 1
'        Wheel1(i) = Mid(strW1, i, 1)
'    Next
'
'    GenerateRandomWheels
    EnigmaProgress.Value = 0

End Sub

Private Sub UpDown1_UpClick()
    TotalWheels = TotalWheels + 2
    
    If TotalWheels > 20 Then
        TotalWheels = 20
    End If
    
    txtNumWheels.Text = CStr(TotalWheels)
End Sub

Private Sub UpDown1_DownClick()
    TotalWheels = TotalWheels - 2
    
    If TotalWheels <= 0 Then
        TotalWheels = 2
    End If
    
    txtNumWheels.Text = CStr(TotalWheels)
End Sub
