VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kana Entry"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Text            =   "katakana.ini"
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load Font Info"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy Kana To Clipboard"
      Height          =   495
      Left            =   5160
      TabIndex        =   4
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   7095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7095
   End
   Begin VB.Label lblWarning 
      Caption         =   "Font map ini not loaded!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Western:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Kana:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FontINI As New INI

Dim str1, str2, str3, str4, str5, str6, str7, str8, str9, str0 As String
Dim strA, strE, strI, strO, strU, strDot As String
Dim strSmallA, strSmallE, strSmallI, strSmallO, strSmallU As String
Dim strSmallTSU, strLength As String
Dim strBA, strBE, strBI, strBO, strBU As String
Dim strKA, strKE, strKI, strKO, strKU As String
Dim strGA, strGE, strGI, strGO, strGU As String
Dim strHA, strHE, strHI, strHO, strHU As String
Dim strMA, strME, strMI, strMO, strMU As String
Dim strNA, strNE, strNI, strNO, strNU, strN As String
Dim strPA, strPE, strPI, strPO, strPU As String
Dim strRA, strRE, strRI, strRO, strRU As String
Dim strSA, strSE, strSHI, strSO, strSU As String
Dim strTA, strTE, strCHI, strTO, strTSU As String
Dim strDA, strDE, strDI, strDO, strDU As String
Dim strYA, strYO, strYU As String
Dim strSmallYA, strSmallYO, strSmallYU As String
Dim strZA, strZE, strZI, strZO, strZU As String
Dim strWA, strWO As String
Dim INILoaded As Boolean

Private Sub Command1_Click()

Clipboard.Clear
Clipboard.SetText (Text2.Text)

End Sub

Private Sub Command2_Click()

If Dir(App.Path & "\" & Text3.Text) = "" Then
    MsgBox ("Fortune cookie say:  Font map ini file: '" & Text3.Text & "' does not exist.")
    Exit Sub
End If

FontINI.inifile = App.Path & "\" & Text3.Text

Text2.Font.Name = FontINI.ReadINI("Font", "Name")
Text2.Font.Size = FontINI.ReadINI("Font", "Size")

str1 = FontINI.ReadINI("Numbers", "1")
str2 = FontINI.ReadINI("Numbers", "2")
str3 = FontINI.ReadINI("Numbers", "3")
str4 = FontINI.ReadINI("Numbers", "4")
str5 = FontINI.ReadINI("Numbers", "5")
str6 = FontINI.ReadINI("Numbers", "6")
str7 = FontINI.ReadINI("Numbers", "7")
str8 = FontINI.ReadINI("Numbers", "8")
str9 = FontINI.ReadINI("Numbers", "9")
str0 = FontINI.ReadINI("Numbers", "0")
strDot = FontINI.ReadINI("Numbers", "Dot")


strA = FontINI.ReadINI("Vowels", "A")
strE = FontINI.ReadINI("Vowels", "E")
strI = FontINI.ReadINI("Vowels", "I")
strO = FontINI.ReadINI("Vowels", "O")
strU = FontINI.ReadINI("Vowels", "U")

strSmallA = FontINI.ReadINI("Vowels", "SmallA")
strSmallE = FontINI.ReadINI("Vowels", "SmallE")
strSmallI = FontINI.ReadINI("Vowels", "SmallI")
strSmallO = FontINI.ReadINI("Vowels", "SmallO")
strSmallU = FontINI.ReadINI("Vowels", "SmallU")

strSmallTSU = FontINI.ReadINI("Vowels", "SmallTSU")
strLength = FontINI.ReadINI("Vowels", "Length")

strBA = FontINI.ReadINI("B", "BA")
strBE = FontINI.ReadINI("B", "BE")
strBI = FontINI.ReadINI("B", "BI")
strBO = FontINI.ReadINI("B", "BO")
strBU = FontINI.ReadINI("B", "BU")

strDA = FontINI.ReadINI("D", "DA")
strDE = FontINI.ReadINI("D", "DE")
strDI = FontINI.ReadINI("D", "DI")
strDO = FontINI.ReadINI("D", "DO")
strDU = FontINI.ReadINI("D", "DU")

strGA = FontINI.ReadINI("G", "GA")
strGE = FontINI.ReadINI("G", "GE")
strGI = FontINI.ReadINI("G", "GI")
strGO = FontINI.ReadINI("G", "GO")
strGU = FontINI.ReadINI("G", "GU")

strHA = FontINI.ReadINI("H", "HA")
strHE = FontINI.ReadINI("H", "HE")
strHI = FontINI.ReadINI("H", "HI")
strHO = FontINI.ReadINI("H", "HO")
strHU = FontINI.ReadINI("H", "HU")

strKA = FontINI.ReadINI("K", "KA")
strKE = FontINI.ReadINI("K", "KE")
strKI = FontINI.ReadINI("K", "KI")
strKO = FontINI.ReadINI("K", "KO")
strKU = FontINI.ReadINI("K", "KU")

strMA = FontINI.ReadINI("M", "MA")
strME = FontINI.ReadINI("M", "ME")
strMI = FontINI.ReadINI("M", "MI")
strMO = FontINI.ReadINI("M", "MO")
strMU = FontINI.ReadINI("M", "MU")

strNA = FontINI.ReadINI("N", "NA")
strNE = FontINI.ReadINI("N", "NE")
strNI = FontINI.ReadINI("N", "NI")
strNO = FontINI.ReadINI("N", "NO")
strNU = FontINI.ReadINI("N", "NU")
strN = FontINI.ReadINI("N", "N")

strPA = FontINI.ReadINI("P", "PA")
strPE = FontINI.ReadINI("P", "PE")
strPI = FontINI.ReadINI("P", "PI")
strPO = FontINI.ReadINI("P", "PO")
strPU = FontINI.ReadINI("P", "PU")

strRA = FontINI.ReadINI("R", "RA")
strRE = FontINI.ReadINI("R", "RE")
strRI = FontINI.ReadINI("R", "RI")
strRO = FontINI.ReadINI("R", "RO")
strRU = FontINI.ReadINI("R", "RU")

strSA = FontINI.ReadINI("S", "SA")
strSE = FontINI.ReadINI("S", "SE")
strSHI = FontINI.ReadINI("S", "SHI")
strSO = FontINI.ReadINI("S", "SO")
strSU = FontINI.ReadINI("S", "SU")

strTA = FontINI.ReadINI("T", "TA")
strTE = FontINI.ReadINI("T", "TE")
strCHI = FontINI.ReadINI("T", "CHI")
strTO = FontINI.ReadINI("T", "TO")
strTSU = FontINI.ReadINI("T", "TSU")

strWA = FontINI.ReadINI("W", "WA")
strWO = FontINI.ReadINI("W", "WO")

strYA = FontINI.ReadINI("Y", "YA")
strYO = FontINI.ReadINI("Y", "YO")
strYU = FontINI.ReadINI("Y", "YU")
strSmallYA = FontINI.ReadINI("Y", "SmallYA")
strSmallYO = FontINI.ReadINI("Y", "SmallYO")
strSmallYU = FontINI.ReadINI("Y", "SmallYU")

strZA = FontINI.ReadINI("Z", "ZA")
strZE = FontINI.ReadINI("Z", "ZE")
strZI = FontINI.ReadINI("Z", "ZI")
strZO = FontINI.ReadINI("Z", "ZO")
strZU = FontINI.ReadINI("Z", "ZU")

lblWarning.Visible = False

Call Text1_Change

End Sub

Private Sub Form_Load()

Call Command2_Click

End Sub

Private Sub Text1_Change()

Dim strCur
Dim strPrev
Dim strNext
Dim strNextNext
Dim strLetter
Dim strNew

    For i = 1 To Len(Text1.Text)

        strCur = UCase(Mid(Text1.Text, i, 1))
        
        If i > 1 Then
            strPrev = UCase(Mid(Text1.Text, i - 1, 1))
        Else
            strPrev = ""
        End If
        
        If i < Len(Text1.Text) Then
            strNext = UCase(Mid(Text1.Text, i + 1, 1))
        Else
            strNext = ""
        End If
        
        If i < Len(Text1.Text) - 1 Then
            strNextNext = UCase(Mid(Text1.Text, i + 2, 1))
        Else
            strNextNext = ""
        End If
        
        If strCur = "B" Then
            If strNext = "Y" Then
                If strNextNext = "A" Then strLetter = "BYA": i = i + 2
                If strNextNext = "U" Then strLetter = "BYU": i = i + 2
                If strNextNext = "O" Then strLetter = "BYO": i = i + 2
            End If
            If strNext = "A" Then strLetter = "BA": i = i + 1
            If strNext = "E" Then strLetter = "BE": i = i + 1
            If strNext = "I" Then strLetter = "BI": i = i + 1
            If strNext = "O" Then strLetter = "BO": i = i + 1
            If strNext = "U" Then strLetter = "BU": i = i + 1
            If strNext = "B" Then strLetter = "Q"
        End If
        If strCur = "C" Then
            If strNext = "Y" Then
                If strNextNext = "A" Then strLetter = "KYA": i = i + 2
                If strNextNext = "U" Then strLetter = "KYU": i = i + 2
                If strNextNext = "O" Then strLetter = "KYO": i = i + 2
            End If
            If strNext = "H" Then
                If strNextNext = "A" Then strLetter = "CHA": i = i + 2
                If strNextNext = "E" Then strLetter = "CHE": i = i + 2
                If strNextNext = "I" Then strLetter = "CHI": i = i + 2
                If strNextNext = "O" Then strLetter = "CHO": i = i + 2
                If strNextNext = "U" Then strLetter = "CHU": i = i + 2
            End If
            If strNext = "A" Then strLetter = "KA": i = i + 1
            If strNext = "E" Then strLetter = "KE": i = i + 1
            If strNext = "I" Then strLetter = "KI": i = i + 1
            If strNext = "O" Then strLetter = "KO": i = i + 1
            If strNext = "U" Then strLetter = "KU": i = i + 1
            If strNext = "C" Then strLetter = "Q"
            If strNext = "K" Then strLetter = "Q"
        End If
        If strCur = "D" Then
            If strNext = "Y" Then
                If strNextNext = "A" Then strLetter = "DYA": i = i + 2
                If strNextNext = "E" Then strLetter = "DYE": i = i + 2
                If strNextNext = "I" Then strLetter = "DI": i = i + 2
                If strNextNext = "O" Then strLetter = "DYO": i = i + 2
                If strNextNext = "U" Then strLetter = "DYU": i = i + 2
            End If
            If strNext = "A" Then strLetter = "DA": i = i + 1
            If strNext = "E" Then strLetter = "DE": i = i + 1
            If strNext = "I" Then strLetter = "DI": i = i + 1
            If strNext = "O" Then strLetter = "DO": i = i + 1
            If strNext = "U" Then strLetter = "DU": i = i + 1
            If strNext = "D" Then strLetter = "Q"
        End If
        If strCur = "F" Then
            If strNext = "U" Then strLetter = "HU": i = i + 1
        End If
        If strCur = "G" Then
            If strNext = "Y" Then
                If strNextNext = "A" Then strLetter = "GYA": i = i + 2
                If strNextNext = "U" Then strLetter = "GYU": i = i + 2
                If strNextNext = "O" Then strLetter = "GYO": i = i + 2
            End If
            If strNext = "A" Then strLetter = "GA": i = i + 1
            If strNext = "E" Then strLetter = "GE": i = i + 1
            If strNext = "I" Then strLetter = "GI": i = i + 1
            If strNext = "O" Then strLetter = "GO": i = i + 1
            If strNext = "U" Then strLetter = "GU": i = i + 1
            If strNext = "G" Then strLetter = "Q"
        End If
        If strCur = "H" Then
            If strNext = "Y" Then
                If strNextNext = "A" Then strLetter = "HYA": i = i + 2
                If strNextNext = "U" Then strLetter = "HYU": i = i + 2
                If strNextNext = "O" Then strLetter = "HYO": i = i + 2
            End If
            If strNext = "A" Then strLetter = "HA": i = i + 1
            If strNext = "E" Then strLetter = "HE": i = i + 1
            If strNext = "I" Then strLetter = "HI": i = i + 1
            If strNext = "O" Then strLetter = "HO": i = i + 1
            If strNext = "U" Then strLetter = "HU": i = i + 1
            If strNext = "G" Then strLetter = "Q"
        End If
        If strCur = "J" Then
            If strNext = "A" Then strLetter = "ZYA": i = i + 1
            If strNext = "E" Then strLetter = "ZYE": i = i + 1
            If strNext = "I" Then strLetter = "ZI": i = i + 1
            If strNext = "O" Then strLetter = "ZYO": i = i + 1
            If strNext = "U" Then strLetter = "ZYU": i = i + 1
            If strNext = "J" Then strLetter = "Q"
        End If
        If strCur = "K" Then
            If strNext = "Y" Then
                If strNextNext = "A" Then strLetter = "KYA": i = i + 2
                If strNextNext = "U" Then strLetter = "KYU": i = i + 2
                If strNextNext = "O" Then strLetter = "KYO": i = i + 2
            End If
            If strNext = "A" Then strLetter = "KA": i = i + 1
            If strNext = "E" Then strLetter = "KE": i = i + 1
            If strNext = "I" Then strLetter = "KI": i = i + 1
            If strNext = "O" Then strLetter = "KO": i = i + 1
            If strNext = "U" Then strLetter = "KU": i = i + 1
            If strNext = "K" Then strLetter = "Q"
            If strNext = "C" Then strLetter = "Q"
        End If
        If strCur = "L" Or strCur = "R" Then
            If strNext = "Y" Then
                If strNextNext = "A" Then strLetter = "RYA": i = i + 2
                If strNextNext = "U" Then strLetter = "RYU": i = i + 2
                If strNextNext = "O" Then strLetter = "RYO": i = i + 2
            End If
            If strNext = "A" Then strLetter = "RA": i = i + 1
            If strNext = "E" Then strLetter = "RE": i = i + 1
            If strNext = "I" Then strLetter = "RI": i = i + 1
            If strNext = "O" Then strLetter = "RO": i = i + 1
            If strNext = "U" Then strLetter = "RU": i = i + 1
            If strNext = "R" Then strLetter = "Q"
            If strNext = "L" Then strLetter = "Q"
        End If
        If strCur = "M" Then
            If strNext = "Y" Then
                If strNextNext = "A" Then strLetter = "MYA": i = i + 2
                If strNextNext = "U" Then strLetter = "MYU": i = i + 2
                If strNextNext = "O" Then strLetter = "MYO": i = i + 2
            End If
            If strNext = "A" Then strLetter = "MA": i = i + 1
            If strNext = "E" Then strLetter = "ME": i = i + 1
            If strNext = "I" Then strLetter = "MI": i = i + 1
            If strNext = "O" Then strLetter = "MO": i = i + 1
            If strNext = "U" Then strLetter = "MU": i = i + 1
            If strNext = "M" Then strLetter = "Q"
            If strNext = "P" Then strLetter = "N"
            If strNext = "B" Then strLetter = "N"
        End If
        If strCur = "N" Then
            If strNext = "Y" Then
                If strNextNext = "A" Then strLetter = "NYA": i = i + 2
                If strNextNext = "U" Then strLetter = "NYU": i = i + 2
                If strNextNext = "O" Then strLetter = "NYO": i = i + 2
            End If
            If strNext = "A" Then strLetter = "NA": i = i + 1
            If strNext = "E" Then strLetter = "NE": i = i + 1
            If strNext = "I" Then strLetter = "NI": i = i + 1
            If strNext = "O" Then strLetter = "NO": i = i + 1
            If strNext = "U" Then strLetter = "NU": i = i + 1
            If strNext = "N" Then strLetter = "N": i = i + 1
            If strLetter = "" Then strLetter = "N"
        End If
        If strCur = "P" Then
            If strNext = "Y" Then
                If strNextNext = "A" Then strLetter = "PYA": i = i + 2
                If strNextNext = "U" Then strLetter = "PYU": i = i + 2
                If strNextNext = "O" Then strLetter = "PYO": i = i + 2
            End If
            If strNext = "A" Then strLetter = "PA": i = i + 1
            If strNext = "E" Then strLetter = "PE": i = i + 1
            If strNext = "I" Then strLetter = "PI": i = i + 1
            If strNext = "O" Then strLetter = "PO": i = i + 1
            If strNext = "U" Then strLetter = "PU": i = i + 1
            If strNext = "P" Then strLetter = "Q"
        End If
        If strCur = "Q" Then strLetter = "Q"
        If strCur = "S" Then
            If strNext = "H" Then
                If strNextNext = "A" Then strLetter = "SHA": i = i + 2
                If strNextNext = "E" Then strLetter = "SHE": i = i + 2
                If strNextNext = "I" Then strLetter = "SHI": i = i + 2
                If strNextNext = "O" Then strLetter = "SHO": i = i + 2
                If strNextNext = "U" Then strLetter = "SHU": i = i + 2
            End If
            If strNext = "Y" Then
                If strNextNext = "A" Then strLetter = "SHA": i = i + 2
                If strNextNext = "E" Then strLetter = "SHE": i = i + 2
                If strNextNext = "I" Then strLetter = "SHI": i = i + 2
                If strNextNext = "O" Then strLetter = "SHO": i = i + 2
                If strNextNext = "U" Then strLetter = "SHU": i = i + 2
            End If
            If strNext = "A" Then strLetter = "SA": i = i + 1
            If strNext = "E" Then strLetter = "SE": i = i + 1
            If strNext = "I" Then strLetter = "SHI": i = i + 1
            If strNext = "O" Then strLetter = "SO": i = i + 1
            If strNext = "U" Then strLetter = "SU": i = i + 1
            If strNext = "P" Then strLetter = "Q"
        End If
        If strCur = "T" Then
            If strNext = "Y" Then
                If strNextNext = "A" Then strLetter = "CHA": i = i + 2
                If strNextNext = "E" Then strLetter = "CHE": i = i + 2
                If strNextNext = "I" Then strLetter = "CHI": i = i + 2
                If strNextNext = "O" Then strLetter = "CHO": i = i + 2
                If strNextNext = "U" Then strLetter = "CHU": i = i + 2
            End If
            If strNext = "A" Then strLetter = "TA": i = i + 1
            If strNext = "E" Then strLetter = "TE": i = i + 1
            If strNext = "I" Then strLetter = "CHI": i = i + 1
            If strNext = "O" Then strLetter = "TO": i = i + 1
            If strNext = "U" Then strLetter = "TSU": i = i + 1
            If strNext = "S" And strNextNext = "U" Then strLetter = "TSU": i = i + 2
            If strNext = "T" Then strLetter = "Q"
            If strNext = "C" And strNextNext = "H" Then strLetter = "Q"
        End If
        If strCur = "W" Then
            If strNext = "A" Then strLetter = "WA": i = i + 1
            If strNext = "O" Then strLetter = "WO": i = i + 1
            If strNext = "WU" Then strLetter = "U": i = i + 1
            If strNext = "W" Then strLetter = "Q"
        End If
        If strCur = "Y" Then
            If strNext = "A" Then strLetter = "YA": i = i + 1
            If strNext = "O" Then strLetter = "YO": i = i + 1
            If strNext = "U" Then strLetter = "YU": i = i + 1
            If strNext = "Y" Then strLetter = "Q"
        End If
        If strCur = "Z" Then
            If strNext = "Y" Then
                If strNextNext = "A" Then strLetter = "ZYA": i = i + 2
                If strNextNext = "E" Then strLetter = "ZYE": i = i + 2
                If strNextNext = "I" Then strLetter = "ZI": i = i + 2
                If strNextNext = "O" Then strLetter = "ZYO": i = i + 2
                If strNextNext = "U" Then strLetter = "ZYU": i = i + 2
            End If
            If strNext = "A" Then strLetter = "ZA": i = i + 1
            If strNext = "E" Then strLetter = "ZE": i = i + 1
            If strNext = "I" Then strLetter = "ZI": i = i + 1
            If strNext = "O" Then strLetter = "ZO": i = i + 1
            If strNext = "U" Then strLetter = "ZU": i = i + 1
            If strNext = "Z" Then strLetter = "Q"
        End If
        If strCur = "A" Then
            If strNext = "A" Then strLetter = "A-": i = i + 1
            If strPrev = "A" Then strLetter = "-"
            If strLetter = "" Then strLetter = "A"
        End If
        If strCur = "E" Then
            If strNext = "E" Then strLetter = "E-": i = i + 1
            If strPrev = "E" Then strLetter = "-"
            If strLetter = "" Then strLetter = "E"
        End If
        If strCur = "I" Then
            If strNext = "I" Then strLetter = "I-": i = i + 1
            If strPrev = "I" Then strLetter = "i"
            If strLetter = "" Then strLetter = "I"
        End If
        If strCur = "O" Then
            If strNext = "O" Then strLetter = "O-": i = i + 1
            If strPrev = "O" Then strLetter = "-"
            If strLetter = "" Then strLetter = "O"
        End If
        If strCur = "U" Then
            If strNext = "U" Then strLetter = "U-": i = i + 1
            If strPrev = "U" Then strLetter = "-"
            If strLetter = "" Then strLetter = "U"
        End If
    If strCur = "-" Then strLetter = "-"
    If strCur = "1" Then strLetter = "1"
    If strCur = "2" Then strLetter = "2"
    If strCur = "3" Then strLetter = "3"
    If strCur = "4" Then strLetter = "4"
    If strCur = "5" Then strLetter = "5"
    If strCur = "6" Then strLetter = "6"
    If strCur = "7" Then strLetter = "7"
    If strCur = "8" Then strLetter = "8"
    If strCur = "9" Then strLetter = "9"
    If strCur = "0" Then strLetter = "0"
    If strCur = "." Then strLetter = "."
    If strCur = " " Then strLetter = " "

    If strLetter = "A" Then strLetter = strA
    If strLetter = "E" Then strLetter = strE
    If strLetter = "I" Then strLetter = strI
    If strLetter = "O" Then strLetter = strO
    If strLetter = "U" Then strLetter = strU
    If strLetter = "a" Then strLetter = strSmallA
    If strLetter = "e" Then strLetter = strSmallE
    If strLetter = "i" Then strLetter = strSmallI
    If strLetter = "o" Then strLetter = strSmallO
    If strLetter = "u" Then strLetter = strSmallU
    
    If strLetter = "A-" Then strLetter = strA & strLength
    If strLetter = "E-" Then strLetter = strE & strLength
    If strLetter = "I-" Then strLetter = strI & strSmallI
    If strLetter = "O-" Then strLetter = strO & strSmallU
    If strLetter = "U-" Then strLetter = strU & strLength
    
    If strLetter = "BA" Then strLetter = strBA
    If strLetter = "BE" Then strLetter = strBE
    If strLetter = "BI" Then strLetter = strBI
    If strLetter = "BO" Then strLetter = strBO
    If strLetter = "BU" Then strLetter = strBU
    If strLetter = "BYA" Then strLetter = strBI & strSmallYA
    If strLetter = "BYO" Then strLetter = strBI & strSmallYO
    If strLetter = "BYU" Then strLetter = strBI & strSmallYU

    If strLetter = "DA" Then strLetter = strDA
    If strLetter = "DE" Then strLetter = strDE
    If strLetter = "DI" Then strLetter = strDI
    If strLetter = "DO" Then strLetter = strDO
    If strLetter = "DU" Then strLetter = strDU
    If strLetter = "DYA" Then strLetter = strDI & strSmallYA
    If strLetter = "DYO" Then strLetter = strDI & strSmallYO
    If strLetter = "DYU" Then strLetter = strDI & strSmallYU
    If strLetter = "DYE" Then strLetter = strDI & strSmallE
    
    If strLetter = "GA" Then strLetter = strGA
    If strLetter = "GE" Then strLetter = strGE
    If strLetter = "GI" Then strLetter = strGI
    If strLetter = "GO" Then strLetter = strGO
    If strLetter = "GU" Then strLetter = strGU
    If strLetter = "GYA" Then strLetter = strGI & strSmallYA
    If strLetter = "GYO" Then strLetter = strGI & strSmallYO
    If strLetter = "GYU" Then strLetter = strGI & strSmallYU

    If strLetter = "HA" Then strLetter = strHA
    If strLetter = "HE" Then strLetter = strHE
    If strLetter = "HI" Then strLetter = strHI
    If strLetter = "HO" Then strLetter = strHO
    If strLetter = "HU" Then strLetter = strHU
    If strLetter = "HYA" Then strLetter = strHI & strSmallYA
    If strLetter = "HYO" Then strLetter = strHI & strSmallYO
    If strLetter = "HYU" Then strLetter = strHI & strSmallYU

    If strLetter = "KA" Then strLetter = strKA
    If strLetter = "KE" Then strLetter = strKE
    If strLetter = "KI" Then strLetter = strKI
    If strLetter = "KO" Then strLetter = strKO
    If strLetter = "KU" Then strLetter = strKU
    If strLetter = "KYA" Then strLetter = strKI & strSmallYA
    If strLetter = "KYO" Then strLetter = strKI & strSmallYO
    If strLetter = "KYU" Then strLetter = strKI & strSmallYU

    If strLetter = "MA" Then strLetter = strMA
    If strLetter = "ME" Then strLetter = strME
    If strLetter = "MI" Then strLetter = strMI
    If strLetter = "MO" Then strLetter = strMO
    If strLetter = "MU" Then strLetter = strMU
    If strLetter = "MYA" Then strLetter = strMI & strSmallYA
    If strLetter = "MYO" Then strLetter = strMI & strSmallYO
    If strLetter = "MYU" Then strLetter = strMI & strSmallYU
    
    If strLetter = "NA" Then strLetter = strNA
    If strLetter = "NE" Then strLetter = strNE
    If strLetter = "NI" Then strLetter = strNI
    If strLetter = "NO" Then strLetter = strNO
    If strLetter = "NU" Then strLetter = strNU
    If strLetter = "NYA" Then strLetter = strNI & strSmallYA
    If strLetter = "NYO" Then strLetter = strNI & strSmallYO
    If strLetter = "NYU" Then strLetter = strNI & strSmallYU
    If strLetter = "N" Then strLetter = strN
    
    If strLetter = "PA" Then strLetter = strPA
    If strLetter = "PE" Then strLetter = strPE
    If strLetter = "PI" Then strLetter = strPI
    If strLetter = "PO" Then strLetter = strPO
    If strLetter = "PU" Then strLetter = strPU
    If strLetter = "PYA" Then strLetter = strPI & strSmallYA
    If strLetter = "PYO" Then strLetter = strPI & strSmallYO
    If strLetter = "PYU" Then strLetter = strPI & strSmallYU
    
    If strLetter = "Q" Then strLetter = strSmallTSU
    
    If strLetter = "RA" Then strLetter = strRA
    If strLetter = "RE" Then strLetter = strRE
    If strLetter = "RI" Then strLetter = strRI
    If strLetter = "RO" Then strLetter = strRO
    If strLetter = "RU" Then strLetter = strRU
    If strLetter = "RYA" Then strLetter = strRI & strSmallYA
    If strLetter = "RYO" Then strLetter = strRI & strSmallYO
    If strLetter = "RYU" Then strLetter = strRI & strSmallYU
    
    If strLetter = "SA" Then strLetter = strSA
    If strLetter = "SE" Then strLetter = strSE
    If strLetter = "SO" Then strLetter = strSO
    If strLetter = "SU" Then strLetter = strSU
    
    If strLetter = "SHA" Then strLetter = strSHI & strSmallYA
    If strLetter = "SHE" Then strLetter = strSHI & strSmallE
    If strLetter = "SHI" Then strLetter = strSHI
    If strLetter = "SHO" Then strLetter = strSHI & strSmallYO
    If strLetter = "SHU" Then strLetter = strSHI & strSmallYU
    
    If strLetter = "TA" Then strLetter = strTA
    If strLetter = "TE" Then strLetter = strTE
    If strLetter = "TO" Then strLetter = strTO
    If strLetter = "TSU" Then strLetter = strTSU
    
    If strLetter = "CHA" Then strLetter = strCHI & strSmallYA
    If strLetter = "CHE" Then strLetter = strCHI & strSmallE
    If strLetter = "CHI" Then strLetter = strCHI
    If strLetter = "CHO" Then strLetter = strCHI & strSmallYO
    If strLetter = "CHU" Then strLetter = strCHI & strSmallYU
    
    If strLetter = "WA" Then strLetter = strWA
    If strLetter = "WO" Then strLetter = strWO
    
    If strLetter = "YA" Then strLetter = strYA
    If strLetter = "YO" Then strLetter = strYO
    If strLetter = "YU" Then strLetter = strYU
    
    If strLetter = "ZA" Then strLetter = strZA
    If strLetter = "ZE" Then strLetter = strZE
    If strLetter = "ZI" Then strLetter = strZI
    If strLetter = "ZO" Then strLetter = strZO
    If strLetter = "ZU" Then strLetter = strZU
    If strLetter = "ZYA" Then strLetter = strZI & strSmallYA
    If strLetter = "ZYO" Then strLetter = strZI & strSmallYO
    If strLetter = "ZYU" Then strLetter = strZI & strSmallYU
    If strLetter = "ZYE" Then strLetter = strZI & strSmallE

    If strLetter = "-" Then strLetter = strLength

    If strLetter = "1" Then strLetter = str1
    If strLetter = "2" Then strLetter = str2
    If strLetter = "3" Then strLetter = str3
    If strLetter = "4" Then strLetter = str4
    If strLetter = "5" Then strLetter = str5
    If strLetter = "6" Then strLetter = str6
    If strLetter = "7" Then strLetter = str7
    If strLetter = "8" Then strLetter = str8
    If strLetter = "9" Then strLetter = str9
    If strLetter = "0" Then strLetter = str0
    If strLetter = "." Then strLetter = strDot
    
    strNew = strNew & strLetter
    strLetter = ""
    
    Next

Text2.Text = strNew

End Sub
