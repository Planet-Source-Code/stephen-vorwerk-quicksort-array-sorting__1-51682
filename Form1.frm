VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00D0C0AC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QuickSorting"
   ClientHeight    =   7590
   ClientLeft      =   4215
   ClientTop       =   1080
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   7050
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00797453&
      Caption         =   "&Clear Text Box"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00797453&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdAverage 
      BackColor       =   &H00797453&
      Caption         =   "&Average and Reset "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7120
      Width           =   2175
   End
   Begin VB.TextBox txtTime2 
      BackColor       =   &H00EDEADE&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2895
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   4200
      Width           =   6975
   End
   Begin VB.OptionButton opt10000 
      BackColor       =   &H00BDA688&
      Caption         =   "10000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox txtTime 
      BackColor       =   &H00EDEADE&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5040
      TabIndex        =   3
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00797453&
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   1455
   End
   Begin VB.ListBox lstNumbers 
      BackColor       =   &H00BDA688&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   2085
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.ListBox lstNumbers 
      BackColor       =   &H00BDA688&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   2085
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   2295
   End
   Begin VB.OptionButton opt1000 
      BackColor       =   &H00BDA688&
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   5
      Top             =   2640
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.OptionButton opt30000 
      BackColor       =   &H00BDA688&
      Caption         =   "30000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   6855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Numbers to test:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4980
      TabIndex        =   12
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Time of test:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4980
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00DACDBE&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   4920
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00BDA688&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   4920
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00DACDBE&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   4920
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00BDA688&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Left            =   4920
      Top             =   2280
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Stephen Vorwerk
'Bubble Sort Sort-Time Analyzing

'Controls:
'   lstNumbers(0)   : This ListBox holds the original randomly generated numbers
'   lstNumbers(1)   : This ListBox holds the sorted numbers after using Bubble Sort
'   txtTime         : This TextBox displays the time for the last sorted array.
'   txtTime2        : This TextBox displays all the times run so far, as well as
'   >>>>>>>>>>>>>>>>:      the average time taken per set of tests.
'   opt1000         : This OptionButton chooses 1000 numbers to be sorted.
'   opt10000        : This OptionButton chooses 10000 numbers to be sorted.
'   opt30000        : This OptionButton chooses 30000 numbers to be sorted.

Option Explicit

'   The Global Variable Tests(Single) is used to keep track of how many
'   times the user has tested the time in each set of tests.
'   The Global Variable Sum(Single) is used to hold the sum of all the
'   times taken per test per set.

Dim Tests As Integer, Sum As Single

Private Sub cmdAverage_Click()
    Dim Average As Single
    
'   Finds the average time for the set and displays it.
    
    Average = Sum / Tests
    txtTime2.Text = txtTime2.Text & "Average time for " & Tests & _
                    " sort tests: " & Str$(Average) & vbCrLf
    txtTime2.SelStart = Len(txtTime2.Text)
    
'   Resets the number of Tests and the Sum for a new set.

    Tests = 0
    Sum = 0
End Sub

Private Sub cmdClear_Click()
'   Clears txtTime2 of all it's information.
    txtTime2.Text = ""
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdStart_Click()

'   N() is the dynamic array, re-dimensioned depending on the
'   user-preference. Amount is the UBound of N(), the highest
'   index number. Start is used to calculate the time of the
'   sort, holding the starting time. TotalTime is the Start
'   time minus the time when the sorting is complete to
'   calculate the time the sort took. ChangeMade is a boolean
'   used in the Bubble Sort technique to make it stop sorting.

    Dim N() As Single, Amount As Long, I As Integer
    Dim Start As Single, TotalTime As Single
    
'   Set the variables for the new test and clear previous info.
    Tests = Tests + 1
    
    lstNumbers(0).Clear
    lstNumbers(1).Clear
    
'   Find and ReDim N() to the amount of numbers to be sorted.
'   Then make Amount the highest index of N() and fill N()
'   with randomly generated numbers
    
    If opt1000.Value = True Then ReDim N(1 To 1000)
    If opt10000.Value = True Then ReDim N(1 To 10000)
    If opt30000.Value = True Then ReDim N(1 To 30000)
    Amount = UBound(N)
    
    For I = 1 To Amount
        Randomize
        N(I) = Int(Rnd * 1000000) + 1
        lstNumbers(0).AddItem N(I)
    Next I
    
'   Start the sort and keep sorting until no change
'   was made to the array.
    
    Start = Timer
    QuickSort N(), Amount
    
'   Calculate the time taken for the sort and the new sum.
'   Then display the time of the sort in the txtTime and txtTime2.
'   Set the SelStart of txtTime to the last character so that
'   the user doesnt have to scroll down the text box to see the time.
    
    TotalTime = Timer - Start
    Sum = Sum + TotalTime
    txtTime.Text = Str$(TotalTime) & "s"
    txtTime2.Text = txtTime2.Text & "Test #" & LTrim$(Str$(Tests)) & "(" & _
        LTrim$(Str$(Amount)) & " numbers):" & Str$(TotalTime) & _
        "s" & vbCrLf
    txtTime2.SelStart = Len(txtTime2.Text)
    
'   Display the newly sorted array in lstNumbers(1)

    For I = 1 To Amount
        lstNumbers(1).AddItem N(I)
    Next I
End Sub
Private Sub QuickSort(ByRef ArrayName As Variant, ByVal Size As Single)
    Dim L As Single, R As Single, I As Single, LeftEnd As Single, RightEnd As Single
    
    LeftEnd = 1
    RightEnd = Size
    Dim S() As Single, SPtr As Single
    ReDim S(1 To 100) As Single
    SPtr = 0
    Push S(), SPtr, 1
    Push S(), SPtr, Size
    
    While Not EmptyStack(SPtr)
        Pop S(), SPtr, RightEnd
        Pop S(), SPtr, LeftEnd
        
        L = LeftEnd + 1
        R = RightEnd
        SWAP ArrayName(LeftEnd), ArrayName((RightEnd - LeftEnd) \ 2 + LeftEnd)
        
        While L < R
            While L < R And ArrayName(L) <= ArrayName(LeftEnd)
                L = L + 1
            Wend
            While L < R And ArrayName(R) > ArrayName(LeftEnd)
                R = R - 1
            Wend
            SWAP ArrayName(L), ArrayName(R)
        Wend
        
        If ArrayName(L) > ArrayName(LeftEnd) Then L = L - 1
        SWAP ArrayName(LeftEnd), ArrayName(L)
        
        If L - LeftEnd > 1 Then
            Push S(), SPtr, LeftEnd
            Push S(), SPtr, L - 1
        End If
        If RightEnd - L > 1 Then
            Push S(), SPtr, L + 1
            Push S(), SPtr, RightEnd
        End If
    Wend

End Sub

Private Function EmptyStack(ByVal StackPointer As Single) As Boolean
    If StackPointer > 0 Then EmptyStack = False Else EmptyStack = True
End Function

Private Sub Push(ByRef Stack As Variant, ByRef StackPointer As Single, _
                ByVal StackItem As Single)
    If UBound(Stack) < StackPointer + 1 Then ReDim Preserve Stack(1 To StackPointer + 10)
    StackPointer = StackPointer + 1
    Stack(StackPointer) = StackItem
End Sub

Private Sub Pop(ByRef Stack As Variant, ByRef StackPointer As Single, _
               ByRef StackItem As Single)
    If Not EmptyStack(StackPointer) Then
        StackItem = Stack(StackPointer)
        StackPointer = StackPointer - 1
    End If
End Sub

Private Sub SWAP(A As Variant, B As Variant)
    Dim T As Variant
    T = A
    A = B
    B = T
End Sub

Private Sub Form_Load()
'   Set the default values

    Tests = 0
    Sum = 0
End Sub

