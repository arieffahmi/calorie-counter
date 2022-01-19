VERSION 5.00
Begin VB.Form calorieCounter 
   Caption         =   "calorieCounter"
   ClientHeight    =   10275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   ScaleHeight     =   10275
   ScaleWidth      =   15960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   495
      Left            =   7680
      TabIndex        =   23
      Top             =   9360
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   9600
      TabIndex        =   22
      Top             =   9360
      Width           =   1695
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   5760
      TabIndex        =   21
      Top             =   9360
      Width           =   1695
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   3840
      TabIndex        =   20
      Top             =   9360
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Summary"
      Height          =   2055
      Left            =   3720
      TabIndex        =   15
      Top             =   6960
      Width           =   7575
      Begin VB.Label lblSum 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3480
         TabIndex        =   25
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label lblCount 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3480
         TabIndex        =   24
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label lblTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3480
         TabIndex        =   19
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label10 
         Caption         =   "Count For Item Entered :"
         Height          =   375
         Left            =   960
         TabIndex        =   18
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label9 
         Caption         =   "Total Of Calories : "
         Height          =   375
         Left            =   960
         TabIndex        =   17
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label8 
         Caption         =   "Sum of accumulated calories : "
         Height          =   375
         Left            =   960
         TabIndex        =   16
         Top             =   1440
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Calories"
      Height          =   2295
      Left            =   3720
      TabIndex        =   8
      Top             =   4320
      Width           =   7575
      Begin VB.Label lblProtein1 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3480
         TabIndex        =   14
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label lblCarbohydrate1 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3480
         TabIndex        =   13
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label lblFat1 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3480
         TabIndex        =   12
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label7 
         Caption         =   "4 per grams Protein :"
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "4 per grams Carbohydrate : "
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "9 per grams Fats : "
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   3720
      TabIndex        =   1
      Top             =   1800
      Width           =   7575
      Begin VB.TextBox txtProtein 
         Height          =   375
         Left            =   3480
         TabIndex        =   7
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox txtCarbohydrate 
         Height          =   375
         Left            =   3480
         TabIndex        =   6
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtFat 
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label4 
         Caption         =   "Enter Your Protein(gram) : "
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Enter Your Carbohydrate(gram):"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Enter Your Fat(gram) : "
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "CALORIE COUNTER"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   0
      Top             =   600
      Width           =   7335
   End
End
Attribute VB_Name = "calorieCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const caloriesFat               As Integer = 9
Const caloriesCarbohydrate      As Integer = 4
Const caloriesProtein           As Integer = 4
Dim intSum                      As Integer
Dim intCount                    As Integer


Private Sub cmdCalculate_Click()

    Dim curFat                  As Currency
    Dim curCarbohydrate         As Currency
    Dim curProtein              As Currency
    Dim curFat1                 As Currency
    Dim curCarbohydrate1        As Currency
    Dim curProtein1             As Currency
    Dim curTotal                As Currency
    
    'convert input values to numeric number
    curFat = Val(txtFat.Text)
    curCarbohydrate = Val(txtCarbohydrate.Text)
    curProtein = Val(txtProtein.Text)
    
    'calculate values for each calories
    curFat1 = curFat * caloriesFat
    curCarbohydrate1 = curCarbohydrate * caloriesCarbohydrate
    curProtein1 = curProtein * caloriesProtein
    
    'Format and display answer calories
    lblFat1.Caption = FormatNumber(curFat1, 0)
    lblCarbohydrate1.Caption = FormatNumber(curCarbohydrate1, 0)
    lblProtein1.Caption = FormatNumber(curProtein1, 0)
    
    'calculate total calories
    curTotal = curFat1 + curCarbohydrate1 + curProtein1
    
    'Format and display summary values
    lblTotal.Caption = FormatNumber(curTotal, 0)
    
    intCount = intCount + 1
    lblCount.Caption = FormatNumber(intCount, 0)
    
    intSum = intSum + curTotal
    lblSum.Caption = FormatNumber(intSum, 0)
    
    
    
    
End Sub

Private Sub cmdClear_Click()
    'clear previous amounts from form
    txtFat.Text = ""
    txtCarbohydrate.Text = ""
    txtProtein.Text = ""
    lblFat1.Caption = ""
    lblCarbohydrate1.Caption = ""
    lblProtein1.Caption = ""
    lblTotal.Caption = ""
    txtFat.SetFocus
End Sub

Private Sub cmdExit_Click()
    'Exit the Project
    End
End Sub

Private Sub cmdPrint_Click()
    'Print the Form
    PrintForm
End Sub

