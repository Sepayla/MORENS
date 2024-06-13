VERSION 5.00
Begin VB.Form PayrollDash 
   Caption         =   "Form1"
   ClientHeight    =   9480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15330
   LinkTopic       =   "Form1"
   ScaleHeight     =   9480
   ScaleWidth      =   15330
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   8535
      Left            =   12360
      TabIndex        =   49
      Top             =   360
      Width           =   2295
      Begin VB.CommandButton closebttn 
         Caption         =   "Close"
         Height          =   975
         Left            =   480
         Picture         =   "PayrollDash.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   6960
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Print"
         Height          =   975
         Left            =   480
         Picture         =   "PayrollDash.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   5640
         Width           =   1335
      End
      Begin VB.CommandButton findbttn 
         Caption         =   "Find"
         Height          =   975
         Left            =   480
         Picture         =   "PayrollDash.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton deletebttn 
         Caption         =   "Delete"
         Height          =   975
         Left            =   480
         Picture         =   "PayrollDash.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton savebttn 
         Caption         =   "Save"
         Height          =   975
         Left            =   480
         Picture         =   "PayrollDash.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton addbttn 
         Caption         =   "Add"
         Height          =   975
         Left            =   480
         Picture         =   "PayrollDash.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "List of Deductions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4215
      Left            =   6360
      TabIndex        =   36
      Top             =   4680
      Width           =   5415
      Begin VB.CommandButton Command3 
         Caption         =   "Compute Deductions"
         Height          =   615
         Left            =   240
         TabIndex        =   48
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Net Income"
         Height          =   615
         Left            =   240
         TabIndex        =   47
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox txtsss1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   42
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txttax 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   41
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtphil 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   40
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtpag 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   39
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox txtnetincome 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2280
         TabIndex        =   38
         Top             =   3480
         Width           =   2655
      End
      Begin VB.TextBox txttotdeduction 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   37
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label Label22 
         Caption         =   "SSS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "Tax with Held  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "PhilHealth:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "Pag-Ibig"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   1920
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Breakdown of Wages"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4215
      Left            =   480
      TabIndex        =   24
      Top             =   4680
      Width           =   5415
      Begin VB.CommandButton cmdGross 
         Caption         =   "Compute Gross Pay"
         Height          =   495
         Left            =   2280
         MaskColor       =   &H8000000F&
         TabIndex        =   35
         Top             =   3480
         Width           =   2655
      End
      Begin VB.TextBox txtGrossPay 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   29
         Top             =   3000
         Width           =   2655
      End
      Begin VB.TextBox txtmeal 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   28
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox txtperhour 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   27
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox txtperday 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   26
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txt15th 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   25
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label16 
         Caption         =   "Gross Pay:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label15 
         Caption         =   "Allowance:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "Rate per Hour:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Rate per Day:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Rate per 15th day of the Month"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   30
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Deduction Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3495
      Left            =   6360
      TabIndex        =   13
      Top             =   360
      Width           =   5415
      Begin VB.TextBox txtdatehired 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   18
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtSSS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   17
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txttin 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   16
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtphilhealth 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   15
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox txtpagibig 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   14
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label12 
         Caption         =   "Date Hired:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "SSS#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "TIN#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "PhilHealth"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Pag-Ibig"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   2880
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transaction Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   5415
      Begin VB.TextBox txtdateTo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   12
         Top             =   2880
         Width           =   2655
      End
      Begin VB.TextBox txtdatefrom 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   11
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox txtmonthlysalary 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   10
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox txtname 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   9
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtempID 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txttranno 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Date Covered To:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Date Covered From:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Monthly Salary:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Complete Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Employee ID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Transaction No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "PayrollDash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub addbttn_Click()
On Error Resume Next

txttranno.SelStart = 0
txttranno.SelLength = Len(txttranno.Text)
txttranno.SetFocus

txttranno.Text = ""
txtempID.Text = ""
txtname.Text = ""
txtmonthlysalary.Text = ""
txtdatefrom.Text = ""
txtdateTo.Text = ""
txtdatehired.Text = ""
txtSSS.Text = ""
txttin.Text = ""
txtphilhealth.Text = ""
txtpagibig.Text = ""
txt15th.Text = ""
txtperday.Text = ""
txtperhour.Text = ""
txtmeal.Text = ""
txtGrossPay.Text = ""
txtsss1.Text = ""
txttax.Text = ""
txtphil.Text = ""
txtpag.Text = ""
txttotdeduction.Text = ""
txtnetincome.Text = ""
End Sub

Private Sub closebttn_Click()
Unload Me
End Sub

Private Sub cmdGross_Click()
Dim xrate15 As Double
Dim xSalary As Double
Dim xrateperday As Double
Dim xrateperhour As Double
Dim xmeal As Double
Dim xGross As Double

xSalary = txtmonthlysalary.Text
xrate15 = xSalary / 2
txt15th.Text = xrate15

xrateperday = txtmonthlysalary.Text / 26
txtperday.Text = xrateperday

xrateperhour = txtperday.Text / 8
txtperhour.Text = xrateperhour

xmeal = 500

txtmeal.Text = xmeal

xGross = xmeal + xrate15

txtGrossPay.Text = xGross
End Sub

Private Sub findbttn_Click()
txttranno.SelStart = 0
txttranno.SelLength = Len(txttranno.Text)
txttranno.SetFocus

openrstPayroll "SELECT * FROM payroll WHERE tranno='" & Trim(txttranno.Text) & "'"
    If Not rstPayroll.EOF Then
        With rstPayroll
            txttranno.Text = .Fields("tranno").Value
            txtempID.Text = .Fields("employeeid").Value
            txtdatefrom.Text = .Fields("datefrom").Value
            txtdateTo.Text = .Fields("dateto").Value
            txt15th.Text = .Fields("rate15").Value
            txtperday.Text = .Fields("rateperday").Value
            txtperhour.Text = .Fields("rateperhour").Value
            txtmeal.Text = .Fields("meal").Value
            txtGrossPay.Text = .Fields("grosspay").Value
            txtdatehired.Text = .Fields("datehired").Value
            txtSSS.Text = .Fields("sssno").Value
            txttin.Text = .Fields("tinno").Value
            txtphilhealth.Text = .Fields("philhealthno").Value
            txtpagibig.Text = .Fields("pagibigno").Value
            txtsss1.Text = .Fields("sss").Value
            txttax.Text = .Fields("tax").Value
            txtpag.Text = .Fields("pagibig").Value
            txtphil.Text = .Fields("philhealth").Value
            txttotdeduction.Text = .Fields("totaldeduction").Value
            txtnetincome.Text = .Fields("netincome").Value
        End With
    Else
        ' Record not found, notify the user
        MsgBox "Transaction number not found."
    End If
End Sub

Private Sub Form_Load()
openWORKSPACEODBC
openconPayroll
End Sub

Private Sub savebttn_Click()
openrstPayroll "SELECT * FROM payroll where tranno='" & Trim(txttranno.Text) & "'"
If Not rstPayroll.EOF Then
'if not found
    With rstPayroll
        .Edit
            .Fields("tranno").Value = Trim(txttranno.Text)
            .Fields("employeeid").Value = Trim(txtempID.Text)
            .Fields("datefrom").Value = Trim(txtdatefrom.Text)
            .Fields("dateto").Value = Trim(txtdateTo.Text)
            .Fields("rate15").Value = Trim(txt15th.Text)
            .Fields("rateperday").Value = Trim(txtperday.Text)
            .Fields("rateperhour").Value = Trim(txtperhour.Text)
            .Fields("meal").Value = Trim(txtmeal.Text)
            .Fields("grosspay").Value = Trim(txtGrossPay.Text)
            .Fields("datehired").Value = Trim(txtdatehired.Text)
            .Fields("sssno").Value = Trim(txtSSS.Text)
            .Fields("tinno").Value = Trim(txttin.Text)
            .Fields("philhealthno").Value = Trim(txtphilhealth.Text)
            .Fields("pagibigno").Value = Trim(txtpagibig.Text)
            .Fields("sss").Value = Trim(txtsss1.Text)
            .Fields("tax").Value = Trim(txttax.Text)
            .Fields("pagibig").Value = Trim(txtpag.Text)
            .Fields("philhealth").Value = Trim(txtphil.Text)
            .Fields("totaldeduction").Value = Trim(txttotdeduction.Text)
            .Fields("netincome").Value = Trim(txtnetincome.Text)
            
        .Update
        
    End With
Else
    'not found
        With rstPayroll
            .AddNew
                .Fields("tranno").Value = Trim(txttranno.Text)
                .Fields("employeeid").Value = Trim(txtempID.Text)
                .Fields("datefrom").Value = Trim(txtdatefrom.Text)
                .Fields("dateto").Value = Trim(txtdateTo.Text)
                .Fields("rate15").Value = Trim(txt15th.Text)
                .Fields("rateperday").Value = Trim(txtperday.Text)
                .Fields("rateperhour").Value = Trim(txtperhour.Text)
                .Fields("meal").Value = Trim(txtmeal.Text)
                .Fields("grosspay").Value = Trim(txtGrossPay.Text)
                .Fields("datehired").Value = Trim(txtdatehired.Text)
                .Fields("sssno").Value = Trim(txtSSS.Text)
                .Fields("tinno").Value = Trim(txttin.Text)
                .Fields("philhealthno").Value = Trim(txtphilhealth.Text)
                .Fields("pagibigno").Value = Trim(txtpagibig.Text)
                .Fields("sss").Value = Trim(txtsss1.Text)
                .Fields("tax").Value = Trim(txttax.Text)
                .Fields("pagibig").Value = Trim(txtpag.Text)
                .Fields("philhealth").Value = Trim(txtphil.Text)
                .Fields("totaldeduction").Value = Trim(txttotdeduction.Text)
                .Fields("netincome").Value = Trim(txtnetincome.Text)
            .Update
            
        End With
End If
End Sub
