VERSION 5.00
Begin VB.Form frmPrimos 
   BackColor       =   &H00FFC0C0&
   Caption         =   "NUMEROS PRIMOS"
   ClientHeight    =   9930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   ScaleHeight     =   9930
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstCantidadPrimos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9030
      Left            =   4320
      TabIndex        =   14
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox txtPorcentajePrimos 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   5040
      Width           =   1215
   End
   Begin VB.ListBox lstConteo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9030
      Left            =   9600
      TabIndex        =   10
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox txtAproxPrimos 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtTotalPrimos 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ListBox lstResta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9030
      Left            =   6960
      TabIndex        =   4
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox txtMeta 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Text            =   "2"
      Top             =   720
      Width           =   1215
   End
   Begin VB.ListBox lstPrimos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9030
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton cmdCalculaPrimo 
      Caption         =   "Calcula"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Cantidad de Primos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   15
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "% Primos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Conteo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   11
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Aprox Primos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Total Primos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "1 - Primos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   5
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Numeros Primos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmPrimos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'* PROYECTO      : CALCULA PRIMOS
'* CONTENIDO     : PERMITE CALCULAR NÚMEROS PRIMOS
'* VERSION       : 1.1
'* AUTORES       : MIGUEL QUINTEIRO PIÑERO / MIGUEL QUINTEIRO FERNANDEZ
'* INICIO        : 16 DE JUNIO DE 2013
'* ACTUALIZACION : 16 DE JUNIO DE 2013
'****************************************************************************************
Option Explicit

' Declaracion de variables
Dim miNumero As Double
Dim miRaizEntera As Double
Dim miResto As Double
Dim miPrueba As Double
Dim miDivisor As Double
Dim miIndice As Double
Dim miTotalPrimos As Double
Dim miAproxPrimos As Double
Dim miConteo As Double
Dim noesprimo As Boolean
Dim miAcumulaNoPrimo As Long
Dim miAcumulaPrimo As Long


' Calcular numeros primos
Private Sub cmdCalculaPrimo_Click()

'Abre archivo de salida
  Open "Salida.txt" For Output As #1
  Open "Conjetura.txt" For Output As #2
  Open "Primos.txt" For Output As #3

  ' Limpiar los listbox
  lstPrimos.Clear
  lstCantidadPrimos.Clear
  lstResta.Clear
  lstConteo.Clear

  ' Almacenar los primeros primos
  lstPrimos.AddItem 2
  Print #3, "2"
  lstCantidadPrimos.AddItem 1
  lstResta.AddItem (1 - 2)
  lstConteo.AddItem 0
  Print #1, "0"
  lstPrimos.AddItem 3
  Print #3, "3"
  lstCantidadPrimos.AddItem 2
  lstResta.AddItem (1 - 3)
  lstConteo.AddItem 0
  Print #1, "0"
  Print #1, "1"

  Print #2, "1", "0"
  Print #2, "2", "1"
  Print #2, "3", "2"
  Print #2, "4", "2"



  ' Inicializar las variables
  miNumero = Val(txtMeta)
  miTotalPrimos = 2
  miConteo = 0

  miAcumulaNoPrimo = 1
  miAcumulaPrimo = 5

  ' Ciclo de busqueda
  For miPrueba = 5 To miNumero Step 2
    noesprimo = True
    miRaizEntera = Int(Sqr(miPrueba))
    For miIndice = 1 To (lstPrimos.ListCount - 1)
      lstPrimos.ListIndex = miIndice
      miDivisor = Val(lstPrimos.Text)
      If miDivisor > miRaizEntera Then
        ' Si es primo
        miTotalPrimos = miTotalPrimos + 1
        miConteo = 0
        lstPrimos.AddItem miPrueba
        Print #3, miPrueba

        lstCantidadPrimos.AddItem miTotalPrimos
        lstResta.AddItem (1 - miPrueba)
        lstConteo.AddItem miConteo
        Print #1, miConteo
        miIndice = lstPrimos.ListCount - 1
        noesprimo = False
      Else
        ' Si no es primo
        miResto = miPrueba Mod miDivisor
        If miResto = 0 Then
          miIndice = lstPrimos.ListCount - 1
        End If
      End If
    Next miIndice

    If noesprimo Then
      miAcumulaNoPrimo = miAcumulaNoPrimo + miPrueba + (miPrueba - 1)
      miConteo = miConteo + 1
      lstConteo.AddItem miConteo
      Print #1, miConteo
    Else
      miAcumulaPrimo = miAcumulaPrimo + miPrueba
    End If

    Print #2, miPrueba, miTotalPrimos
    Print #2, miPrueba + 1, miTotalPrimos
  Next miPrueba

  ' Otras operaciones
  miAproxPrimos = miNumero / Log(miNumero)

  ' Muestra los resultados
  txtTotalPrimos.Text = miTotalPrimos
  txtAproxPrimos.Text = Format(miAproxPrimos, "0.00")
  txtPorcentajePrimos.Text = miTotalPrimos * 100 / Val(txtMeta.Text)
  'txtAcumulaNoPrimo.Text = miAcumulaNoPrimo
  'txtAcumulaPrimo.Text = miAcumulaPrimo

  ' Cierrra el archivo
  Close #1
  Close #2
  Close #3
End Sub

