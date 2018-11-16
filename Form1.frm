VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1260
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   1260
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHello 
      Height          =   555
      Left            =   1650
      TabIndex        =   2
      Top             =   60
      Width           =   2205
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create PDF"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   690
      Width           =   1215
   End
   Begin VB.Label lblHello 
      Alignment       =   1  'Right Justify
      Caption         =   "Hellow World:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   210
      Width           =   1485
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    ' Create a simple PDF file using the mjwPDF class
    Dim objPDF As New mjwPDF
    
    ' Set the PDF title and filename
    objPDF.PDFTitle = "Test PDF Document"
    objPDF.PDFFileName = "c:\Temp\testABC.pdf"

    ' We must tell the class where the PDF fonts are located
    objPDF.PDFLoadAfm = App.Path & "\Temp\Fonts"

    ' View the PDF file after we create it
    objPDF.PDFView = True

    ' Begin our PDF document
    objPDF.PDFBeginDoc
        ' Set the font name, size, and style
        objPDF.PDFSetFont FONT_ARIAL, 15, FONT_BOLD

        ' Set the text color
        objPDF.PDFSetTextColor = vbBlue

        ' Set the text we want to print
        objPDF.PDFTextOut "Hello, World! From www.vb6.us"

    ' End our PDF document (this will save it to the filename)
    objPDF.PDFEndDoc
    
    objPDF.PDFFileName = "c:\Temp\test.pdf"
    
    txtHello.Text = "Hello World"
    lblHello.Caption = "Happy Live!!!"
    
    MsgBox "Final project and merge to master folder.", vbInformation + vbOKOnly, Me.Caption
    
    txtHello.Text = "Finish Work."
    lblHello.Caption = "Are you sure???"
End Sub
