VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   1575
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   75
      TabIndex        =   0
      Top             =   900
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type WAVEFORMATEX
    wFormatTag                  As Integer
    nChannels                   As Integer
    nSamplesPerSec              As Long
    nAvgBytesPerSec             As Long
    nBlockAlign                 As Integer
    wBitsPerSample              As Integer
    cbSize                      As Integer
End Type

Private WithEvents cWavConcat As clsWavConcatenate
Attribute cWavConcat.VB_VarHelpID = -1

Private Sub cWavConcat_FileChanged(file As String)
    Label2 = file
End Sub

Private Sub cWavConcat_Progress(percent As Integer)
    Label1 = percent & "%"
End Sub

Private Sub Form_Load()

    Set cWavConcat = New clsWavConcatenate
    Dim wfx          As WAVEFORMATEX

    Me.Show

    With wfx
        .cbSize = 12                    ' standard
        .nAvgBytesPerSec = 44100 * 4    ' bytes/s
        .nBlockAlign = 4                ' block align
        .nChannels = 2                  ' channels
        .nSamplesPerSec = 44100         ' samples/s
        .wBitsPerSample = 16            ' bits/sample
        .wFormatTag = 1                 ' PCM
    End With

    cWavConcat.OutputFormat = VarPtr(wfx)

    ' add wavs
    If Not cWavConcat.WaveAdd("C:\Windows\Media\tada.wav") Then
        MsgBox "Wave #1 not supported!"
    End If

    If Not cWavConcat.WaveAdd("C:\WINDOWS\Media\notify.wav") Then
        MsgBox "Wave #2 not supported!"
    End If

    ' do the job
    If Not cWavConcat.WaveConcatenate("C:\cool.wav") Then
        MsgBox "Failed!", vbExclamation, "Error"
    Else
        MsgBox "Finished!", vbInformation, "Ok"
    End If

    ' clean up
    Set cWavConcat = Nothing
    Unload Me

End Sub
