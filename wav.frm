VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "wav.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Dim filePath As String
    Dim fileContent As String
    Dim i As Long
    Dim char As String
    Dim sampleRate As Long: sampleRate = 44100
    Dim duration As Double: duration = 0.3 ' segundos por nota
    Dim samples() As Integer
    Dim pos As Long: pos = 0

    filePath = Application.GetOpenFilename("Text Files (*.txt), *.txt")
    If filePath = "False" Then Exit Sub

    fileContent = ReadFileText(filePath)
    ReDim samples(1 To Len(fileContent) * sampleRate * duration)

    For i = 1 To Len(fileContent)
        char = Mid(fileContent, i, 1)
        If char Like "[0-9A-H]" Then
            Dim freq As Double
            freq = NoteFrequency(char)
            If freq > 0 Then
                GenerateSineWave samples, pos, freq, duration, sampleRate
                pos = pos + duration * sampleRate
            End If
        End If
    Next i

    WriteWav "output.wav", samples, sampleRate
    MsgBox "WAV gerado: output.wav"
End Sub

Private Function ReadFileText(path As String) As String
    Dim f As Integer: f = FreeFile
    Open path For Input As #f
    ReadFileText = Input$(LOF(f), #f)
    Close #f
End Function

Private Function NoteFrequency(ch As String) As Double
    Select Case ch
        Case "0": NoteFrequency = 261.63 ' C4
        Case "1": NoteFrequency = 293.66 ' D4
        Case "2": NoteFrequency = 329.63 ' E4
        Case "3": NoteFrequency = 349.23 ' F4
        Case "4": NoteFrequency = 392#   ' G4
        Case "5": NoteFrequency = 440#   ' A4
        Case "6": NoteFrequency = 493.88 ' B4
        Case "7": NoteFrequency = 523.25 ' C5
        Case "8": NoteFrequency = 587.33 ' D5
        Case "9": NoteFrequency = 659.25 ' E5
        Case "A": NoteFrequency = 440#   ' A4
        Case "B": NoteFrequency = 493.88 ' B4
        Case "C": NoteFrequency = 523.25 ' C5
        Case "D": NoteFrequency = 587.33 ' D5
        Case "E": NoteFrequency = 659.25 ' E5
        Case "F": NoteFrequency = 698.46 ' F5
        Case "G": NoteFrequency = 783.99 ' G5
        Case "H": NoteFrequency = 880#   ' A5
        Case Else: NoteFrequency = 0
    End Select
End Function

Private Sub GenerateSineWave(ByRef arr() As Integer, startPos As Long, freq As Double, duration As Double, rate As Long)
    Dim i As Long
    Dim total As Long: total = duration * rate
    For i = 0 To total - 1
        arr(startPos + i + 1) = Int(10000 * Sin(2 * 3.14159 * freq * i / rate))
    Next i
End Sub

Private Sub WriteWav(filename As String, data() As Integer, rate As Long)
    Dim f As Integer: f = FreeFile
    Dim i As Long, datasize As Long
    datasize = UBound(data) * 2

    Open filename For Binary As #f

    Put #f, , "RIFF"
    Put #f, , CLng(36 + datasize)
    Put #f, , "WAVE"
    Put #f, , "fmt "
    Put #f, , CLng(16)
    Put #f, , CInt(1)          ' PCM
    Put #f, , CInt(1)          ' Mono
    Put #f, , CLng(rate)       ' SampleRate
    Put #f, , CLng(rate * 2)   ' ByteRate
    Put #f, , CInt(2)          ' BlockAlign
    Put #f, , CInt(16)         ' BitsPerSample
    Put #f, , "data"
    Put #f, , CLng(datasize)

    For i = 1 To UBound(data)
        Put #f, , CInt(data(i))
    Next i

    Close #f
End Sub

