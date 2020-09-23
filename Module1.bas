Attribute VB_Name = "Module1"
Public Const DivN = 14
Public Const C1 = 261.63
Public Const CS1 = 277.18
Public Const D1 = 293.66
Public Const DS1 = 311.13
Public Const E1 = 329.63
Public Const F1 = 349.23
Public Const FS1 = 369.99
Public Const G1 = 392
Public Const GS1 = 415.3
Public Const A1 = 440
Public Const AS1 = 466.16
Public Const B1 = 493.88
Public Const P1 = 20000
Global Tempo  As Long
Global LastTempo As Long
Global LastKeyColor As Variant
Global LastKey As Integer
Public cCount As Integer
Global TempX As Integer, TempY As Integer, IntCnt As Integer
Public Const Red = &HC0&
Public Const Blue = &HC00000
Public Const Green = &HC000&
Public Const Dur = 150
Public Const StepF = 150
Global StepStart As Long

Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long


















