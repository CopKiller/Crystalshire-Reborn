Attribute VB_Name = "modTypes"
Option Explicit

Public Item() As ItemRec

Private Type ItemRec
    Name As String * NAME_LENGTH
End Type
