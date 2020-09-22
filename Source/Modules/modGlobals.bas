Attribute VB_Name = "modGlobals"

' udt structure to hold the data.
Public Type udtType
    rndValue As Double
    lastName As String
    firstName As String
End Type

' Main array to play with
Public udtData() As udtType
Public numUDTData As Long


' Arrays to hold the last and first names so we only have to load from the text files once.
Public lastNames() As String
Public firstNames() As String
Public numfirstNames As Long
Public numlastNames As Long

