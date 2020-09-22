Attribute VB_Name = "modUDT"
'---------------------------------------------------------------------------------------
' Procedure : randomizeArray
' Author    : bikeguy
' Date      : 5/14/2012
' Purpose   : Assigns a random value to rndValue property of each element, and then calls
'             quickSortRand to reorder the array based on this property
'---------------------------------------------------------------------------------------
Public Sub randomizeArray()
    Dim ctr As Long
    For ctr = 0 To numUDTData - 1
        udtData(ctr).rndValue = Rnd
    Next
    quickSortRand udtData, 0, numUDTData - 1
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : loadNameData
' Author    : nie1jlw
' Date      : 5/14/2012
' Purpose   : Loads and parses last and first name data from the files in the Data folder under the
'             the main app folder.  This data is from the US census bureau and is kind of neat.
' Note      : There are MUCH better (read "faster") ways to do this.  If you need a very good
'             function to read data search for "GetFileQuick" on the Planet site.  Excellent, stable
'             and very, very fast.
'---------------------------------------------------------------------------------------
Public Sub loadNameData()
    Dim iLoc As Integer     ' Index of where the first space is found
    Dim iLine As String     ' Holder for each line read in.
    Dim iHand As Integer    ' Handle variable
    Dim iFile As String     ' Name of current file being processed
    
    ' Get the last names
    iFile = App.Path & "\Data\Lastnames.txt"
    
    ' Check for file existence
    If Len(Dir(iFile)) > 0 Then
        ' If found, get a handle
        iHand = FreeFile()
        
        ' Open the file
        Open iFile For Input As #iHand
        
        Do While Not EOF(iHand)
            ' Get a line to work on, and put it into the variable iLine
            Line Input #iHand, iLine
            ' Trim the line of leading and trailing spaces
            iLine = Trim$(iLine)
            If Len(iLine) > 0 Then
                ' If the line has data, add it to the array of lastnames
                numlastNames = numlastNames + 1
                ReDim Preserve lastNames(numlastNames)
                lastNames(numlastNames - 1) = iLine
                
            End If
        Loop
        ' Close the file
        Close #iHand
    End If
    
    ' Get the first names
    iFile = App.Path & "\Data\Firstnames.txt"
    
    ' Check for file existence
    If Len(Dir(iFile)) > 0 Then
        ' If found, get a handle
        iHand = FreeFile()
        
        ' Open the file
        Open iFile For Input As #iHand
        
        Do While Not EOF(iHand)
            ' Get a line to work on, and put it into the variable iLine
            Line Input #iHand, iLine
            ' Trim the line of leading and trailing spaces
            iLine = Trim$(iLine)
            If Len(iLine) > 0 Then
                ' If the line has data, add it to the array of firstnames
                numfirstNames = numfirstNames + 1
                ReDim Preserve firstNames(numfirstNames)
                firstNames(numfirstNames - 1) = iLine

            End If
        Loop
        ' Close the file
        Close #iHand
    End If
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : createUDTArray
' Author    : bikeGuy
' Date      : 5/14/2012
' Purpose   : Build an array of UDT items.  numToBuild parameter defines the number of
'             items to build.
'---------------------------------------------------------------------------------------
'
Public Function createUDTArray(numToBuild As Long) As udtType()

    Dim rArray() As udtType ' Return value - array of UDTs, each containing a last and first name
                            ' and rndValue
    Dim numRarray As Long   ' Number to keep track of side of UDT array
    Dim lNdx As Long        ' Index of last name to add
    Dim fNdx As Long        ' Index of first name to add
    
    ' Seed the randomizer
    Randomize Timer
    
    
    ' Check to see if the data has already been loaded
    If numlastNames = 0 Or numfirstNames = 0 Then
        ' Not there, so load it
        loadNameData
    End If
    
    ' Build each UDT item
    For ctr = 0 To numToBuild
        numRarray = numRarray + 1
        ReDim Preserve rArray(numRarray)
        ' Get a random last name index
        lNdx = Int(Rnd * numlastNames) + 1
        ' Get a random first name index
        fNdx = Int(Rnd * numfirstNames) + 1
        
        ' Build the data item
        rArray(numRarray - 1).lastName = lastNames(lNdx)
        rArray(numRarray - 1).firstName = lastNames(fNdx)
        rArray(numRarray - 1).rndValue = Rnd
    Next
    createUDTArray = rArray
End Function
