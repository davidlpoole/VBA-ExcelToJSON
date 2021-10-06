Sub exportJSON()

  ''''''''''''' Declare variables
  Dim data() As String
  Dim Rows, Columns As Integer
  Dim tempName, tempVal, tempField, tempRow, tempObj As String
    
    
  '''''''''''''' Declare output file name and location
  ' File must be saved before running (TODO: handle error)
  Dim outFile, collectionName As String
  myFile = ActiveWorkbook.Path & "\output.json"
  collectionName = "data"
      
  
  '''''''''''''' Calculate the number of rows and columns in the sheet
  numCols = WorksheetFunction.CountA(Worksheets(1).Rows(1))
  numRows = WorksheetFunction.CountA(Worksheets(1).Columns(1)) - 1
  ReDim data(numRows, numCols)
  
  
  '''''''''''''' Iterate through cells and store into array
  For j = 1 To numRows + 1      ' For each row
    For i = 1 To numCols + 1    ' For each column in the row
      data(j - 1, i - 1) = Trim(Cells(j, i).Value)
    Next
  Next

  '''''''''''''' Iterate through the array building up the JSON object
  tempObj = ""
    
  ' Iterate over rows (j)
  For j = 1 To numRows
    tempRow = ""
      
      ' Iterate over the columns (i) in each row (j)
      For i = 1 To numCols
        tempName = ""
        tempVal = ""
        tempField = ""
        tempName = Chr(34) + data(0, i - 1) + Chr(34)  ' Name of the field with double quotes
        tempVal = Chr(34) + data(j, i - 1) + Chr(34)   ' Value of the field with double quotes
        tempField = tempName + ": " & tempVal          ' Combine name and value
        tempRow = tempRow + tempField                  ' Combine this field with previous
        If i < numCols Then tempRow = tempRow + ","    ' Skip the comma on last column
      Next ' Next i (column in spreadsheet)
    tempObj = tempObj + "  " + "{" + tempRow + "}"   ' wrap row in curly braces and add a carriage return
    If j < numRows Then tempObj = tempObj + ","       ' skip the comma on last row
    tempObj = tempObj + vbCrLf + "  "                ' add a carriage return
  Next ' Next j (row in spreadsheet)
    
    
  ' Wrap row data with curly braces and a 'collection name'
  tempObj = "{" + vbCrLf + "  " _
  + Chr(34) + collectionName + Chr(34) + ": [" + vbCrLf + "  " _
  + tempObj _
  + "]" + vbCrLf + _
  "}"
    
    
  ''''''''''''''' Create a file and output the JSON data to it
  Open myFile For Output As #1
  Print #1, tempObj
  Close #1
    
End Sub



