Attribute VB_Name = "AddToStorageUnit"
Attribute VB_Description = "L_TO_CREATE_SINGLE - Add to Storage Unit"

Sub AddToStorageUnit()

    ' Declare variables For SAP connection And Function Call
    Dim Functions As Object
    Dim myConnection As Object
    Dim RfcCallTransaction As Object
    Dim ExportParams As Object
    Dim ImportParams As Object
    Dim cellStorageUnit As String
    Dim i As Long
    Dim lastRow As Long
    Dim firstRow As Long
    Dim resultRow As Long
    Dim reasonRow As String

    Dim cellMovementType As String
    Dim cellStorageLocation As String
    Dim formatedStorageLocation As String
    Dim cellPlant As String
    Dim cellWarehouseNumber As String
    Dim cellReason As String
    Dim columnMaterial As String
    Dim columnQuantity As String
    Dim columFromStorageType As String
    Dim columFromStorageBin As String
    Dim columToStorageType As String
    Dim columToStorageBin As String
    Dim columToStorageUnit As String
    Dim columDate As String
    Dim columReason As String
    Dim columResult As String
    Dim cellPercentage As String
    Dim totalRows As Long
    Dim percentageCompleted As Double
    Dim currentIteration As Long


    cellMovementType = "B1"
    cellStorageLocation = "B2"
    cellPlant = "B3"
    cellWarehouseNumber = "B4"
    cellReason = "B5"
    cellPercentage = "B6"

    columnMaterial = "A"
    columnQuantity = "B"
    columFromStorageType = "C"
    columFromStorageBin = "D"
    columToStorageType = "E"
    columToStorageBin = "F"
    columToStorageUnit = "G"
    columDate = "H"
    columReason = "I"
    columResult = "J"
    firstRow = 9

    ' Check If all required cells have values
    Dim missingCells As String
    missingCells = ""
    
    If Trim(ThisWorkbook.ActiveSheet.Range(cellMovementType).Value) = "" Then
        missingCells = missingCells & "MovementType " & cellMovementType & " is missing." & vbCrLf
    End If
    
    If Trim(ThisWorkbook.ActiveSheet.Range(cellStorageLocation).Value) = "" Then
        missingCells = missingCells & "StorageLocation " & cellStorageLocation & " is missing." & vbCrLf
    End If
    
    If Trim(ThisWorkbook.ActiveSheet.Range(cellPlant).Value) = "" Then
        missingCells = missingCells & "Plant " & cellPlant & " is missing." & vbCrLf
    End If
    
    If Trim(ThisWorkbook.ActiveSheet.Range(cellWarehouseNumber).Value) = "" Then
        missingCells = missingCells & "WarehouseNumber " & cellWarehouseNumber & " is missing." & vbCrLf
    End If
    
    If Trim(ThisWorkbook.ActiveSheet.Range(cellReason).Value) = "" Then
        missingCells = missingCells & "Reason " & cellReason & " is missing." & vbCrLf
    End If
    
    If missingCells <> "" Then
        MsgBox "Please fill in the following required cells:" & vbCrLf & missingCells, vbExclamation
        Exit Sub
    End If


    ' Create SAP Function object And connection
    Set Functions = CreateObject("SAP.Functions")
    Set myConnection = Functions.Connection

    ' Set SAP connection parameters (Set these outside the Loop)
    ' myConnection.ApplicationServer = ""
    ' myConnection.SystemNumber = ""
    ' myConnection.Client = ""
    ' myConnection.User = ""
    ' myConnection.Password = ""

    ' Comment this section And uncomment avobe To Do tests As this will prompt For user And password every time
    If myConnection.Logon(0, False) <> True Then
        ' MsgBox "Cannot Log on To SAP"
    End If

    ' Connect To SAP (connect once before the Loop)
    If Not myConnection.Logon(0, True) Then
        MsgBox "Unable To connect To SAP", vbCritical
     Exit Sub
    End If

    ' Set the range of rows To process
    lastRow = ThisWorkbook.ActiveSheet.Cells(Rows.Count, columnMaterial).End(xlUp).Row
    resultRow = ThisWorkbook.ActiveSheet.Cells(Rows.Count, columResult).End(xlUp).Row
    reasonRow = ThisWorkbook.ActiveSheet.Range(cellReason).Value
    If resultRow > firstRow Then
        firstRow = resultRow + 1
    End If
    totalRows = lastRow - resultRow


    formatedStorageLocation = Format(ThisWorkbook.ActiveSheet.Range(cellStorageLocation).Value, "0000")   

        On Error Goto Cleanup

            ' Loop through each row
            For i = firstRow To lastRow
                currentIteration = currentIteration + 1 
                ' Reinitialize the RFC Function For each iteration
                Set RfcCallTransaction = Functions.Add("L_TO_CREATE_SINGLE")

                ' Set the import parameters For the RFC Function
                Set ExportParams = RfcCallTransaction.Exports
                cellStorageUnit = Format(ThisWorkbook.ActiveSheet.Range(columToStorageUnit & i).Value, "00000000000000000000")
                ' Set export parameters For the current row
                ExportParams("I_BWLVS").Value = ThisWorkbook.ActiveSheet.Range(cellMovementType).Value
                ExportParams("I_WERKS").Value = ThisWorkbook.ActiveSheet.Range(cellPlant).Value
                ExportParams("I_LGNUM").Value = ThisWorkbook.ActiveSheet.Range(cellWarehouseNumber).Value
                ExportParams("I_LGORT").Value = formatedStorageLocation
                ExportParams("I_ANFME").Value = ThisWorkbook.ActiveSheet.Range(columnQuantity & i).Value
                ExportParams("I_MATNR").Value = ThisWorkbook.ActiveSheet.Range(columnMaterial & i).Value
                ExportParams("I_ALTME").Value = ""
                ExportParams("I_LETYP").Value = "001"
                ' FROM SOURCE
                ExportParams("I_VLTYP").Value = ThisWorkbook.ActiveSheet.Range(columFromStorageType & i).Value
                ExportParams("I_VLBER").Value = "001"
                ExportParams("I_VLPLA").Value = ThisWorkbook.ActiveSheet.Range(columFromStorageBin & i).Value
                ' To DESTINATION
                ExportParams("I_NLTYP").Value = ThisWorkbook.ActiveSheet.Range(columToStorageType & i).Value
                ExportParams("I_NLBER").Value = "001"
                ExportParams("I_NLPLA").Value = ThisWorkbook.ActiveSheet.Range(columToStorageBin & i).Value
                ExportParams("I_NLENR").Value = cellStorageUnit


                ThisWorkbook.ActiveSheet.Range(columDate & i).Value = Now
                ThisWorkbook.ActiveSheet.Range(columReason & i).Value = reasonRow
                ThisWorkbook.ActiveSheet.Range(cellPercentage).Value = (currentIteration / totalRows)


                ' Call the Function
                If RfcCallTransaction.Call = True Then
                    ' Process the result
                    Set ImportParams = RfcCallTransaction.Imports
                    Dim TransferOrderNumber As String
                    TransferOrderNumber = ImportParams("E_TANUM").Value

                    ' Display the transfer order number
                    ThisWorkbook.ActiveSheet.Range(columResult & i).Value = TransferOrderNumber

                Else
                    ' Handle RFC Call failure
                    Dim ErrorMessage As String
                    ErrorMessage = "Error: " & RfcCallTransaction.Exception
                    ThisWorkbook.ActiveSheet.Range(columResult & i).Value = ErrorMessage
                    Debug.Print ErrorMessage

                    ' Log off And reinitialize connection
                    myConnection.Logoff
                    If Not myConnection.Logon(0, True) Then
                        MsgBox "Unable To reconnect To SAP", vbCritical
                     Exit Sub
                    End If            

                    ' Reinitialize RFC Function object
                    Set RfcCallTransaction = Functions.Add("L_TO_CREATE_SINGLE")
                    Set ExportParams = RfcCallTransaction.Exports
                End If

            Next i

 Cleanup:
            ' Log off from SAP
            If Not myConnection Is Nothing Then
                myConnection.Logoff
            End If

            If Err.Number <> 0 Then
                MsgBox "An error occurred: " & Err.Description, vbCritical
            End If

            ThisWorkbook.ActiveSheet.Range(cellPercentage).Value = 0
            ThisWorkbook.ActiveSheet.Range(cellReason).Value = ""

End Sub