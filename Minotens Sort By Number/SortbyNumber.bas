Attribute VB_Name = "SortbyNumber"
'###############################(v.1)#
'####### SORT BY NUMBER MODULE #######
'######### Coded by: minoten #########
'############ Keith Simons ###########
'######## minoten@hotmail.com ########
'#####################################

'This module can freely be used/modified/released
'but please keep the above lines intact.


'Our function that is called on to sort a listview
Function sortbynum(strListview As ListView, strTempListview As ListView, strColumn As Integer, Descending As Boolean)
'strListview is your listview you want to sort
'strTempListview is a listview you just make to hold temporary information
'strColumn is the column number you would like to sort
'Descending is wether or not you want your listview to be in descending or ascending order

'Lets dim our variables
Dim strCount, strCount2, strCount3 As Integer
Dim strNumber, strTempNum As Integer
Dim stritsgood As Boolean
Dim strSplitVar, strSplit1, strSplit2, strSplit3, strSplit4, strSplit5, strSplit6, strSplit7, strSplit8, strSplit9 As Integer

'lets call on the function to prepare our listviews
prepareit strListview, strTempListview

'Set our variables

'used to loop through each row in our main listview
strCount = 2
'used to loop through each row in our temporary listview
strCount2 = 1
'this variable is used to let the do-loop know when you've found its spot in the listview
stritsgood = False

'Because of how you call on the first column of listview, we have to make a custom one.
If strColumn = 0 Then

    'Loop until we've gone through all listview rows
    Do Until strCount > strListview.ListItems.Count
    
        'Reset our variables on each loop
        strCount2 = 1
        strNumber = strListview.ListItems(strCount).Text
        strCount3 = 1
        
        
        '#### BEGIN SPEED OPTIMIZING CODE ####'
        'This code speeds us up a TON when dealing with big numbers
        'It basically splits the TempListview into 5 parts, cutting time to find its spot down 1/5
        
        'See if the TempListview listcount is larger than 29, if so, then split it and send the new strCount
        If strTempListview.ListItems.Count > 29 Then
        
            'Set our row splits throughout the TempListview using the split variable.
            strSplitVar = 0.2
'           strSplitVar = 0.1
            strSplit1 = Int(strTempListview.ListItems.Count * (strSplitVar * 1))
            strSplit2 = Int(strTempListview.ListItems.Count * (strSplitVar * 2))
            strSplit3 = Int(strTempListview.ListItems.Count * (strSplitVar * 3))
            strSplit4 = Int(strTempListview.ListItems.Count * (strSplitVar * 4))
'           strSplit5 = Int(strTempListview.ListItems.Count * (strSplitVar * 5))
'           strSplit6 = Int(strTempListview.ListItems.Count * (strSplitVar * 6))
'           strSplit7 = Int(strTempListview.ListItems.Count * (strSplitVar * 7))
'           strSplit8 = Int(strTempListview.ListItems.Count * (strSplitVar * 8))
'           strSplit9 = Int(strTempListview.ListItems.Count * (strSplitVar * 9))

            'If the given number is less than the TempListview split number then set it 0 to find the actual number
            If Int(strNumber) < Int(strTempListview.ListItems(strSplit1).Text) Then
            strCount2 = 1
            Else
                'IF the given number < given # then set it back to first split to search
                If Int(strNumber) < Int(strTempListview.ListItems(strSplit2).Text) Then
                strCount2 = strSplit1
                Else
                    'same as above but with diff var
                    If Int(strNumber) < Int(strTempListview.ListItems(strSplit3).Text) Then
                    strCount2 = strSplit2
                    Else
                        'same as above but with diff var
                        If Int(strNumber) < Int(strTempListview.ListItems(strSplit4).Text) Then
                        strCount2 = strSplit3
                        Else
                            'This is our last split, so its gotta be here
                            strCount2 = strSplit4

                        End If
                    End If
                End If
            End If
        
        End If
        '#### END SPEED OPTIMIZING CODE ####'
        
                'Loop through each row in the TempListview till strgood is true or it goes over the TempListview count
                Do Until stritsgood = True Or strCount2 > strTempListview.ListItems.Count
                  
                    'If the current number in listview is smaller than the number in TempListview then insert it into that line
                    If Int(strNumber) < Int(strTempListview.ListItems(strCount2).Text) Then
                    
                    transfer strListview, strTempListview, Int(strCount), Int(strCount2)
                          
                    'Let our do-loop know we've sorted it
                    stritsgood = True
                    
                    Else
                    
                        'If the current TempListview number it is on is the last one then just add it to the end
                        If strCount2 = strTempListview.ListItems.Count Then

                        transfer strListview, strTempListview, Int(strCount), Int(strCount2) + 1
                            
                        'Let our do-loop know we've sorted it
                        stritsgood = True
                        
                        End If
                        
                    End If

                'Add a number to our TempListview row number and loop it
                strCount2 = strCount2 + 1
                Loop
                
                'When moving to the next row in Listview, we want the do-loop to start on a clean slate
                stritsgood = False

    'Add a number to the Listview row number and loop it
    strCount = strCount + 1
    Loop

'If it is not the first column...
Else

    'Loop until we've gone through every row in listview
    Do Until strCount > strListview.ListItems.Count
    
        'Reset our variables on every loop
        strCount2 = 1
        strNumber = strListview.ListItems(strCount).ListSubItems(strColumn).Text
        strCount3 = 1
        
        
        '#### BEGIN SPEED OPTIMIZING CODE ####'
        'This code speeds us up a TON when dealing with big numbers varying in length
        'It basically splits the TempListview into 5 parts, cutting time to find its spot down 1/5
        
        'See if the TempListview listcount is larger than 29, if so, then split it and send the new strCount
        If strTempListview.ListItems.Count > 29 Then
        
            'Set our row splits throughout the TempListview using the split variable.
            strSplitVar = 0.2
'           strSplitVar = 0.1
            strSplit1 = Int(strTempListview.ListItems.Count * (strSplitVar * 1))
            strSplit2 = Int(strTempListview.ListItems.Count * (strSplitVar * 2))
            strSplit3 = Int(strTempListview.ListItems.Count * (strSplitVar * 3))
            strSplit4 = Int(strTempListview.ListItems.Count * (strSplitVar * 4))
'           strSplit5 = Int(strTempListview.ListItems.Count * (strSplitVar * 5))
'           strSplit6 = Int(strTempListview.ListItems.Count * (strSplitVar * 6))
'           strSplit7 = Int(strTempListview.ListItems.Count * (strSplitVar * 7))
'           strSplit8 = Int(strTempListview.ListItems.Count * (strSplitVar * 8))
'           strSplit9 = Int(strTempListview.ListItems.Count * (strSplitVar * 9))

            'If the given number is less than the TempListview split number then set it 0 to find the actual number
            If Int(strNumber) < Int(strTempListview.ListItems(strSplit1).ListSubItems(strColumn).Text) Then
            strCount2 = 1
            Else
                'IF the given number < given # then set it back to first split to search
                If Int(strNumber) < Int(strTempListview.ListItems(strSplit2).ListSubItems(strColumn).Text) Then
                strCount2 = strSplit1
                Else
                    'same as above but with diff var
                    If Int(strNumber) < Int(strTempListview.ListItems(strSplit3).ListSubItems(strColumn).Text) Then
                    strCount2 = strSplit2
                    Else
                        'same as above but with diff var
                        If Int(strNumber) < Int(strTempListview.ListItems(strSplit4).ListSubItems(strColumn).Text) Then
                        strCount2 = strSplit3
                        Else
                            'This is our last split, so its gotta be here
                            strCount2 = strSplit4

                        End If
                    End If
                End If
            End If
        
        End If
        '#### END SPEED OPTIMIZING CODE ####'
        
            
                'Loop until we've gone through every row in TempListview
                Do Until stritsgood = True Or strCount2 > strTempListview.ListItems.Count
                
                    'If the listview number is less than the templistview one, we add it to TempListview
                    If Int(strNumber) < Int(strTempListview.ListItems(strCount2).ListSubItems(strColumn).Text) Then

                    transfer strListview, strTempListview, Int(strCount), Int(strCount2)
                    
                    'Let our do-loop know we've sorted it
                    stritsgood = True
                    
                    Else
                    
                        'If we hit the last row in TempListview, we just add it to the end
                        If strCount2 = strTempListview.ListItems.Count Then
                        
                        transfer strListview, strTempListview, Int(strCount), Int(strCount2) + 1
                        
                        'Let our do-loop know we've sorted it
                        stritsgood = True
                        
                        End If
                        
                    End If

                'Add a number to go to the next row in TempListview and loop it
                strCount2 = strCount2 + 1
                Loop
                
                'Reset our variable on each loop so they have a clean slate
                stritsgood = False
            
    'Add a number to go to the next row in Listview and loop it
    strCount = strCount + 1
    Loop


End If

'We are done sorting the data from listview to templistview.
'We need to take the templistview data and throw it into listview
transferall strListview, strTempListview, Descending

End Function

Function transfer(strListview As ListView, strTempListview As ListView, strRow As Integer, strInsertRow As Integer)
'Dim our variable
Dim strCount3

'strListview is the listview we are pulling data from
'strTempListview is the listview we are adding data to
'strRow is the row we are pulling data from
'strInsertrow is the row we are adding data to

'Set our variable
strCount3 = 1

    'Just in case one of our columns is empty, we don't want it to error
    On Error Resume Next:
    
    'We transfer the first column data from our original row to our other listview
    strTempListview.ListItems.Add strInsertRow, , strListview.ListItems(strRow).Text
    strTempListview.ListItems(strInsertRow).Bold = strListview.ListItems(strRow).Bold
    strTempListview.ListItems(strInsertRow).ForeColor = strListview.ListItems(strRow).ForeColor
    strTempListview.ListItems(strInsertRow).Key = strListview.ListItems(strRow).Key
    strTempListview.ListItems(strInsertRow).Tag = strListview.ListItems(strRow).Tag
    strTempListview.ListItems(strInsertRow).ToolTipText = strListview.ListItems(strRow).ToolTipText
    
        'If the listview we are transfering from has checkboxes, we are sure to include their value
        If strListview.Checkboxes = True Then
        strTempListview.ListItems(strInsertRow).Checked = strListview.ListItems(strRow).Checked
        End If
    
    'Loop through all columns except the first (we just took care of that)
    Do Until strCount3 > strListview.ColumnHeaders.Count - 1
    
    'We transfer the data from our original row to our other listview (except first row)
    strTempListview.ListItems(strInsertRow).ListSubItems.Add , , strListview.ListItems(strRow).ListSubItems(strCount3).Text
    strTempListview.ListItems(strInsertRow).ListSubItems(strCount3).Bold = strListview.ListItems(strRow).ListSubItems(strCount3).Bold
    strTempListview.ListItems(strInsertRow).ListSubItems(strCount3).ForeColor = strListview.ListItems(strRow).ListSubItems(strCount3).ForeColor
    strTempListview.ListItems(strInsertRow).ListSubItems(strCount3).Key = strListview.ListItems(strRow).ListSubItems(strCount3).Key
    strTempListview.ListItems(strInsertRow).ListSubItems(strCount3).ReportIcon = strListview.ListItems(strRow).ListSubItems(strCount3).ReportIcon
    strTempListview.ListItems(strInsertRow).ListSubItems(strCount3).Tag = strListview.ListItems(strRow).ListSubItems(strCount3).Tag
    strTempListview.ListItems(strInsertRow).ListSubItems(strCount3).ToolTipText = strListview.ListItems(strRow).ListSubItems(strCount3).ToolTipText
                     
                     
     'Add one to columnheaders and then loop
     strCount3 = strCount3 + 1
     Loop


End Function

Function transferall(strListview As ListView, strTempListview As ListView, strDecending As Boolean)
'Dim our variable
Dim strCount As Integer

'Clear our list we are transfering to
strListview.ListItems.Clear

'If we want to order to be descending, then...
If strDecending = True Then

'Set strCount to the row number for our list we are adding from
strCount = strTempListview.ListItems.Count

    'Loop through our templistview BACKWARDS, generating a DESCENDING order
    Do Until strCount = 0
    
    'Transfer the current row to our main listview
    transfer strTempListview, strListview, strCount, strListview.ListItems.Count + 1
    
    'Subtract one from our row number and loop
    strCount = strCount - 1
    Loop

'If we want to order to be ascending
Else

'Set our variable (theres no 0's in listview)
strCount = 1

    'Loop through our tempListview until we go over the row number
    Do Until strTempListview.ListItems.Count < strCount
    
    'Transfer the current row to our main listview
    transfer strTempListview, strListview, strCount, strListview.ListItems.Count + 1
    
    'Add one to our row number and loop
    strCount = strCount + 1
    Loop

End If

End Function

Function prepareit(strListview As ListView, strTempListview As ListView)
'Dim our variable
Dim strCount As Integer

'Change our tempListview to report mode, clear it, set it to no checkboxes, and clear it
strTempListview.View = lvwReport
strTempListview.ColumnHeaders.Clear
strTempListview.Checkboxes = False
strTempListview.ListItems.Clear

'Set sort order for main listview to 1, makes sorting a bit quicker in some cases
strListview.SortOrder = 1
strListview.Sorted = True

'Unsort it, or else it'll mess us all up (trying to sort while adding them)
strListview.Sorted = False

    'If the listview we are copying from has checkboxes, make our templistview have theem too
    If strListview.Checkboxes = True Then
    strTempListview.Checkboxes = True
    End If

'Set our variable to 1 (no 0 column)
strCount = 1
     
     'Loop until our strCount is more than the column header count
     Do Until strCount > strListview.ColumnHeaders.Count

     'Copy the column from our main listview to our temporary listview (even keeps the text)
     strTempListview.ColumnHeaders.Add , , strListview.ColumnHeaders(strCount).Text
     
     'Add one to our count, increasing the column header, then loop
     strCount = strCount + 1
     Loop

     'Transfer the first row to the templist
     transfer strListview, strTempListview, 1, 1

End Function

