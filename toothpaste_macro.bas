Attribute VB_Name = "Module1"
Sub main()

    Dim num_cities As Integer 'number of cities and sheets in the data file
    num_cities = 0
    Dim parameter_options() As String 'carries all the parameter options from the first sheet
    Dim num_parameter_options As Integer
    num_parameter_options = 0
    Dim sums() As Double 'carries sums from each sheet of the data file to the output file
    year_ = Cells(3, 3) 'user input
    data_file = Cells(3, 4) + Cells(3, 5)
    output_file = Cells(3, 6) + Cells(3, 7)
    sheet_name = Cells(3, 8)
    parameter_ = Cells(3, 2)
    starting_row = 1 'expected starting row as derived from first sheet
    Dim year_category_content As String 'used to account for differences in placement of desired year column between sheets
    Dim year_content As String
    
    Call gather_cities(data_file, output_file, sheet_name, num_cities) 'gather cities and place on output sheet
    Workbooks(data_file).Sheets(3).Activate 'go to first data sheet
    starting_row = find_row(1, "MANUFACTURER", output_file, 1) + 1 'find row on which input starts
    year_category_content = Cells(starting_row - 2, year_).MergeArea.Cells(1, 1) 'take note of content of user desired year
    year_content = Cells(starting_row - 1, year_)
    Call sort_column(parameter_, ActiveSheet.Index, starting_row) 'gather parameter options and place on output sheet
    Call gather_parameter_options(data_file, output_file, parameter_options(), num_parameter_options, starting_row, parameter_)
    
    Dim input_row As Integer 'used to track row position in data file
    Dim mod_param_opt() As String 'the parameter options of the given sheet
    Dim mod_num_param_opt As Integer
    
    For i = 1 To num_cities 'go through all sheets, gather data, and place on output sheet
        input_row = starting_row
        Workbooks(data_file).Sheets(i + 1).Activate 'go to next sheet
        Call sort_column(parameter_, ActiveSheet.Index, input_row)
        ReDim sums(0) 'prepare variables for the new sheet
        ReDim mod_param_opt(0)
        mod_num_param_opt = 0
        
        outer_sum_column = find_column(input_row - 2, year_category_content, output_file, 1) 'find correct column for desired year
        sum_column = find_column(input_row - 1, year_content, output_file, outer_sum_column)
        
        While Cells(input_row, 1) <> "" 'continue until no more data
            current_sum = 0
            
            If Cells(input_row, sum_column) <> "NA" Then 'add first value to sum
                current_sum = Cells(input_row, sum_column)
            End If
            
            parameter_option = Cells(input_row, parameter_) 'consider the first parameter option
            mod_num_param_opt = mod_num_param_opt + 1
            input_row = input_row + 1
            
            While Cells(input_row, parameter_) = parameter_option 'continue adding to sum until parameter option changes
                If Cells(input_row, sum_column) <> "NA" Then
                    current_sum = current_sum + Cells(input_row, sum_column)
                End If
                input_row = input_row + 1
            Wend
            
            ReDim Preserve mod_param_opt(mod_num_param_opt) 'add to array of options
            mod_param_opt(UBound(mod_param_opt)) = parameter_option
            ReDim Preserve sums(mod_num_param_opt) 'add to array of sums
            sums(UBound(sums)) = current_sum
        Wend

        mod_num_param_opt = mod_num_param_opt + 1 'add total to array
        ReDim Preserve mod_param_opt(mod_num_param_opt)
        mod_param_opt(UBound(mod_param_opt)) = "TOTAL"
        ReDim Preserve sums(mod_num_param_opt)
        If Cells(input_row, sum_column) <> "NA" Then
                sums(UBound(sums)) = Cells(input_row, sum_column)
        Else
            sums(UBound(sums)) = 0
        End If

        Workbooks(output_file).Activate 'transfer data to output file
        
        If mod_num_param_opt = num_parameter_options Then 'if no parameter options were missing, print on output file as normal
            For j = 3 To mod_num_param_opt + 2
                Cells(j, i + 1) = sums(j - 2)
            Next
            GoTo Unmodified
        End If
        
        m = 1
        n = 1
        
        While Not (m > num_parameter_options Or n > mod_num_param_opt) 'otherwise parse through arrays in parallel
            If parameter_options(m) = mod_param_opt(n) Then 'if parameter option is included, add corresponding sum to output file
                Cells(m + 2, i + 1) = sums(n)
                m = m + 1
                n = n + 1
            Else 'if parameter option is missing, input "0" into output file
                Cells(m + 2, i + 1) = 0
                m = m + 1
            End If
        Wend
        
Unmodified:
        Cells(num_parameter_options + 2, i + 1).NumberFormat = "0.0" 'format to single precision
    Next

End Sub

Sub gather_cities(ByVal data_file As String, ByVal output_file As String, ByVal sheet_name As String, ByRef num_cities As Integer)
    Dim cities() As String 'an array to hold all city names
    input_ = ""
    
    Workbooks(data_file).Sheets(1).Activate 'go to data window
    row_ = find_row(3, "Nat Total/CN", output_file, 1) 'find starting row
    
    While row_ <= Cells.SpecialCells(xlCellTypeLastCell).Row 'while end of list not yet reached
        input_ = Cells(row_, 3) 'take in cell contents to parse
        city_name = ""
        city_start = 0
        
        For i = Len(input_) - 3 To 1 Step -1 'find contents before "/CN"
            If Mid(input_, i, 1) = "/" Or Mid(input_, i, 1) = " " Then
                city_start = i + 1
                Exit For
            End If
        Next
        
        For i = city_start To Len(input_) - 3 'record city name
            city_name = city_name + Mid(input_, i, 1)
        Next
        
        num_cities = num_cities + 1 'place into array
        ReDim Preserve cities(num_cities)
        cities(UBound(cities)) = city_name
        row_ = row_ + 1
    Wend
    
    Workbooks(output_file).Activate 'place city names in output file
    Sheets.Add After:=Sheets(Sheets.Count)
    ActiveSheet.Name = sheet_name
    
    For i = 1 To num_cities
        Cells(2, i + 1) = cities(i)
    Next
End Sub

Function find_row(ByVal desired_column As Integer, ByVal desired_contents As String, ByVal output_file As String, ByVal starting_row As Integer) As Integer 'finds the row at which the desired contents appear in specified column
    find_row = starting_row
    
    While Cells(find_row, desired_column) <> desired_contents And find_row <= Cells.SpecialCells(xlCellTypeLastCell).Row 'search for row until found
        find_row = find_row + 1
    Wend
End Function

Function find_column(ByVal desired_row As Integer, ByVal desired_contents As String, ByVal output_file As String, ByVal starting_col As Integer) As Integer 'finds the column at which the desired contents appear in specified row
    find_column = starting_col
    
    While Cells(desired_row, find_column).MergeArea.Cells(1, 1) <> desired_contents And find_column <= Cells.SpecialCells(xlCellTypeLastCell).Column 'search for column until found
        find_column = find_column + 1
    Wend
End Function

Sub sort_column(ByVal column_num As Integer, ByVal sheet_num As Integer, ByVal row_num As Integer)
    ActiveWorkbook.Sheets(sheet_num).Sort.SortFields.Clear
    ActiveWorkbook.Sheets(sheet_num).Sort.SortFields.Add2 Key:=Columns(column_num), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Sheets(sheet_num).Sort
        .SetRange Range(Cells(row_num, 1), Cells(Cells.SpecialCells(xlCellTypeLastCell).Row, Cells.SpecialCells(xlCellTypeLastCell).Column))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub gather_parameter_options(ByVal data_file As String, ByVal output_file As String, ByRef parameter_options() As String, ByRef num_parameter_options As Integer, ByVal input_row As Integer, ByVal parameter_ As Integer)
    ReDim parameter_options(0) 'array to hold parameter options
    
    While Cells(input_row, 1) <> "" 'continue until end of data
            parameter_option = Cells(input_row, parameter_) 'take in first parameter option
            num_parameter_options = num_parameter_options + 1
            input_row = input_row + 1
            
            While Cells(input_row, parameter_) = parameter_option 'parse until new option found
                input_row = input_row + 1
            Wend
            
            ReDim Preserve parameter_options(num_parameter_options) 'add to array of options
            parameter_options(UBound(parameter_options)) = parameter_option
    Wend
    
    num_parameter_options = num_parameter_options + 1 'add total to array
    ReDim Preserve parameter_options(num_parameter_options)
    parameter_options(UBound(parameter_options)) = "TOTAL"
        
    Workbooks(output_file).Activate 'transfer information to output file
    
    For j = 3 To num_parameter_options + 2
        Cells(j, 1) = parameter_options(j - 2)
    Next
End Sub




