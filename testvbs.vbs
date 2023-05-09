' Declare and initialize the array
Dim numbers(9)
numbers(0) = 10
numbers(1) = 5
numbers(2) = 8
numbers(3) = 2
numbers(4) = 3
numbers(5) = 123
numbers(6) = 7
numbers(7) = 6
numbers(8) = 4
numbers(9) = 9

' Call the Quicksort function to sort the array
Quicksort numbers, LBound(numbers), UBound(numbers)

WScript.Echo("LBOUND: " & LBound(numbers))
WScript.Echo("UBOUND: " & UBound(numbers))

' Display the sorted array
For i = 0 To UBound(numbers)
    WScript.Echo numbers(i)
Next

' Quicksort algorithm
Sub Quicksort(arr, low, high)
    If low < high Then
        ' Partition the array and get the pivot index
        Dim pivotIndex
        pivotIndex = Partition(arr, low, high)
        
        ' Recursively sort the sub-arrays before and after the pivot
        Quicksort arr, low, pivotIndex - 1
        Quicksort arr, pivotIndex + 1, high
    End If
End Sub

' Partition function for Quicksort
Function Partition(arr, low, high)
    Dim pivot
    pivot = arr(high) ' Choose the last element as the pivot
    Dim i
    i = low - 1 ' Index of the smaller element
    
    For j = low To high - 1
        If arr(j) <= pivot Then
            ' Swap arr(i) and arr(j)
            i = i + 1
            Swap arr, i, j
        End If
    Next
    
    ' Swap arr(i+1) and arr(high) (put the pivot in the correct position)
    Swap arr, i + 1, high
    
    ' Return the pivot index
    Partition = i + 1
End Function

' Swap function for exchanging array elements
Sub Swap(arr, i, j)
    Dim temp
    temp = arr(i)
    arr(i) = arr(j)
    arr(j) = temp
End Sub