Option Explicit

Private queue As Collection

' Initialize the queue
Private Sub Class_Initialize()
    Set queue = New Collection
End Sub

' Add an item to the queue
Public Sub Enqueue(item As clsUDPQueueItem)
    queue.Add item
End Sub

' Remove and return the first item in the queue
Public Function Dequeue() As clsUDPQueueItem
    If queue.count = 0 Then
        Set Dequeue = Nothing
        Exit Function
    End If
    
    Set Dequeue = queue(1)
    queue.Remove 1
End Function

' Peek at the first item without removing it
Public Function Peek() As clsUDPQueueItem
    If queue.count = 0 Then
        Set Peek = Nothing
    Else
        Set Peek = queue(1)
    End If
End Function

' Number of items in the queue
Public Property Get count() As Long
    count = queue.count
End Property

' Check if the queue is empty
Public Property Get IsEmpty() As Boolean
    IsEmpty = (queue.count = 0)
End Property

' Clear the queue
Public Sub Clear()
    Set queue = New Collection
End Sub

' Remove and return the first item if available, else return Nothing
Public Function TryDequeue() As clsUDPQueueItem
    If queue.count = 0 Then
        Set TryDequeue = Nothing
    Else
        Set TryDequeue = queue(1)
        queue.Remove 1
    End If
End Function

