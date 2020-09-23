VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Form1"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.TextBox Text2 
      Height          =   1815
      Left            =   3240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   2
      Top             =   480
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Result"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°
'This is a very simple code to demonstrate how to search a list box
'for an item dynamically by entering a word/phrase in a text area.
'Double clicking on an item at the List, will cause infos about it
'to appear at the text area on the right.
'Of course the content of the list (and that from the item-infos too),
'can be loaded from text/other files... In this example though, the list items
'(and the infos) are coded.
'Feel free to use this code or parts of it for your applications :)
'°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°°

Private Sub Form_Load()
    'Filling the ListBox with various options
    'NOTE: you must set the "Sorted" property of the ListBox to true,
    'when you want to create an alphabetically sorted List...
    'Also make sure that the ListBox has (at least) a vertical scroll bar...
    With List1
        .AddItem "cat"
        .AddItem "dog"
        .AddItem "doggy"
        .AddItem "Elephant"
        .AddItem "elephant"
        .AddItem "lion"
        .AddItem "2 little lions"
        .AddItem "5 little lions"
        .AddItem "tiger"
        .AddItem "The Rock Show :)"
        .AddItem "an option"
        .AddItem "another option"
        .AddItem "guess what, another option"
        .AddItem "eagle"
        .AddItem "...these are just examples"
        .AddItem "zzzzZZZZZ..."
        .AddItem "it works!"
        .AddItem "system overload"
        .AddItem "are you looking for this?"
        .AddItem "my inspiration has left me..."
    End With
    'Let's make the first item highlight (selected):
    List1.ListIndex = 0
End Sub

Private Sub List1_DblClick()
    'Double clicking on an item of the List will cause the
    'information about the item to appear at the Textarea on the right.
    'You can add your code here for the method how you'll load the infos
    'about the selected item. What here happens is just for the example...
    Text2.Text = "Here will appear the information about the selected item: " & vbCrLf & List1.Text
End Sub

Private Sub Text1_Change()
    Dim i As Integer                            'a counter
    Dim strSearchText As String                 'the text to search for
    'To prevent the situation that the first option will be highlighted,
    'when you delete what you entered in the text area... :
    If Text1.Text = "" Then Exit Sub
    'It begins... :
    For i = 0 To List1.ListCount - 1            'We go all the List through...
        strSearchText = Mid(List1.List(i), 1, Len(Text1.Text))
        If Text1.Text Like strSearchText Then   'We found it ;-)
            List1.ListIndex = i                 'The List "springs" to that item...
            Exit For                            'We don't need to search anymore
        End If
    Next
End Sub
