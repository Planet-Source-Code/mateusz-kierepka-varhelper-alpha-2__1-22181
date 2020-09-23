VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FastCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'*********************************************************************************************
'
' FastCollection Class
'
'*********************************************************************************************
'
' Author: Eduardo A. Morcillo
' E-Mail: e_morcillo@yahoo.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Distribution: You can freely use this code in your own applications but you
'               can't publish this code in a web site, online service, or any
'               other media, without my express permission.
'
' Usage: at your own risk.
'
' Tested on: Windows 98
'
' History:
'          05/28/2000 * 2 bugs fixed:
'                         o pvGetItemByKey did not return the
'                           correct pointer.
'                         o pvRemoveKey and RemoveByKey fail
'                           with some keys.
'
'          05/27/2000 * Changes:
'                         o By popular demand you can now
'                           use numbers as keys.
'                         o The items array is accesed through
'                           a non-vb allocated SAFEARRAY.
'                         o The hash table was converted to a
'                           non-VB allocated SAFEARRAY
'                         o New methods:
'                           ItemByIndex
'                           ItemByKey
'                           RemoveByKey
'                           RemoveByIndex
'                         o The INCREMENT constant
'                           was converted to a property.
'          05/15/2000 * Changes:
'                         o KeyFromIndex function was removed.
'                         o Clear performance improved.
'                     * New features:
'                         o Key property. Use it to get/set
'                           item keys.
'                         o Keys function. It returns all
'                           keys in a string array.
'                         o Item property can be set.
'                         o CompareMode property. Use it to
'                           set how keys are compared.
'                     * Bugs Fixed:
'                         o Keys can be numbers. Now if you try
'                           to use a number as key the add method
'                           will raise the error 5. The same error
'                           will occurs if you try to set the
'                           key to a number using the Key property.
'
'          04/28/2000 * Fixed a bug on the Add method that
'                       didn't m_Increment Object reference
'                       count causing objects to be destroyed.
'
'          04/27/2000 * Fixed a bug on the Add method: the
'                       items on the array were not moved
'                       when the before/after parameters
'                       contains a key.
'
'          04/26/2000 * Fixed a bug in the addition
'                       of keys to the linked list.
'
'          04/25/2000 * The class was released.
'
'*********************************************************************************************
Option Explicit

' ==== public variables (properties) ====

' Setting this property to True
' items can be replaced.
Public AllowUpdates As Boolean

' ==== Private Constants ====

' Number of items in the key table
' If you increase this number the
' collection will run faster with
' keys, but will require more
' memory (the size of the hash
' table is 4 * MAXHASH bytes).
Const MAXHASH As Long = 2048

'
' Linked list item format
'
' Type LinkedListItem
'    Next As Long
'    Item As Long
'    Key  As Long
' End Type
'
' Offsets on the linked list items
Const LL_NEXT     As Long = 0       ' Next item in the list
Const LL_ITEMPTR  As Long = 4    ' Pointer to the item
Const LL_KEYPTR   As Long = 8     ' Pointer to the key

' Size of liked list items in bytes
Const LL_ITEMSIZE As Long = 12

'
' Item format
'
' Type Item
'     Value As Variant
'     Key   As Long
' End Type
'
' Offsets on the collection items
Const CI_VARIANT  As Long = 0
Const CI_KEYPTR   As Long = 16

' Size of each item in bytes
Const CI_ITEMSIZE As Long = 20

' Name of this class. Used to raise errors
Const CLASSNAME = "FastCollection."

' ==== public Constants ====

Enum CompareModes
  CompareBinary
  CompareText
End Enum

' ==== Private variables ====

' Items table
Private m_Items() As Long                 ' Array of items
Private m_ItemsSA As SAFEARRAY_1D         ' SAFEARRAY of items

' Hash table
Private m_HashTable() As Long             ' Table of keys
Private m_HashSA As SAFEARRAY_1D          ' SAFEARRAY of the hash Table

' Other variables
Private m_ArraySize As Long               ' Size in bytes of the items array
Private m_Count As Long                   ' Number of items in the collection
Private m_hHeap As Long                   ' Handle of the heap object
Private m_CompareMode As CompareModes
Private m_CompareFlags As Long
Private m_Increment As Long

' ==== API Declarations ====

Const LOCALE_USER_DEFAULT = &H400&

Private Declare Function LHashValOfNameSys Lib "oleaut32" (ByVal syskind As Long, ByVal lcid As Long, ByVal szName As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal wstr As Long) As Long

Private Declare Function SysStringLen Lib "oleaut32" (ByVal bstr As Any) As Long
Private Declare Function SysStringByteLen Lib "oleaut32" (ByVal bstr As Any) As Long

Private Enum CompareResults
  VARCMP_LT = 0
  VARCMP_EQ = 1
  VARCMP_GT = 2
  VARCMP_NULL = 3
End Enum

Private Enum VarBstrCmpFlags
  NORM_IGNORECASE = &H1        ' Ignore case.
  NORM_IGNORENONSPACE = &H2    ' Ignore nonspacing characters.
  NORM_IGNORESYMBOLS = &H4     ' Ignore symbols.
  NORM_IGNOREWIDTH = &H8       ' Ignore string width.
  NORM_IGNOREKANATYPE = &H40   ' Ignore Kana type.
  NORM_IGNOREKASHIDA = &H40000 ' Ignore Arabic kashida characters.
End Enum

Private Declare Function VarBstrCmp Lib "oleaut32" ( _
                ByVal bstrLeft As Long, _
                ByVal bstrRight As Long, _
                ByVal lcid As Long, _
                ByVal dwFlags As VarBstrCmpFlags) As CompareResults

Private Sub Class_Initialize()

 Const PROCNAME = "Initialize"
 Dim lPtr As Long

  ' Initialize the m_Increment property
  m_Increment = 4000

  ' Create a Heap object
  m_hHeap = mdlIEnumVariant.HeapCreate(0, 0, 0) ' 2 * m_Increment + 4 * MAXHASH + 100, 0)

  If m_hHeap Then

    ' Allocate memory for the items array
    lPtr = mdlIEnumVariant.HeapAlloc(m_hHeap, HEAP_ZERO_MEMORY, m_Increment)

    If lPtr Then

      ' Fill the SAFEARRAY structure
      With m_ItemsSA
        .Bounds(0).cElements = 0   ' 0 elements
        .Bounds(0).lLbound = 1     ' 1 based index
        .cDims = 1
        .cbElements = 4            ' each element has 4 bytes
        .pvData = lPtr
      End With

      ' Set the array size in bytes
      m_ArraySize = m_Increment

      ' Copy a pointer to m_ItemsSA to
      ' the m_Items array.
      mdlIEnumVariant.MoveMemory ByVal VarPtrArray(m_Items, m_Items), VarPtr(m_ItemsSA), 4

      ' Allocate the Hash table
      lPtr = mdlIEnumVariant.HeapAlloc(m_hHeap, HEAP_ZERO_MEMORY, MAXHASH * 4)

      If lPtr Then

        ' Fill the SAFEARRAY structure
        With m_HashSA
          .Bounds(0).cElements = MAXHASH   ' MAXHASH elements
          .Bounds(0).lLbound = 0           ' 0 based index
          .cDims = 1
          .cbElements = 4                  ' each element has 4 bytes
          .pvData = lPtr
        End With

        ' Put a pointer to m_ItemsSA in
        ' the m_HashTable array
        mdlIEnumVariant.MoveMemory ByVal VarPtrArray(m_HashTable, m_HashTable), VarPtr(m_HashSA), 4

       Else

        ' Destroy the heap object
        HeapDestroy m_hHeap

        Err.Raise 7, CLASSNAME & PROCNAME

      End If

     Else

      ' Destroy the heap object
      HeapDestroy m_hHeap

      Err.Raise 7, CLASSNAME & PROCNAME

    End If

   Else

    Err.Raise 7, CLASSNAME & PROCNAME

  End If
   
End Sub

Private Sub Class_Terminate()

 ' Remove all items

  pvClear

End Sub

':) Ulli's Code Formatter V2.0 (2001-01-23 10:53:05) 181 + 1234 = 1415 Lines

Public Sub Add(Item As Variant, Optional Key As String, Optional Before As Variant, Optional After As Variant)

 Const PROCNAME = "Add"  ' Procedure name
 Dim lpArray As Long     ' Pointer to the first element of the array
 Dim lpItem As Long      ' Pointer to the item
 Dim lpKey As Long       ' Pointer to key BSTR
 Dim lIdx As Long        ' Index of the new item

  ' Check if the array has
  ' to be expanded
  If m_ArraySize < ((m_Count + 1) * 4) Then

    ' m_Increment the array size
    m_ArraySize = m_ArraySize + m_Increment

    With m_ItemsSA

      ' Reallocate the array memory
      lpArray = mdlIEnumVariant.HeapReAlloc(m_hHeap, 0, .pvData, m_ArraySize)

      ' Update the array pointer
      If lpArray Then
        .pvData = lpArray
       Else
        Err.Raise 7, CLASSNAME & PROCNAME
      End If

    End With

  End If

  ' Allocate memory for the item
  lpItem = mdlIEnumVariant.HeapAlloc(m_hHeap, HEAP_ZERO_MEMORY, CI_ITEMSIZE)

  If lpItem Then

    ' Check if the item has to be
    ' added to a specific position
    If Not IsMissing(Before) Then

      If IsNumeric(Before) Then

        ' Get the item index
        lIdx = Before - 1

        ' Check if the index is valid
        If lIdx >= 0 And lIdx < m_Count Then

          If (m_Count - lIdx) Then

            ' Move the items on the array
            ' to insert the new one
            mdlIEnumVariant.MoveMemory _
                                       ByVal m_ItemsSA.pvData + 4 * (lIdx + 1), _
                                       ByVal m_ItemsSA.pvData + 4 * lIdx, _
                                       4 * (m_Count - lIdx)
          End If

         Else

          GoTo InvalidParameterError

        End If

       Else

        ' Get the item index
        lIdx = IndexFromKey(Before) - 1

        If lIdx >= 0 Then

          If (m_Count - lIdx) Then

            ' Move the items on the array
            mdlIEnumVariant.MoveMemory _
                                       ByVal m_ItemsSA.pvData + 4 * (lIdx + 1), _
                                       ByVal m_ItemsSA.pvData + 4 * lIdx, _
                                       4 * (m_Count - lIdx)
          End If

         Else

          GoTo InvalidParameterError

        End If

      End If

     ElseIf Not IsMissing(After) Then

      If IsNumeric(After) Then

        ' Set the item index
        lIdx = After

        ' Check if the index is valid
        If lIdx >= 0 And lIdx < m_Count Then

          If (m_Count - lIdx) Then

            ' Move the items on the array
            mdlIEnumVariant.MoveMemory _
                                       ByVal m_ItemsSA.pvData + 4 * (lIdx + 1), _
                                       ByVal m_ItemsSA.pvData + 4 * lIdx, _
                                       4 * (m_Count - lIdx)
          End If

         Else

          GoTo InvalidParameterError

        End If

       Else

        ' Set the item index
        lIdx = IndexFromKey(After)

        If lIdx <> -1 Then

          If (m_Count - lIdx) Then

            ' Move the items on the array
            mdlIEnumVariant.MoveMemory _
                                       ByVal m_ItemsSA.pvData + 4 * (lIdx + 1), _
                                       ByVal m_ItemsSA.pvData + 4 * lIdx, _
                                       4 * (m_Count - lIdx)
          End If

         Else

          GoTo InvalidParameterError

        End If

      End If

     Else

      ' Set the item index
      lIdx = m_Count

    End If

    ' Add the item to the key list
    If Len(Key) > 0 Then

      If pvGetItemByKey(Key, lpItem) Then

        ' Release the memory used by the item
        mdlIEnumVariant.HeapFree m_hHeap, 0, lpItem

        ' The key is already used
        Err.Raise 457, CLASSNAME & PROCNAME

      End If

    End If

    ' Copy the item to the allocated memory
    mdlIEnumVariant.VariantCopyIndPtrVar lpItem, Item

    ' m_Increment the count
    m_Count = m_Count + 1

    ' m_Increment the array element count
    m_ItemsSA.Bounds(0).cElements = m_Count

    ' Copy the pointer to the array
    m_Items(lIdx + 1) = lpItem

   Else

    Err.Raise 7, CLASSNAME & PROCNAME

  End If

Exit Sub

InvalidParameterError:

  ' Release the item memory
  mdlIEnumVariant.HeapFree m_hHeap, 0, lpItem

  ' Raise the error
  Err.Raise 5, CLASSNAME & PROCNAME

End Sub


Public Sub Clear()

 ' Clear the items

  pvClear

  ' Reinitialize the collection
  Class_Initialize

End Sub

'
' CompareMode
'
' Returns the key compare mode
'
Public Property Get CompareMode() As CompareModes
Attribute CompareMode.VB_Description = "Returns/Sets the mode used to compare keys."

  CompareMode = m_CompareMode

End Property

'
' CompareMode
'
' Sets the key compare mode
'
Public Property Let CompareMode(ByVal NewMode As CompareModes)

 Const PROCNAME = "CompareMode (Let)"

  ' CompareMode can only
  ' be changed if the
  ' collection does not have
  ' items.
  If m_Count = 0 Then

    m_CompareMode = NewMode

    Select Case m_CompareMode
     Case CompareText
      m_CompareFlags = NORM_IGNORECASE
     Case Else
      m_CompareFlags = 0
    End Select

   Else

    Err.Raise 387, CLASSNAME & PROCNAME, "Cannot change CompareMode. The collection contains items."

  End If

End Property

'
' Count
'
' Returns the number of items in the collection
'
Public Property Get Count() As Long

  Count = m_Count

End Property

'
' Increment
'
' Sets the quantity of bytes in which the array
' of items will be increased
'
Public Property Get Increment() As Long

  Increment = m_Increment

End Property

'
' Increment
'
' Sets the quantity of bytes in which the array
' of items will be increased
'
Public Property Let Increment(ByVal NewIncrement As Long)

 ' Round to a multiple of 4

  m_Increment = NewIncrement \ 4 * 4

  If m_Increment < 4 Then m_Increment = 4

End Property

'
' IndexFromKey
'
' Returns the index of an item given its key or -1
' if the key was not found
'
Public Function IndexFromKey(ByVal Key As String) As Long

 Dim lIdx As Long, lPtr As Long, lPtr2 As Long

  ' Find the key in the keys table
  lPtr = pvGetItemByKey(Key)

  If lPtr Then

    ' The item was found. Iterate
    ' the items array to find the
    ' same Ptr and return that index.

    For lIdx = 1 To m_Count

      ' Compare pointers
      If lPtr = m_Items(lIdx) Then Exit For

    Next lIdx

    IndexFromKey = lIdx

   Else

    IndexFromKey = -1

  End If

End Function

Public Property Get Item(ByVal Key As Variant) As Variant
Attribute Item.VB_Description = "Returns/Sets the item value."

 Const PROCNAME = "Item (Get)"
 Dim lIdx As Long, lPtr As Long

  Select Case VarType(Key)

   Case vbString

    lPtr = pvGetItemByKey(Key)

    If lPtr <> 0 Then
      ' Return a copy of the
      ' stored variant
      mdlIEnumVariant.VariantCopyVarPtr Item, lPtr

     Else

      Err.Raise 381, CLASSNAME & PROCNAME

    End If

   Case vbByte, vbInteger, vbLong, vbSingle, vbDouble

    ' Convert the variant to a Long
    lIdx = CLng(Key)

    ' Raise an error if the index
    ' is invalid
    If lIdx < 1 Or lIdx > m_Count Then Err.Raise 381, CLASSNAME & PROCNAME

    ' Copy the item
    mdlIEnumVariant.VariantCopyVarPtr Item, m_Items(lIdx)

   Case Else

    ' Raise an error
    Err.Raise 5, CLASSNAME & PROCNAME

  End Select

End Property

Public Property Let Item(ByVal Key As Variant, ByVal NewValue As Variant)

 Const PROCNAME = "Item (Let)"
 Dim lIdx As Long, lPtr As Long

  If AllowUpdates Then

    Select Case VarType(Key)

     Case vbString

      ' Find the item by its key
      lPtr = pvGetItemByKey(Key)

      ' Replace the item
      mdlIEnumVariant.VariantCopyIndPtrVar lPtr, NewValue

     Case vbByte, vbInteger, vbLong, vbSingle, vbDouble

      lIdx = CLng(Key)

      If lIdx < 1 Or lIdx > m_Count Then Err.Raise 381, CLASSNAME & PROCNAME

      ' Replace the item
      mdlIEnumVariant.VariantCopyIndPtrVar m_Items(lIdx), NewValue

     Case Else

      Err.Raise 5, CLASSNAME & PROCNAME

    End Select

   Else

    ' Set not supported (read-only property)
    Err.Raise 383, CLASSNAME & PROCNAME

  End If

End Property

Public Property Set Item(ByVal Key As Variant, ByVal NewValue As Variant)

 Const PROCNAME = "Item (Set)"
 Dim lIdx As Long, lPtr As Long

  If AllowUpdates Then

    Select Case VarType(Key)

     Case vbString

      ' Find the item by its key
      lPtr = pvGetItemByKey(Key)

      ' Replace the item
      mdlIEnumVariant.VariantCopyIndPtrVar lPtr, NewValue

     Case vbByte, vbInteger, vbLong, vbSingle, vbDouble

      lIdx = CLng(Key)

      If lIdx < 1 Or lIdx > m_Count Then Err.Raise 381, CLASSNAME & PROCNAME

      ' Replace the item
      mdlIEnumVariant.VariantCopyIndPtrVar m_Items(lIdx), NewValue

     Case Else

      Err.Raise 5, CLASSNAME & PROCNAME

    End Select

   Else

    ' Set not supported (read-only property)
    Err.Raise 383, CLASSNAME & PROCNAME

  End If

End Property

'
' ItemByIndex
'
' Returns an item given its index.
'
Public Function ItemByIndex(ByVal Index As Long) As Variant

 Const PROCNAME = "ItemByIndex"

  ' Raise an error if the index
  ' is invalid. You can remove
  ' this line and get the error
  ' raised by VB.
  If Index < 1 Or Index > m_Count Then Err.Raise 381, CLASSNAME & PROCNAME

  ' Copy the item
  mdlIEnumVariant.VariantCopyVarPtr ItemByIndex, m_Items(Index)

End Function

'
' ItemByKey
'
' Returns an item given its key
'
Public Function ItemByKey(Key As String) As Variant

 Const PROCNAME = "ItemByKey"
 Dim lPtr As Long

  lPtr = pvGetItemByKey(Key)

  If lPtr <> 0 Then

    ' Return a copy of the
    ' stored variant
    mdlIEnumVariant.VariantCopyVarPtr ItemByKey, lPtr

   Else

    Err.Raise 381, CLASSNAME & PROCNAME

  End If

End Function

'
' Get Key
'
' Returns the key used by the item of the given index
'
Public Property Get Key(ByVal Index As Long) As String
Attribute Key.VB_Description = "Returns an item key given its index."

 Const PROCNAME = "Key (Get)"
 Dim lpKey As Long
 Dim lKeyLen As Long
 Dim lpItem As Long

  If Index < 1 Or Index > m_Count Then Err.Raise 381, CLASSNAME & PROCNAME

  ' Get the pointer to the key string
  mdlIEnumVariant.MoveMemory lpKey, ByVal m_Items(Index) + CI_KEYPTR, 4

  ' lpKey is <> 0 if the
  ' item has a key
  If lpKey Then

    ' Get the key length
    lKeyLen = SysStringLen(lpKey)

    ' Copy the string
    Key = Space$(lKeyLen)
    mdlIEnumVariant.MoveMemory ByVal StrPtr(Key), ByVal lpKey, lKeyLen * 2

  End If

End Property

'
' Let Key
'
' Returns the key used by the item of the given index
'
Public Property Let Key(ByVal Index As Long, ByVal NewKey As String)

 Const PROCNAME = "Key (LET)"
 Dim lpKey As Long
 Dim lpItem As Long

  ' Check if the index is valid
  If Index < 1 Or Index > m_Count Then Err.Raise 381, CLASSNAME & PROCNAME

  ' Get the pointer to the key string
  mdlIEnumVariant.MoveMemory lpKey, ByVal m_Items(Index) + CI_KEYPTR, 4

  If Len(NewKey) = 0 Then

    ' The new key is empty so
    ' just remove the old
    pvRemoveKey lpKey

   Else

    ' Try to add the new key
    If pvGetItemByKey(NewKey, lpItem) = 0 Then

      ' The new key was added to the
      ' liked list. Now remove the old key.
      pvRemoveKey lpKey

     Else

      ' The new key is used by other item
      Err.Raise 457, CLASSNAME & PROCNAME

    End If

  End If

End Property

'
' KeyExists
'
' Returns if a given key is already used to
' reference an item in the collection
'
Public Function KeyExists(ByVal Key As String) As Boolean

  KeyExists = pvGetItemByKey(Key) <> 0

End Function

'
' Keys
'
' Returns a string array containing all the keys
' used by items. The array is not sorted.
'
Public Function Keys() As Variant

 Const PROCNAME = "Keys"
 Dim lIdx As Long
 Dim lpKey As Long
 Dim lKeyIdx As Long
 Dim aKeys() As String

  For lIdx = 1 To m_Count

    ' Get the pointer to the key
    mdlIEnumVariant.MoveMemory lpKey, ByVal m_Items(lIdx) + CI_KEYPTR, 4

    If lpKey Then

      ReDim aKeys(0 To lKeyIdx)

      ' Copy the string
      'aKeys(lKeyIdx) = SysAllocString(lpKey)

    End If

  Next lIdx

  ' Return the array
  Keys = aKeys

End Function

'
' NewEnum
'
' Returns an enumerator object which
' used by For Each.
'
' THIS FUNCTION MUST HAVE THE Procedure ID = -4.
' (change it in the Tools/Procedure Attributes dialog)
'
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

 ' Create the lightweight IEnumVARIANT object

  Set NewEnum = CreateIEnumVARIANT(m_hHeap, m_ItemsSA)

End Function

'
' pvClear
'
' Removes all items and releases the memory used
' by the collection
'
Private Sub pvClear()

 Dim lIdx As Long

  ' Clear all variants on the
  ' array and free the memory
  For lIdx = 1 To m_Count
    ' Clear the variant
    mdlIEnumVariant.VariantClear m_Items(lIdx)
  Next lIdx

  ' Reset the item count
  m_Count = 0

  ' Clear the items SAFEARRAY (= Erase m_Items)
  m_ItemsSA.Bounds(0).cElements = 0
  m_ItemsSA.pvData = 0

  ' Clear the hash table SAFEARRAY (= Erase m_HashTable)
  m_HashSA.Bounds(0).cElements = 0
  m_HashSA.pvData = 0

  ' Destroy the heap object. This
  ' will release all memory allocated
  ' with HeapAlloc (including the keys)
  mdlIEnumVariant.HeapDestroy m_hHeap

End Sub

'
' pvGetItemByKey
'
' Returns a pointer to the item given its key.
'
' Parameters:
' ----------
' Key       - Key of the item to search or add
' lpNewItem - Optional pointer to a item to add.
'             If lpNewItem <> 0 and the key is
'             not found the key is added.
'
Private Function pvGetItemByKey(ByVal Key As String, _
                                Optional ByVal lpNewItem As Long) As Long

 Dim lHash As Long
 Dim lpItem As Long
 Dim lpKey1 As Long
 Dim lpKey2 As Long
 Dim lpNext As Long
 Dim bFound As Long

  ' Convert the key to
  ' lower case if CompareMode
  ' is set to CompareText
  If m_CompareMode = CompareText Then Key = LCase$(Key)

  ' Get the string pointer
  lpKey1 = StrPtr(Key)

  ' Calculate a hash value to use
  ' as index on keys table
  lHash = LHashValOfNameSys(0, 0, lpKey1) Mod MAXHASH

  If m_HashTable(lHash) = 0 Then

    ' If a pointer to an item
    ' was passed to the function
    ' add it to the list
    If lpNewItem Then

      ' Allocate memory for the item
      lpItem = mdlIEnumVariant.HeapAlloc(m_hHeap, HEAP_ZERO_MEMORY, LL_ITEMSIZE)

      ' Copy the item pointer
      mdlIEnumVariant.MoveMemory ByVal lpItem + LL_ITEMPTR, lpNewItem, 4

      ' Allocate memory and copy the key
      lpKey2 = pvSysAllocString(lpKey1)

      ' Copy the string pointer to the item
      mdlIEnumVariant.MoveMemory ByVal lpNewItem + CI_KEYPTR, lpKey2, 4

      ' Copy the string pointer to the LL item
      mdlIEnumVariant.MoveMemory ByVal lpItem + LL_KEYPTR, lpKey2, 4

      ' Copy the pointer to the table
      m_HashTable(lHash) = lpItem

    End If

   Else

    ' Get the pointer to the first item
    lpNext = m_HashTable(lHash)

    Do

      lpItem = lpNext

      ' Get the pointer to the key string
      mdlIEnumVariant.MoveMemory lpKey2, ByVal lpItem + LL_KEYPTR, 4

      ' Compare the keys length.
      If Len(Key) = SysStringLen(lpKey2) Then

        ' Compare the keys
        If VarBstrCmp(lpKey1, lpKey2, LOCALE_USER_DEFAULT, m_CompareFlags) = VARCMP_EQ Then

          ' Both keys are equal
          bFound = True
          Exit Do

        End If

      End If

      ' Go with the next item
      mdlIEnumVariant.MoveMemory lpNext, ByVal lpItem + LL_NEXT, 4

    Loop Until lpNext = 0

    If bFound Then

      ' Get the item pointer
      mdlIEnumVariant.MoveMemory pvGetItemByKey, ByVal lpItem + LL_ITEMPTR, 4

     Else

      ' Add the item
      If lpNewItem Then

        ' Allocate memory for the item
        lpNext = mdlIEnumVariant.HeapAlloc(m_hHeap, HEAP_ZERO_MEMORY, LL_ITEMSIZE)

        ' Set the Next pointer in the previous item
        mdlIEnumVariant.MoveMemory ByVal lpItem, lpNext, 4

        ' Copy the item pointer
        mdlIEnumVariant.MoveMemory ByVal lpNext + LL_ITEMPTR, lpNewItem, 4

        ' Allocate memory and copy the key
        lpKey2 = pvSysAllocString(lpKey1)

        ' Copy the string pointer to the item
        mdlIEnumVariant.MoveMemory ByVal lpNewItem + CI_KEYPTR, lpKey2, 4

        ' Copy the string pointer to the LL item
        mdlIEnumVariant.MoveMemory ByVal lpNext + LL_KEYPTR, lpKey2, 4

      End If

    End If

  End If

End Function


'
' pvRemoveKey
'
' Removes a key from the key list
'
' Parameters:
' -----------
' lpKey - Pointer to the key
'
Private Sub pvRemoveKey(ByVal lpKey As Long)

 Dim lHash As Long
 Dim lpItem As Long
 Dim lpPrev As Long
 Dim lpNext As Long
 Dim lpKey2 As Long
 Dim bAreEqual As Boolean

  ' Calculate the array index
  lHash = LHashValOfNameSys(0, 0, lpKey) Mod MAXHASH

  ' Get the pointer to the
  ' first element in the linked List
  lpItem = m_HashTable(lHash)

  Do

    ' Get the key pointer stored
    ' in the item
    mdlIEnumVariant.MoveMemory lpKey2, ByVal lpItem + LL_KEYPTR, 4

    ' Compare the strings pointers
    If lpKey = lpKey2 Then

      bAreEqual = True

     ElseIf SysStringLen(lpKey2) = SysStringLen(lpKey) Then

      ' If the string length are equal
      ' compare the strings
      bAreEqual = VarBstrCmp(lpKey2, lpKey, LOCALE_USER_DEFAULT, m_CompareFlags) = VARCMP_EQ

    End If

    If bAreEqual Then

      ' Move the pointer to the
      ' previous item
      If lpPrev Then

        ' Get the pointer to the next
        ' element
        mdlIEnumVariant.MoveMemory ByVal lpPrev + LL_NEXT, ByVal lpItem + LL_NEXT, 4

       Else

        mdlIEnumVariant.MoveMemory m_HashTable(lHash), ByVal lpItem + LL_NEXT, 4

      End If

      ' Release the string
      pvSysFreeString lpKey

      ' Release the memory
      ' used by the key item
      mdlIEnumVariant.HeapFree m_hHeap, 0, lpItem

      ' Get out of the loop
      Exit Do

    End If

    ' Store this item pointer
    lpPrev = lpItem

    ' Go to the next item
    mdlIEnumVariant.MoveMemory lpItem, ByVal lpItem + LL_NEXT, 4

  Loop Until lpItem = 0

End Sub

'
' pvSysAllocString
'
' Allocates a BSTR in the collection heap
'
' Parameters:
' -----------
' LPOLESTR - Pointer to the source string
'
' Returns: A pointer to the new string
'
Private Function pvSysAllocString(ByVal LPOLESTR As Long) As Long

 Dim lByteLen As Long

  ' Get the string lenght in bytes
  lByteLen = SysStringByteLen(LPOLESTR)

  ' Allocate the memory in the heap
  pvSysAllocString = mdlIEnumVariant.HeapAlloc(m_hHeap, HEAP_ZERO_MEMORY, lByteLen + 6)

  ' Copy the string at the
  ' begining of buffer
  mdlIEnumVariant.MoveMemory ByVal pvSysAllocString, lByteLen, 4

  ' Move the pointer so it points to
  ' the first byte after the size
  pvSysAllocString = pvSysAllocString + 4

  ' Copy the string
  mdlIEnumVariant.MoveMemory ByVal pvSysAllocString, ByVal LPOLESTR, lByteLen

End Function

'
' pvSysFreeString
'
' Releases a BSTR allocated with pvSysAllocString
'
' Parameters:
' ----------
' LPOLESTR - Pointer to the string to release
'
Private Sub pvSysFreeString(ByVal LPOLESTR As Long)

  mdlIEnumVariant.HeapFree m_hHeap, 0, LPOLESTR - 4

End Sub

'
' Remove
'
' Removes an item from the collection given
' its key or index
'
' Parameters:
' -----------
' Key - Key or index of the item to remove
'
Public Sub Remove(ByVal Key As Variant)

 Const PROCNAME = "Remove"
 Dim lIdx As Long
 Dim lpArrayItem As Long
 Dim lpItem As Long
 Dim lpKey As Long
 Dim lUbound As Long

  Select Case VarType(Key)

   Case vbString

    RemoveByKey Key

   Case vbByte, vbInteger, vbLong, vbSingle, vbDouble

    RemoveByIndex Key

   Case Else

  End Select

End Sub

'
' RemoveByIndex
'
' Removes an item from the collection given its index
'
' Parameters:
' -----------
' Index - Index of the item to remove
'
Public Sub RemoveByIndex(ByVal Index As Long)

 Const PROCNAME = "RemoveByIndex"
 Dim lpKey As Long, lPtr As Long

  ' Check if Index is valid
  If Index < 1 Or Index > m_Count Then Err.Raise 381, CLASSNAME & PROCNAME

  ' Get the pointer to the item
  lPtr = m_Items(Index)

  ' Clear the item
  mdlIEnumVariant.VariantClear lPtr

  ' Get the item key
  mdlIEnumVariant.MoveMemory lpKey, ByVal lPtr + CI_KEYPTR, 4

  ' Remove the key
  If lpKey Then pvRemoveKey lpKey

  ' Destroy the item
  mdlIEnumVariant.HeapFree m_hHeap, 0, lPtr

  ' Move the items on the array
  If Index < m_Count Then
    MoveMemory _
               ByVal m_ItemsSA.pvData + 4 * (Index - 1), _
               ByVal m_ItemsSA.pvData + 4 * Index, _
               4 * (m_Count - Index)
  End If

  ' Decrement the items count
  m_Count = m_Count - 1

  ' Decrement the array elements count
  m_ItemsSA.Bounds(0).cElements = m_Count

  ' Release part of the
  ' table if there're too many
  ' items allocated but not
  ' used.
  If (m_Count * 4) < m_ArraySize - (m_Increment * 1.5) Then
    m_ItemsSA.pvData = HeapReAlloc(m_hHeap, 0, m_ItemsSA.pvData, m_ArraySize - m_Increment)
    m_ArraySize = m_ArraySize - m_Increment
  End If

End Sub


'
' RemoveByKey
'
' Removes an item from the collection given its key
'
' Parameters:
' -----------
' Key - Key of the item to remove
'
Public Sub RemoveByKey(ByVal Key As String)

 Const PROCNAME = "RemoveByKey"
 Dim lHash As Long
 Dim lpItem As Long
 Dim lpKey1 As Long
 Dim lpKey2 As Long
 Dim lpPrev As Long
 Dim lpNext As Long
 Dim lIdx As Long
 Dim bFound As Long

  ' Convert the key to
  ' lower case if CompareMode
  ' is set to CompareText
  If m_CompareMode = CompareText Then Key = LCase$(Key)

  ' Get the string pointer
  lpKey1 = StrPtr(Key)

  ' Calculate a hash value to use
  ' as index on keys table
  lHash = LHashValOfNameSys(0, 0, lpKey1) Mod MAXHASH
   
  If m_HashTable(lHash) Then
      
    ' Get the pointer to the first item
    lpNext = m_HashTable(lHash)

    Do

      ' Get the pointer to the key string
      mdlIEnumVariant.MoveMemory lpKey2, ByVal lpNext + LL_KEYPTR, 4

      ' Compare the keys length.
      If Len(Key) = SysStringLen(lpKey2) Then

        ' Compare the keys
        If VarBstrCmp(lpKey1, lpKey2, LOCALE_USER_DEFAULT, m_CompareFlags) = VARCMP_EQ Then

          ' Release the key
          pvSysFreeString lpKey2

          ' Get the pointer to the item
          mdlIEnumVariant.MoveMemory lpItem, ByVal lpNext + LL_ITEMPTR, 4

          If lpPrev Then
               
            ' Move the Next pointer to the previous key item
            mdlIEnumVariant.MoveMemory ByVal lpPrev + LL_NEXT, ByVal lpNext + LL_NEXT, 4
                  
           Else
               
            ' Move the Next pointer to the hash table
            mdlIEnumVariant.MoveMemory m_HashTable(lHash), ByVal lpNext + LL_NEXT, 4
                  
          End If

          ' Release the key item
          mdlIEnumVariant.HeapFree m_hHeap, 0, lpNext

          For lIdx = 1 To m_Count \ 2

            If m_Items(lIdx) = lpItem Then
              lIdx = lIdx
              Exit For
             ElseIf m_Items(m_Count - lIdx) = lpItem Then
              lIdx = m_Count - lIdx
              Exit For
             Else

            End If

          Next lIdx

          ' Move the items
          MoveMemory _
                     ByVal m_ItemsSA.pvData + 4 * (lIdx - 1), _
                     ByVal m_ItemsSA.pvData + 4 * lIdx, _
                     4 * (m_Count - lIdx)

          ' Clear the item
          mdlIEnumVariant.VariantClear lpItem

          ' Destroy the item
          mdlIEnumVariant.HeapFree m_hHeap, 0, lpItem

          ' Decrement the items count
          m_Count = m_Count - 1

          ' Decrement the array elements count
          m_ItemsSA.Bounds(0).cElements = m_Count

          If (m_Count * 4) < m_ArraySize - (m_Increment * 1.5) Then
            m_ItemsSA.pvData = HeapReAlloc(m_hHeap, 0, m_ItemsSA.pvData, m_ArraySize - m_Increment)
            m_ArraySize = m_ArraySize - m_Increment
          End If

          bFound = True
          Exit Do

        End If

      End If

      lpPrev = lpNext

      ' Go with the next item
      mdlIEnumVariant.MoveMemory lpNext, ByVal lpNext + LL_NEXT, 4

    Loop Until lpNext = 0

  End If

  If Not bFound Then
      
    ' Invalid argument (the key not exists)
    Err.Raise 5, CLASSNAME & PROCNAME

  End If

End Sub

