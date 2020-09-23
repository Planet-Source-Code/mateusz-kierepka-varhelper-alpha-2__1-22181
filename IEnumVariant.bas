Attribute VB_Name = "mdlIEnumVariant"
'*********************************************************************************************
'
' FastCollection Class
'
' IEnumVARIANT light-weight object
'
'*********************************************************************************************
'
' Author: Eduardo Morcillo
' E-Mail: edanmo@geocities.com
' Web Page: http://www.domaindlx.com/e_morcillo
'
' Distribution: You can freely use this code in your own applications but you
'               can't publish this code in a web site, online service, or any
'               other media,  without my express permission.
'
' Usage: at your own risk.
'
' Tested on: Windows 98
'
' History:
'          05/27/2000 * The object uses the
'                       SAFEARRAY from the
'                       FastCollection class.
'          04/26/2000 * Fixed a bug on the Next_
'                       method that which didn't
'                       return the item.
'          04/25/2000 * Code was released
'
'*********************************************************************************************
Option Explicit

Private Type IEnumVARIANT  ' Object struct
  vtable As Long       ' Pointer to vtable
  RefCount As Long     ' Reference count
  hHeap As Long        ' Handle of heap object used to create the object
  Items() As Long      ' Array of items
  MaxIdx As Long       ' Number of items
  CurrentIndex As Long ' Current index
End Type

Private Type UUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0 To 7) As Byte
End Type

Const sIID_IUnknown = "{00000000-0000-0000-C000-000000000046}"
Const sIID_IEnumVARIANT = "{00020404-0000-0000-C000-000000000046}"

Dim IID_IUnknown As UUID
Dim IID_IEnumVARIANT As UUID

' ==== API Declarations ====

Type SAFEARRAYBOUND

  cElements As Long      ' Element count
  lLbound As Long        ' LBound
End Type

Type SAFEARRAY_1D

  cDims As Integer       ' Number of dimensions
  fFeatures As Integer   ' Flags
  cbElements As Long     ' Length of each element
  cLocks As Long         ' Lock count
  pvData As Long         ' Pointer to the data
  Bounds(0 To 0) As SAFEARRAYBOUND   ' Array of dimensions
End Type

Public Const HEAP_ZERO_MEMORY = &H8&

Declare Function HeapCreate Lib "kernel32" (ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As Long
Declare Function HeapDestroy Lib "kernel32" (ByVal hHeap As Long) As Long
Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Declare Function HeapReAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long, ByVal dwBytes As Long) As Long
Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
Declare Function HeapSize Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long

Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long)

Declare Function VariantCopyIndPtrVar Lib "oleaut32" Alias "VariantCopyInd" (ByVal pvargDest As Long, pvargSrc As Variant) As Long
Declare Function VariantCopyVarPtr Lib "oleaut32" Alias "VariantCopy" (pvargDest As Variant, ByVal pvargSrc As Long) As Long
Declare Function VariantClear Lib "oleaut32" (ByVal pvarg As Long) As Long

Private Declare Function IIDFromString Lib "ole32" (ByVal lpsz As Long, rguid As UUID) As Long
Private Declare Function IsEqualGUID Lib "ole32" (rguid1 As UUID, rguid2 As UUID) As Boolean

Declare Function VarPtrArray Lib "kernel32" Alias "lstrcpyA" (PtrDest() As Any, PtrSrc() As Any) As Long

Const S_FALSE = &H1&
Const E_NOTIMPL = &H80004001
Const E_NOINTERFACE = &H80004002

Private Function AddRef(This As IEnumVARIANT) As Long

 ' Increment the reference count

  This.RefCount = This.RefCount + 1

  ' Return the reference count
  AddRef = This.RefCount

End Function

Private Function AddrOf(ByVal Add As Long) As Long

  AddrOf = Add

End Function

Private Function Clone(This As IEnumVARIANT, NewIEnumVARIANT As IEnumVARIANT) As Long

  Clone = E_NOTIMPL

End Function

Public Function CreateIEnumVARIANT(ByVal hHeap As Long, _
                                   ItemSA As SAFEARRAY_1D) As IUnknown

 Dim vtable(0 To 6) As Long
 Dim IEnm As IEnumVARIANT, lObjPtr As Long

  ' Initialize IIDs
  IIDFromString StrPtr(sIID_IEnumVARIANT), IID_IEnumVARIANT
  IIDFromString StrPtr(sIID_IUnknown), IID_IUnknown

  ' Create the v-table
  vtable(0) = AddrOf(AddressOf QueryInterface) ' IUnknown.QueryInterface
  vtable(1) = AddrOf(AddressOf AddRef)         ' IUnknown.AddRef
  vtable(2) = AddrOf(AddressOf Release)        ' IUnknown.Release
  vtable(3) = AddrOf(AddressOf Next_)          ' IEnumVARIANT.Next
  vtable(4) = AddrOf(AddressOf Skip)           ' IEnumVARIANT.Skip
  vtable(5) = AddrOf(AddressOf Reset)          ' IEnumVARIANT.Reset
  vtable(6) = AddrOf(AddressOf Clone)          ' IEnumVARIANT.Clone

  ' Fill a temporary IEnumVariant struct
  With IEnm

    ' Copy the pointer to
    ' the SAFEARRAY to the array
    MoveMemory ByVal VarPtrArray(.Items, .Items), VarPtr(ItemSA), 4

    .CurrentIndex = 1
    .MaxIdx = ItemSA.Bounds(0).cElements
    .hHeap = hHeap
    .RefCount = 1

    ' Allocate memory for the vtable
    .vtable = HeapAlloc(hHeap, HEAP_ZERO_MEMORY, 28)

    ' Copy the v-table
    MoveMemory ByVal .vtable, vtable(0), 28

  End With

  ' Allocate memory for the object
  lObjPtr = HeapAlloc(hHeap, HEAP_ZERO_MEMORY, LenB(IEnm))

  ' Copy the struct to the allocated memory
  MoveMemory ByVal lObjPtr, IEnm, LenB(IEnm)

  ' Remove the SAFEARRAY struct
  ' from the temporary IEnumVARIANT UDT
  MoveMemory ByVal VarPtrArray(IEnm.Items, IEnm.Items), 0&, 4

  ' Copt the pointer to the return value
  MoveMemory CreateIEnumVARIANT, lObjPtr, 4

End Function

':) Ulli's Code Formatter V2.0 (2001-01-23 10:53:24) 95 + 170 = 265 Lines


Private Function QueryInterface(This As IEnumVARIANT, riid As UUID, lObj As Long) As Long

  If IsEqualGUID(riid, IID_IUnknown) Or _
     IsEqualGUID(riid, IID_IEnumVARIANT) Then

    ' Return a pointer to
    ' this object
    lObj = VarPtr(This)

    ' Increment the reference count
    This.RefCount = This.RefCount + 1

   Else

    ' Set the return value to "Nothing"
    lObj = 0

    ' Return the error
    QueryInterface = E_NOINTERFACE

  End If

End Function


Private Function Release(This As IEnumVARIANT) As Long

 ' Decrement the reference count

  This.RefCount = This.RefCount - 1

  ' Return the reference count
  Release = This.RefCount

  ' Destroy the object if
  ' the reference count is 0
  If This.RefCount = 0 Then

    ' Remove the reference from
    ' the items array
    MoveMemory ByVal VarPtrArray(This.Items, This.Items), 0&, 4

    ' Release the memory
    ' used by the v-table
    HeapFree This.hHeap, 0, This.vtable

    ' Release the object itself
    HeapFree This.hHeap, 0, VarPtr(This)

  End If

End Function

Private Function Reset(This As IEnumVARIANT) As Long

  This.CurrentIndex = 1

End Function

Private Function Skip(This As IEnumVARIANT, ByVal celt As Long) As Long

  This.CurrentIndex = This.CurrentIndex + celt

End Function

Private Function Next_(This As IEnumVARIANT, ByVal celt As Long, rgVar As Variant, ByVal lpCeltFetched As Long) As Long

  With This

    If .CurrentIndex <= .MaxIdx Then

      ' Return a copy of the
      ' stored variant
      VariantCopyVarPtr rgVar, .Items(.CurrentIndex)

      ' Increment the index
      .CurrentIndex = .CurrentIndex + 1

      If lpCeltFetched Then MoveMemory ByVal lpCeltFetched, 1, 4

     Else

      If lpCeltFetched Then MoveMemory ByVal lpCeltFetched, 0, 4

      Next_ = S_FALSE

    End If

  End With

End Function
