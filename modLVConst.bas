Attribute VB_Name = "modLVConst"
Option Explicit

' (Win95)
Public Const WC_LISTVIEWA As String = "SysListView32"
Public Const WC_LISTVIEW  As String = WC_LISTVIEWA

Const LVM_FIRST As Long = &H1000

'ListView Styles(LVS_)
Public Const LVS_ICON As Long = &H0
Public Const LVS_REPORT As Long = &H1
Public Const LVS_SMALLICON As Long = &H2
Public Const LVS_LIST As Long = &H3
Public Const LVS_TYPEMASK As Long = &H3
Public Const LVS_SINGLESEL As Long = &H4
Public Const LVS_SHOWSELALWAYS As Long = &H8
Public Const LVS_SORTASCENDING As Long = &H10
Public Const LVS_SORTDESCENDING As Long = &H20
Public Const LVS_SHAREIMAGELISTS As Long = &H40
Public Const LVS_NOLABELWRAP As Long = &H80
Public Const LVS_AUTOARRANGE As Long = &H100
Public Const LVS_EDITLABELS As Long = &H200
Public Const LVS_OWNERDATA As Long = &H1000        'IE 3+ only
Public Const LVS_NOSCROLL As Long = &H2000

Public Const LVS_TYPESTYLEMASK As Long = &HFC00

Public Const LVS_ALIGNTOP As Long = &H0
Public Const LVS_ALIGNLEFT As Long = &H800
Public Const LVS_ALIGNMASK As Long = &HC00

Public Const LVS_OWNERDRAWFIXED As Long = &H400
Public Const LVS_NOCOLUMNHEADER As Long = &H4000
Public Const LVS_NOSORTHEADER As Long = &H8000

'------------------------------------------------------------------------
'ListView Messages(LVM_)(Generic)

Public Const LVM_GETBKCOLOR As Long = (LVM_FIRST + 0)
Public Const LVM_SETBKCOLOR As Long = (LVM_FIRST + 1)
Public Const LVM_GETIMAGELIST As Long = (LVM_FIRST + 2)
Public Const LVM_SETIMAGELIST As Long = (LVM_FIRST + 3)
Public Const LVM_GETITEMCOUNT As Long = (LVM_FIRST + 4)

Public Const LVM_DELETEITEM As Long = (LVM_FIRST + 8)
Public Const LVM_DELETEALLITEMS As Long = (LVM_FIRST + 9)
Public Const LVM_GETCALLBACKMASK As Long = (LVM_FIRST + 10)
Public Const LVM_SETCALLBACKMASK As Long = (LVM_FIRST + 11)
Public Const LVM_GETNEXTITEM As Long = (LVM_FIRST + 12)

Public Const LVM_SETITEMPOSITION As Long = (LVM_FIRST + 15)
Public Const LVM_GETITEMPOSITION As Long = (LVM_FIRST + 16)

Public Const LVM_HITTEST As Long = (LVM_FIRST + 18)
Public Const LVM_ENSUREVISIBLE As Long = (LVM_FIRST + 19)
Public Const LVM_SCROLL As Long = (LVM_FIRST + 20)
Public Const LVM_REDRAWITEMS As Long = (LVM_FIRST + 21)
Public Const LVM_ARRANGE As Long = (LVM_FIRST + 22)

Public Const LVM_GETEDITCONTROL As Long = (LVM_FIRST + 24)

Public Const LVM_DELETECOLUMN As Long = (LVM_FIRST + 28)
Public Const LVM_GETCOLUMNWIDTH As Long = (LVM_FIRST + 29)
Public Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)

Public Const LVM_GETHEADER As Long = (LVM_FIRST + 31)  'IE 3+ only

Public Const LVM_CREATEDRAGIMAGE As Long = (LVM_FIRST + 33)
Public Const LVM_GETVIEWRECT As Long = (LVM_FIRST + 34)
Public Const LVM_GETTEXTCOLOR As Long = (LVM_FIRST + 35)
Public Const LVM_SETTEXTCOLOR As Long = (LVM_FIRST + 36)
Public Const LVM_GETTEXTBKCOLOR As Long = (LVM_FIRST + 37)
Public Const LVM_SETTEXTBKCOLOR As Long = (LVM_FIRST + 38)
Public Const LVM_GETTOPINDEX As Long = (LVM_FIRST + 39)
Public Const LVM_GETCOUNTPERPAGE As Long = (LVM_FIRST + 40)
Public Const LVM_GETORIGIN As Long = (LVM_FIRST + 41)
Public Const LVM_UPDATE As Long = (LVM_FIRST + 42)
Public Const LVM_SETITEMSTATE As Long = (LVM_FIRST + 43)
Public Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)
Public Const LVM_SETITEMCOUNT As Long = (LVM_FIRST + 47)
Public Const LVM_SORTITEMS As Long = (LVM_FIRST + 48)
Public Const LVM_SETITEMPOSITION32 As Long = (LVM_FIRST + 49)
Public Const LVM_GETSELECTEDCOUNT As Long = (LVM_FIRST + 50)
Public Const LVM_GETITEMSPACING As Long = (LVM_FIRST + 51)

Public Const LVM_SETICONSPACING As Long = (LVM_FIRST + 53) 'IE 3+ only

Public Const LVM_GETSUBITEMRECT As Long = (LVM_FIRST + 56)
Public Const LVM_SUBITEMHITTEST As Long = (LVM_FIRST + 57)
Public Const LVM_SETCOLUMNORDERARRAY As Long = (LVM_FIRST + 58)
Public Const LVM_GETCOLUMNORDERARRAY As Long = (LVM_FIRST + 59)
Public Const LVM_SETHOTITEM As Long = (LVM_FIRST + 60)
Public Const LVM_GETHOTITEM As Long = (LVM_FIRST + 61)
Public Const LVM_SETHOTCURSOR As Long = (LVM_FIRST + 62)
Public Const LVM_GETHOTCURSOR As Long = (LVM_FIRST + 63)
Public Const LVM_APPROXIMATEVIEWRECT As Long = (LVM_FIRST + 64)
Public Const LVM_SETWORKAREA As Long = (LVM_FIRST + 65)

Public Const LVM_GETSELECTIONMARK As Long = (LVM_FIRST + 66) 'Win32 and IE 4 only
Public Const LVM_SETSELECTIONMARK As Long = (LVM_FIRST + 67) 'Win32 and IE 4 only
Public Const LVM_GETWORKAREA As Long = (LVM_FIRST + 70)      'Win32 and IE 4 only
Public Const LVM_SETHOVERTIME As Long = (LVM_FIRST + 71)     'Win32 and IE 4 only
Public Const LVM_GETHOVERTIME As Long = (LVM_FIRST + 72)     'Win32 and IE 4 only 

'------------------------------------------------------------------------
'ListView Messages(LVM_)(Win95 - specific)

Public Const LVM_GETITEM As Long = (LVM_FIRST + 5)
Public Const LVM_SETITEM As Long = (LVM_FIRST + 6)

Public Const LVM_INSERTITEMA As Long = (LVM_FIRST + 7)
Public Const LVM_INSERTITEM As Long = LVM_INSERTITEMA

Public Const LVM_FINDITEMA As Long = (LVM_FIRST + 13)
Public Const LVM_FINDITEM As Long = LVM_FINDITEMA

Public Const LVM_GETSTRINGWIDTHA As Long = (LVM_FIRST + 17)
Public Const LVM_GETSTRINGWIDTH As Long = LVM_GETSTRINGWIDTHA

Public Const LVM_EDITLABELA As Long = (LVM_FIRST + 23)
Public Const LVM_EDITLABEL As Long = LVM_EDITLABELA

Public Const LVM_GETCOLUMNA As Long = (LVM_FIRST + 25)
Public Const LVM_GETCOLUMN As Long = LVM_GETCOLUMNA

Public Const LVM_SETCOLUMNA As Long = (LVM_FIRST + 26)
Public Const LVM_SETCOLUMN As Long = LVM_SETCOLUMNA

Public Const LVM_INSERTCOLUMNA As Long = (LVM_FIRST + 27)
Public Const LVM_INSERTCOLUMN As Long = LVM_INSERTCOLUMNA

Public Const LVM_GETITEMTEXTA As Long = (LVM_FIRST + 45)
Public Const LVM_GETITEMTEXT As Long = LVM_GETITEMTEXTA

Public Const LVM_SETITEMTEXTA As Long = (LVM_FIRST + 46)
Public Const LVM_SETITEMTEXT As Long = LVM_SETITEMTEXTA

Public Const LVM_GETISEARCHSTRINGA As Long = (LVM_FIRST + 52)
Public Const LVM_GETISEARCHSTRING As Long = LVM_GETISEARCHSTRINGA

Public Const LVM_SETBKIMAGEA As Long = (LVM_FIRST + 68)        'Win32 and IE 4 only
Public Const LVM_GETBKIMAGEA As Long = (LVM_FIRST + 69)        'Win32 and IE 4 only
'Public Const LVBKIMAGE As Long =LVBKIMAGEA                     'Win32 and IE 4 only
'Public Const LPLVBKIMAGE As Long =LPLVBKIMAGEA                 'Win32 and IE 4 only
Public Const LVM_SETBKIMAGE As Long = LVM_SETBKIMAGEA          'Win32 and IE 4 only
Public Const LVM_GETBKIMAGE As Long = LVM_GETBKIMAGEA          'Win32 and IE 4 only

'------------------------------------------------------------------------
'ListView Messages(LVM_)(Unicode - specific)

'Public Const LVM_GETITEM As Long =(LVM_FIRST + 75)
'Public Const LVM_SETITEM As Long =(LVM_FIRST + 76)
'
'Public Const LVM_INSERTITEMW As Long =(LVM_FIRST + 77)
'Public Const LVM_INSERTITEM As Long =LVM_INSERTITEMW
'
'Public Const LVM_FINDITEMW As Long =(LVM_FIRST + 83)
'Public Const LVM_FINDITEM As Long =LVM_FINDITEMW
'
'Public Const LVM_GETSTRINGWIDTHW As Long =(LVM_FIRST + 87)
'Public Const LVM_GETSTRINGWIDTH As Long =LVM_GETSTRINGWIDTHW
'
'Public Const LVM_EDITLABELW As Long =(LVM_FIRST + 118)
'Public Const LVM_EDITLABEL As Long =LVM_EDITLABELW
'
'Public Const LVM_GETCOLUMNW As Long =(LVM_FIRST + 95)
'Public Const LVM_GETCOLUMN As Long =LVM_GETCOLUMNW
'
'Public Const LVM_SETCOLUMNW As Long =(LVM_FIRST + 96)
'Public Const LVM_SETCOLUMN As Long =LVM_SETCOLUMNW
'
'Public Const LVM_INSERTCOLUMNW As Long =(LVM_FIRST + 97)
'Public Const LVM_INSERTCOLUMN As Long =LVM_INSERTCOLUMNW
'
'Public Const LVM_GETITEMTEXTW As Long =(LVM_FIRST + 115)
'Public Const LVM_GETITEMTEXT As Long =LVM_GETITEMTEXTW
'
'Public Const LVM_SETITEMTEXTW As Long =(LVM_FIRST + 116)
'Public Const LVM_SETITEMTEXT As Long =LVM_SETITEMTEXTW
'
'Public Const LVM_GETISEARCHSTRINGW As Long =(LVM_FIRST + 117)
'Public Const LVM_GETISEARCHSTRING As Long =LVM_GETISEARCHSTRINGW
'
'Public Const LVM_GETBKIMAGEW As Long =(LVM_FIRST + 139)        'Win32 and IE 4 only
'Public Const LVM_SETBKIMAGEW As Long =(LVM_FIRST + 138)        'Win32 and IE 4 only
'Public Const LVBKIMAGE As Long =LVBKIMAGEW                     'Win32 and IE 4 only
'Public Const LPLVBKIMAGE As Long =LPLVBKIMAGEW                 'Win32 and IE 4 only
'Public Const LVM_SETBKIMAGE As Long =LVM_SETBKIMAGEW           'Win32 and IE 4 only
'Public Const LVM_GETBKIMAGE As Long =LVM_GETBKIMAGEW           'Win32 and IE 4 only

'------------------------------------------------------------------------
'ListView Extended Style Messages (LVS_EX_) (Win95-specific)

Public Const LVS_EX_GRIDLINES As Long = &H1
Public Const LVS_EX_SUBITEMIMAGES As Long = &H2
Public Const LVS_EX_CHECKBOXES As Long = &H4
Public Const LVS_EX_TRACKSELECT As Long = &H8
Public Const LVS_EX_HEADERDRAGDROP As Long = &H10
Public Const LVS_EX_FULLROWSELECT As Long = &H20      'applies to report mode only
Public Const LVS_EX_ONECLICKACTIVATE As Long = &H40
Public Const LVS_EX_TWOCLICKACTIVATE As Long = &H80
Public Const LVS_EX_FLATSB As Long = &H100            'cannot be cleared - Win32 & IE4 only
Public Const LVS_EX_REGIONAL As Long = &H200          'Win32 & IE4 only
Public Const LVS_EX_INFOTIP As Long = &H400           'listview does InfoTips for you - Win32 & IE4 only

'------------------------------------------------------------------------
'ListView Set Image List Messages (LVSIL_)

Public Const LVSIL_NORMAL As Long = 0
Public Const LVSIL_SMALL As Long = 1
Public Const LVSIL_STATE As Long = 2

'------------------------------------------------------------------------
'ListView Item Format Messages (LVIF_)

Public Const LVIF_TEXT As Long = &H1
Public Const LVIF_IMAGE As Long = &H2
Public Const LVIF_PARAM As Long = &H4
Public Const LVIF_STATE As Long = &H8
Public Const LVIF_INDENT As Long = &H10          'IE 3+ only
Public Const LVIF_NORECOMPUTE As Long = &H800    'IE 3+ only
Public Const LVIF_DI_SETITEM As Long = &H1000

'------------------------------------------------------------------------
'ListView Item State Messages (LVIS_)

Public Const LVIS_FOCUSED As Long = &H1
Public Const LVIS_SELECTED As Long = &H2
Public Const LVIS_CUT As Long = &H4
Public Const LVIS_DROPHILITED As Long = &H8

Public Const LVIS_OVERLAYMASK As Long = &HF00
Public Const LVIS_STATEIMAGEMASK As Long = &HF000

'------------------------------------------------------------------------
'ListView Item Definitions (LVITEM) (Win95)

'Public Const LVITEM As Long =LVITEMA
'Public Const LPLVITEM As Long =LPLVITEMA
'Public Const LV_ITEMA As Long =LVITEMA       'IE 3+ only
'Public Const tagLVITEMA As Long =LV_ITEMA

'ListView Item Definitions (LVITEM) (Unicode)

'Public Const LVITEM As Long =LVITEMW
'Public Const LPLVITEM As Long =LPLVITEMW  'Unicode (NT)
'Public Const LV_ITEM As Long =LVITEM      'IE 3+ only
'Public Const tagLVITEMW As Long =LV_ITEMW

'------------------------------------------------------------------------
'ListView -Misc.Messages

'Public Const INDEXTOSTATEIMAGEMASK(i) ((i) << 12)
Public Const I_INDENTCALLBACK As Long = (-1)              'IE 3+ only
'Public Const LPSTR_TEXTCALLBACKW As Long =((LPWSTR) - 1&) 'Unicode (NT)
'Public Const LPSTR_TEXTCALLBACKA As Long =((LPSTR) - 1&)  'win95

'Public Const LPSTR_TEXTCALLBACK As Long =LPSTR_TEXTCALLBACKW 'Unicode (NT)
'Public Const LPSTR_TEXTCALLBACK As Long =LPSTR_TEXTCALLBACKA 'win95

'------------------------------------------------------------------------
'ListView Notification Item Messages (LVNI_)

Public Const LVNI_ALL As Long = &H0
Public Const LVNI_FOCUSED As Long = &H1
Public Const LVNI_SELECTED As Long = &H2
Public Const LVNI_CUT As Long = &H4
Public Const LVNI_DROPHILITED As Long = &H8

Public Const LVNI_ABOVE As Long = &H100
Public Const LVNI_BELOW As Long = &H200
Public Const LVNI_TOLEFT As Long = &H400
Public Const LVNI_TORIGHT As Long = &H800

'------------------------------------------------------------------------
'ListView Find Item Messages (LVFI_) (Generic)

Public Const LVFI_PARAM As Long = &H1
Public Const LVFI_STRING As Long = &H2
Public Const LVFI_PARTIAL As Long = &H8
Public Const LVFI_WRAP As Long = &H20
Public Const LVFI_NEARESTXY As Long = &H40

'Public Const LV_FINDINFO As Long =LVFINDINFO

'------------------------------------------------------------------------
'ListView Find Item Messages (LVFI_) (Win95)

'Public Const LV_FINDINFOA As Long =LVFINDINFOA
'Public Const LV_FINDINFOA As Long =LVFINDINFOA     'IE 3+ only
'Public Const tagLVFINDINFOA As Long =LV_FINDINFOA
'Public Const LVFINDINFOA As Long =LV_FINDINFOA
'Public Const LVFINDINFO As Long =LVFINDINFOA

'------------------------------------------------------------------------
'ListView Find Item Messages (LVFI_) (Unicode)

'Public Const LV_FINDINFOW As Long =LVFINDINFOW
'Public Const LV_FINDINFOW As Long =LVFINDINFOW     'IE 3+ only
'Public Const tagLVFINDINFOW As Long =LV_FINDINFOW
'Public Const LVFINDINFOW As Long =LV_FINDINFOW
'Public Const LVFINDINFO As Long =LVFINDINFOW

'------------------------------------------------------------------------
'ListView Find ItemRect Messages (LVIR_)

Public Const LVIR_BOUNDS As Long = 0
Public Const LVIR_ICON As Long = 1
Public Const LVIR_LABEL As Long = 2
Public Const LVIR_SELECTBOUNDS As Long = 3
'ListView Hit Test Messages (LVHT_)
Public Const LVHT_NOWHERE As Long = &H1
Public Const LVHT_ONITEMICON As Long = &H2
Public Const LVHT_ONITEMLABEL As Long = &H4
Public Const LVHT_ONITEMSTATEICON As Long = &H8
Public Const LVHT_ONITEM As Long = (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)

Public Const LVHT_ABOVE As Long = &H8
Public Const LVHT_BELOW As Long = &H10
Public Const LVHT_TORIGHT As Long = &H20
Public Const LVHT_TOLEFT As Long = &H40

'Public Const LV_HITTESTINFO As Long =LVHITTESTINFO       'IE 3+ only
'Public Const tagLVHITTESTINFO As Long =LV_HITTESTINFO
'Public Const LVHITTESTINFO As Long =LV_HITTESTINFO

'------------------------------------------------------------------------
'ListView Arrange Messages (LVA_)

Public Const LVA_DEFAULT As Long = &H0
Public Const LVA_ALIGNLEFT As Long = &H1
Public Const LVA_ALIGNTOP As Long = &H2
Public Const LVA_SNAPTOGRID As Long = &H5

'------------------------------------------------------------------------
'ListView Column Messages (LVC_) (Generic)

'Public Const LV_COLUMN As Long =LVCOLUMN       'IE 3+ only

'------------------------------------------------------------------------
'ListView Column Messages (LVC_) (Win95)

'Public Const LV_COLUMNA As Long =LVCOLUMNA               'IE 3+ only
'Public Const tagLVCOLUMNA As Long =LV_COLUMNA
'Public Const LVCOLUMNA As Long =LV_COLUMNA
'Public Const LVCOLUMN As Long =LVCOLUMNA
'Public Const LPLVCOLUMN As Long =LPLVCOLUMNA

'------------------------------------------------------------------------
'ListView Column Messages (LVC_) (Unicode)

'Public Const LV_COLUMNW As Long =LVCOLUMNW       'IE 3+ only
'Public Const tagLVCOLUMNW As Long =LV_COLUMNW
'Public Const LVCOLUMNW As Long =LV_COLUMNW
'Public Const LVCOLUMN As Long =LVCOLUMNW
'Public Const LPLVCOLUMN As Long =LPLVCOLUMNW

'------------------------------------------------------------------------
'ListView Column Flag Messages (LVCF_) (LVC.mask)

Public Const LVCF_FMT As Long = &H1
Public Const LVCF_WIDTH As Long = &H2
Public Const LVCF_TEXT As Long = &H4
Public Const LVCF_SUBITEM As Long = &H8
Public Const LVCF_IMAGE As Long = &H10     'IE 3+ only
Public Const LVCF_ORDER As Long = &H20     'IE 3+ only

'------------------------------------------------------------------------
'ListView Column Format Messages (LVCFMT_) (LVC.fmt)

Public Const LVCFMT_LEFT As Long = &H0
Public Const LVCFMT_RIGHT As Long = &H1
Public Const LVCFMT_CENTER As Long = &H2
Public Const LVCFMT_JUSTIFYMASK As Long = &H3
Public Const LVCFMT_IMAGE As Long = &H800              'IE 3+ only
Public Const LVCFMT_BITMAP_ON_RIGHT As Long = &H1000   'IE 3+ only
Public Const LVCFMT_COL_HAS_IMAGES As Long = &H8000    'IE 4 only

'------------------------------------------------------------------------
'ListView Set Column Width Messages (LVSCW_)

Public Const LVSCW_AUTOSIZE As Long = -1
Public Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

'------------------------------------------------------------------------
'ListView Background Image Flags (LVBKIF_)

Public Const LVBKIF_SOURCE_NONE As Long = &H0      'Win32 and IE 4 only
Public Const LVBKIF_SOURCE_HBITMAP As Long = &H1   'Win32 and IE 4 only
Public Const LVBKIF_SOURCE_URL As Long = &H2       'Win32 and IE 4 only
Public Const LVBKIF_SOURCE_MASK As Long = &H3      'Win32 and IE 4 only
Public Const LVBKIF_STYLE_NORMAL As Long = &H0     'Win32 and IE 4 only
Public Const LVBKIF_STYLE_TILE As Long = &H10      'Win32 and IE 4 only
Public Const LVBKIF_STYLE_MASK As Long = &H10      'Win32 and IE 4 only

'------------------------------------------------------------------------
'ListView Notification Messages (LVN_) (Generic)

'Public Const LVN_ITEMCHANGING As Long =(LVN_FIRST - 0)
'Public Const LVN_ITEMCHANGED As Long =(LVN_FIRST - 1)
'Public Const LVN_INSERTITEM As Long =(LVN_FIRST - 2)
'Public Const LVN_DELETEITEM As Long =(LVN_FIRST - 3)
'Public Const LVN_DELETEALLITEMS As Long =(LVN_FIRST - 4)
'
'Public Const LVN_COLUMNCLICK As Long =(LVN_FIRST - 8)
'Public Const LVN_BEGINDRAG As Long =(LVN_FIRST - 9)
'Public Const LVN_BEGINRDRAG As Long =(LVN_FIRST - 11)
'
'Public Const LVN_ODCACHEHINT As Long =(LVN_FIRST - 13)         'IE 3+ only
'Public Const LVN_ITEMACTIVATE As Long =(LVN_FIRST - 14)
'Public Const LVN_ODSTATECHANGED As Long =(LVN_FIRST - 15)
'
'Public Const LVN_HOTTRACK As Long =(LVN_FIRST - 21)
'
'Public Const LVN_KEYDOWN As Long =(LVN_FIRST - 55)
'Public Const LVN_MARQUEEBEGIN As Long =(LVN_FIRST - 56)        'IE 3+ only
'
''------------------------------------------------------------------------
''ListView Notification Messages (LVN_) (Win95)
'
'Public Const LVN_BEGINLABELEDITA As Long =(LVN_FIRST - 5)
'Public Const LVN_ENDLABELEDITA As Long =(LVN_FIRST - 6)
'
'Public Const LVN_GETDISPINFOA As Long =(LVN_FIRST - 50)
'Public Const LVN_SETDISPINFOA As Long =(LVN_FIRST - 51)
'Public Const LVN_ODFINDITEMA As Long =(LVN_FIRST - 52)       'IE 3+ only
'Public Const LVN_ODFINDITEM As Long =LVN_ODFINDITEMA
'
'Public Const LVN_BEGINLABELEDIT As Long =LVN_BEGINLABELEDITA
'Public Const LVN_ENDLABELEDIT As Long =LVN_ENDLABELEDITA
'Public Const LVN_GETDISPINFO As Long =LVN_GETDISPINFOA
'Public Const LVN_SETDISPINFO As Long =LVN_SETDISPINFOA
'
'Public Const LV_DISPINFOA As Long =NMLVDISPINFOA             'IE 3+ only
'Public Const tagLVDISPINFO As Long =LV_DISPINFO
'Public Const NMLVDISPINFOA As Long =LV_DISPINFOA
'Public Const NMLVDISPINFO As Long =NMLVDISPINFOA

'------------------------------------------------------------------------
'ListView Notification Messages (LVN_) (Unicode)

'Public Const LVN_BEGINLABELEDITW As Long =(LVN_FIRST - 75)
'Public Const LVN_ENDLABELEDITW As Long =(LVN_FIRST - 76)
'
'Public Const LVN_GETDISPINFOW As Long =(LVN_FIRST - 77)
'Public Const LVN_SETDISPINFOW As Long =(LVN_FIRST - 78)
'Public Const LVN_ODFINDITEMW As Long =(LVN_FIRST - 79)         'IE 3+ only
'Public Const LVN_ODFINDITEM As Long =LVN_ODFINDITEMW
'
'Public Const LVN_BEGINLABELEDIT As Long =LVN_BEGINLABELEDITW
'Public Const LVN_ENDLABELEDIT As Long =LVN_ENDLABELEDITW
'Public Const LVN_GETDISPINFO As Long =LVN_GETDISPINFOW
'Public Const LVN_SETDISPINFO As Long =LVN_SETDISPINFOW
'
'Public Const LV_DISPINFOW As Long =NMLVDISPINFOW               'IE 3+ only
'Public Const tagLVDISPINFOW As Long =LV_DISPINFOW
'Public Const NMLVDISPINFOW As Long =LV_DISPINFOW
'Public Const NMLVDISPINFO As Long =NMLVDISPINFOW

':) Ulli's Code Formatter V2.0 (2001-01-23 10:53:22) 458 + 2 As Long =460 Lines
