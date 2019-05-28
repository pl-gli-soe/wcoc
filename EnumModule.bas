Attribute VB_Name = "EnumModule"
'
' __        __        _    _          ____
' \ \      / /__  ___| | _| |_   _   / ___|_____   _____ _ __ __ _  __ _  ___
'  \ \ /\ / / _ \/ _ \ |/ / | | | | | |   / _ \ \ / / _ \ '__/ _` |/ _` |/ _ \
'   \ V  V /  __/  __/   <| | |_| | | |__| (_) \ V /  __/ | | (_| | (_| |  __/
'    \_/\_/ \___|\___|_|\_\_|\__, |  \____\___/ \_/ \___|_|  \__,_|\__, |\___|
'   ___  _ __    / ___|___  _|___/_ _(_) |                         |___/
'  / _ \| '_ \  | |   / _ \| '__/ _` | | |
' | (_) | | | | | |__| (_) | | | (_| | | |
'  \___/|_| |_|  \____\___/|_|  \__,_|_|_|
'
'
'01010111 01100101 01100101 01101011 01101100 01111001  01000011 01101111 01110110 01100101 01110010 01100001 01100111 01100101
'01101111 01101110  01000011 01101111 01110010 01100001 01101001 01101100
'
'
'
'The MIT License (MIT)
'
'Copyright (c) 2019 FORREST
' Mateusz Milewski mateusz.milewski@mpsa.com aka FORREST
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.



Public Enum E_TYPE_OF_LAYOUT
    E_TYPE_OF_LAYOUT_WEEKLY_DEFAULT
    E_TYPE_OF_LAYOUT_WEEKLY_CUSTOM
    E_TYPE_OF_LAYOUT_WEEKLY_UNDEFINED
End Enum

Public Enum E_TYPE_OF_DATA_FROM_CORAIL
    E_TYPE_OF_DATA_FROM_CORAIL_DEFAULT
    E_TYPE_OF_DATA_FROM_CORAIL_2510_CALC
    E_TYPE_OF_DATA_FROM_CORAIL_2510_UPDATED
End Enum

Public Enum E_2510_ELEMENT
    E_2510_SGR_LINE
    E_2510_PRODUCT
    E_2510_QTY
    E_2510_ORDER_NUMBER
    E_2510_DN
    E_2510_DH
    E_2510_S1
    E_2510_ROUTING
    E_2510_D
    E_2510_SELLER
    E_2510_SHIPPER
    E_2510_S2
    E_2510_UM
End Enum
