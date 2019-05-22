Attribute VB_Name = "VersionModule"
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



' version 0.03
' update: 2019-05-22
' ========================================================================================

' small doc about data migration between objects:
' WeeklyCoverage ->
'   Corail Handler ->
'       Screen Class (HTTP Req Handler GET and POST) ->
'           CorailData ->
'               Parser(DomHandler) ->
'                   SuitableData (ExcelData) ->
'                       next? LayoutClass?

' ========================================================================================

' version 0.02
' update: 2019-05-21
' ========================================================================================

' next step: also a second branch on github 2019-05-21
' first POST HTTP with successful download from corail from screen 2510

' ========================================================================================


' version 0.01
' ========================================================================================

' first prototype of weekly coverage on corail
' oop: WeeklyCoverage as main object

' ========================================================================================



