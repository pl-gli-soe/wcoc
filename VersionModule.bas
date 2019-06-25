Attribute VB_Name = "VersionModule"
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



' ========================================================================================

' ==================================== TO DO =============================================

' ========================================================================================
' TT marking
' supplier name / COFOR to be added
' plt name to be added
' Adding schedules (3560 as per DHRQ) - as per most recent calc. date
' in comment  DHRQ / DHEO limit to date only
' "SLOTs" - TBD
' coverage requirements line to be distinguished (font bold + interior  RGB (242,242,242))
'
' http request on screen 2610 WEEKLY ROLLNG to test values if OK - seperate branch
' - new classes req: CorailDataFrom2610, Corail_2610_Screen, SuitableData2610
' http request on screen 3560 - to verify - seperate branch of code to test also
' - new classes req: CorailDataFrom3560, Corail_3560_Screen, SuitableData3560
' http request to check 3040 - with even hourly rqms - with also long range for Tychy
' put common logic which will usable also for Fire Flake Hourly - put same interface
' new classes req: CorailDataFrom3040, Corail_3040_Screen, SuitableData3040
'
' some estimated calc on TTIME
' based on orders using initial diff in dates between DHEO and DHRQ
'
' to consider SuitableXtraDataXXXX - for metioned screens depending on req from osea team
' ========================================================================================


' version 0.11c Paulina ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' update: 2019-06-24
' ========================================================================================
' Class ; WeeklyLayout Ln 336 - 351 -  graphic adjustment : frame format, requirements distinguished
' - ready for plant / COFOR Ln 380
' "SLOTS" - deactivated
' ribbon Corail image - changed
' ========================================================================================



' version 0.11b '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' update: 2019-06-12
' ========================================================================================
' !!! lets agree that : Public Sub main is untouchable - major logic to run only
' general code like main public subs from class WeeklyCoverage
' Class; WeeklyLayout - ILayout_finalTouchOnRep ; remove Gridlines of output
' reallocate remove gridline to seperate sub finalTouch recognize which sheet is which
' Class; WeeklyCoverage - report sheet - hidden ?? - only prototype sheet
' probably will be deleted in next version
' coverage column B - autofit - to be redefined anyway later, but ok
' simple ribbon created - OK -> design zostawiam Tobie (Ribbon Module OK)
' ========================================================================================




' version 0.11a Paulina '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' update: 2019-06-10 09:00
' ========================================================================================
' Class; WeeklyLayout - ILayout_finalTouchOnRep ; remove Gridlines of output
' Class; WeeklyCoverage - report sheet - hidden
' coverage column B - autofit
' simple ribbon created
' ========================================================================================



' version 0.10
' update: 2019-06-03 15:00
' ========================================================================================
'Request URL: http://ei.control.erp.corail.inetpsa.com/getFbpcForProductSummaryList.do?productCode=9819671280
'request Method: Post
'Status Code: 200
'Remote Address: 10.208.3.205:80
'referrer Policy: no -referrer - when - downgrade
' ========================================================================================


' version 0.09
' update: 2019-06-03 13:45
' ========================================================================================

' final touch on prototype output

' ========================================================================================



' version 0.08
' update: 2019-05-30
' ========================================================================================

' first connection with WeeklyLayout which implements from ILayout - instance theLayout
' final solution on excel rep

' ========================================================================================

' version 0.07
' update: 2019-05-29
' ========================================================================================

' komentarz dla implemnetacji i aranzu danych:
' ST (past & today) + AN?
' w obiekcie weekly coverage - wyrzucamy liste do excela
' zakladam, ze jesli data na DHRQ jest do dzis + ma status ST
' oznacza, ze material zostal przyjety - wiadomka w kolumnie order zostajemy z danymi...
'
' EF <=
' in transit?
' jesli mamy taki milestone (conajmniej), to znaczy, ze lecimy juz z materialem
'
' EO - order ino!

' ========================================================================================



' version 0.06
' update: 2019-05-28
' ========================================================================================

' first prototype for shipping data collected into excel cells
' still missing some orders proper config

' ========================================================================================


' version 0.05
' update: 2019-05-24
' ========================================================================================

' corail data from 2510 need to be partially, so there will be an collection instead simple:
'
' Private rawString As String - will be collection of these
' + Private Sub ICorailData_setString(arg As String) will not assign to primitive
' but adding a new one with simple string and dom object into new element
' of metioned collection

' ========================================================================================


' version 0.04
' update: 2019-05-23
' ========================================================================================

' small doc about data migration between objects:
' SuitableData without Main Interface - passing as Variant, so no list after dot.

' ========================================================================================

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



