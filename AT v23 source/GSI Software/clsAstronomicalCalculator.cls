VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAstronomicalCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'   Modified and Converted to VB6 OOPs model by Jose J. Reyes-Tomassini
'   Original VBA functions taken from:  http://www.bodmas.org/kepler/sunrise_vba.txt
'   I pressume that the author of the original program is Keith Burnett, a British physicist

'   Event returns the time of day of sunrise, sunset, or the beginning and end one of
'   three kinds of twilight
'
'   The method used here is adapted from Montenbruck and Pfleger's
'   Astronomy on the Personal Computer, 3rd Ed, section 3.8
'
'   The arguments for the function are as follows...
'
'   Day of year using coded decimal (#day in year + decimal that indicates year) - your date in zone time
'   longitude in degrees, west negative
'   latitude in degrees, north positive
'   time zone in decimal hours, west or 'behind' Greenwich negative
'   event - a code integer representing the event you want as follows
'       1 = sunrise                 -1  = sunset
'       2 = begin civil twilight    -2  = end civil twilight
'       3 = begin nautical twi      -3  = end nautical twi
'       4 = begin astro twi         -4  = end astro twi

'   Events are defined in the enumarated type ASTRALFUNCTION

'   the results are returned as a variant with either a time of day
'   in zone time or a string reporting an 'event not occuring' condition
'   event not occuring can be one of the following
'       ALL_NIGHT    always below horizon, so no rise or set
'       ALL_DAY    always above horizon, so no rise or set
'       N/A    the particular rise or set event does not occur on that day
'
'   The function will produce meaningful results at all latitudes
'   but there will be a small range of latitudes around 67.43 degrees North or South
'   when the function might indicate a sunrise very close to noon (or a sunset
'   very soon after noon) where in fact the Sun is below the horizon all day
'   this behaviour relates to the approximate Sun position formulas in use
'
'   As always, the sunrise / set times relate to an earth which is smooth
'   and has no large obstructions on the horizon - you might get a close
'   approximation to this at sea but rarely on land. Accuracy more than 1 min
'   of time is not worth striving for - atmospheric refraction alone can
'   alter observed rise times by minutes
'

Dim JulianDateNumber As Double
Dim Yeardate As Long
Dim TimeZone As Long
Dim GLong As Double
Dim GLat As Double
Dim SGlat As Double
Dim CGlat As Double

'variables passed between functions
Dim NZ As Integer
Dim Z1 As Double
Dim Z2 As Double
Dim Xe As Double
Dim Ye As Double

Dim SunRightAscension As Double
Dim SunDeclination As Double

Private Type Parabola
    Xe As Double
    Ye As Double
    Z1 As Double
    Z2 As Double
    NZ As Integer
End Type

Dim Sun As Parabola

Dim SinHO(5) As Double

'constants
Const RADS = 0.0174532925
Const HoursPerDay = 24#
Const DaysPerCentury = 36525#

Const p2 = 6.283185307
Const COSEPS = 0.91748
Const SINEPS = 0.39778


Private Function MJD(Year As Integer, month As Integer, Day As Integer) As Double
'
'   takes the year, month and day as a Gregorian calendar date
'   and returns the modified julian day number
'
    Dim DaysSinceStartOfMJD As Double
    Dim NumberOfLeapYears As Double
    
    Const AverageDaysInMonth = 30.6001!
    
    'Only count the year for leapyear calculation if february has passed... If not, adding the day to the MJD will take care of that!
    If (month <= 2) Then
        month = month + 12
        Year = Year - 1
    End If
    'calculate number of leapyears since 0 year and add to total (#leap years = number of days that need to be added for correction)
    NumberOfLeapYears = Fix(Year / 400#) - Fix(Year / 100#) + Fix(Year / 4#)
    DaysSinceStartOfMJD = 365# * Year - 679004
    MJD = DaysSinceStartOfMJD + NumberOfLeapYears + Fix(AverageDaysInMonth * (month + 1)) + Day
End Function


Private Function frac(x As Double) As Double
'
'  returns the fractional part of x as used in minimoon and minisun
'
    Dim a As Double
    a = x - Fix(x)
    frac = a
End Function

Private Function range(Angle As Double) As Double
'
'   returns an angle in degrees in the range 0 to 360
'   used to condition the arguments for the Sun's orbit
'   in function minisun below
'
    Dim modulus As Double
    modulus = Angle Mod 360
    If Sgn(modulus) = -1 Then
        modulus = modulus + 360
    End If
    range = modulus
End Function

Public Function TimeConvert(T As Double) As String
'
'   takes a time as a decimal number of hours between 0 and 23.9999...
'   and returns a string with the time in hhmm format
'
    Dim H As Double
    Dim Min As Double
    H = Fix(T)
    Min = Fix((T - H) * 60# + 0.5)
    TimeConvert = Format(H, "00") + Format(Min, "00")
    
End Function
Private Function LocalMeanSideralTime(MJD As Double, GLong As Double) As Double
'
'  Takes the mjd and the longitude (west negative) and then returns
'  the local sidereal time in hours. Uses Meeus formula 11.4
    
    Dim Local_Sidereal_Time As Double
    Dim Centuries As Double
    Dim Jday As Double
    'get jd
    Jday = MJD - 51544.5
    'get number of centuries
    Centuries = Jday / 36525#
    
    Local_Sidereal_Time = range(280.46061837 + 360.98564736629 * Jday + 0.000387933 * Centuries * Centuries - Centuries * Centuries * Centuries / 38710000#)
    'mean time
    LocalMeanSideralTime = Local_Sidereal_Time / 15# + GLong / 15#
End Function
    
Private Sub MiniSun(JulianCenturies As Double)
'
'   takes julian centuries since J2000.0
'   sets SunRightAscension and SunDeclination to the value of the Sun coordinates at time JulianCenturies
'
'   positions claimed to be within 1 arc min by Montenbruck and Pfleger
'
    
    Dim L As Double
    Dim m As Double
    Dim DL As Double
    Dim SL As Double
    Dim x As Double
    Dim y As Double
    Dim Z As Double
    Dim Rho As Double
    
    m = p2 * frac(0.993133 + 99.997361 * JulianCenturies)
    DL = 6893# * Sin(m) + 72# * Sin(2 * m)
    L = p2 * frac(0.7859453 + m / p2 + (6191.2 * JulianCenturies + DL) / 1296000#)
    SL = Sin(L)
    x = Cos(L)
    y = COSEPS * SL
    Z = SINEPS * SL
    Rho = Sqr(1 - Z * Z)
    SunDeclination = (360# / p2) * Atn(Z / Rho)
    SunRightAscension = (48# / p2) * Atn(y / (x + Rho))
    If SunRightAscension < 0 Then SunRightAscension = SunRightAscension + HoursPerDay
End Sub


Private Sub Quad(Ym As Double, Yz As Double, Yp As Double)
'
'  finds the parabola throuh the three points (-1,ym), (0,yz), (1, yp)
'  and sets the coordinates of the max/min (if any) xe, ye
'  the values of x where the parabola crosses zero (z1, z2)
'  and the nz number of roots (0, 1 or 2) within the interval [-1, 1]
'
    Dim a As Double
    Dim B As Double
    Dim C As Double
    Dim dis As Double
    Dim dX As Double

With Sun
    .NZ = 0
    a = 0.5 * (Ym + Yp) - Yz
    B = 0.5 * (Yp - Ym)
    C = Yz
    If B = 0 Then .Xe = 0 Else .Xe = -B / (2# * a)
    .Ye = (a * .Xe + B) * .Xe + C
    dis = B * B - 4# * a * C
    If (dis > 0) Then
        dX = 0.5 * Sqr(dis) / Abs(a)
        .Z1 = .Xe - dX
        .Z2 = .Xe + dX
        If (Abs(.Z1) <= 1#) Then .NZ = .NZ + 1
        If (Abs(.Z2) <= 1#) Then .NZ = .NZ + 1
        If (.Z1 < -1#) Then .Z1 = .Z2
    End If
End With
End Sub

Private Function SinAltSun(H As Double) As Double
'returns the sine of the altitude of sun at time H
    Dim JD As Double
    Dim Centuries As Double
    
    Dim tau As Double
    Dim SineOfAltitude As Double
    
    JD = JulianDateNumber + (H / HoursPerDay)
    Centuries = (JD - 51544.5) / DaysPerCentury
    MiniSun Centuries
    ' hour angle of object
    tau = 15 * (LocalMeanSideralTime(JD, GLong) - SunRightAscension)
    ' sin(alt) of object using the conversion formulas
    SineOfAltitude = SGlat * Sin(RADS * SunDeclination) + CGlat * Cos(RADS * SunDeclination) * Cos(RADS * tau)
    SinAltSun = SineOfAltitude
End Function

    
Private Function SunEvent(EventType As Integer) As String
'
'   This is the function that does most of the work
'
Dim sglong As Double
Dim Ym As Double
Dim Yz As Double
Dim Yp As Double
Dim Above As Boolean
Dim RiseTime As Double
Dim SetTime As Double
Dim RiseEventOccurred As Boolean
Dim SetEventOccurred As Boolean
Dim EventClass As Integer
Dim H As Double
Dim outstring As String

Dim NoEventOccurred As Boolean
Dim QueryForRiseEvent As Boolean
Dim QueryForSetEvent As Boolean


'Is this a query for a rise or set?
If Sgn(EventType) = -1 Then
    QueryForSetEvent = True
Else
    QueryForRiseEvent = True
End If


'   main loop takes each value of sinho in turn and finds the rise/set
'   events associated with that altitude of the Sun
'
    EventClass = Abs(EventType)
        H = 1
        Ym = SinAltSun(H - 1) - SinHO(EventClass)
        If (Ym > 0) Then Above = True
        '
        '  the while loop finds the sin(alt) for sets of three consecutive
        '  hours, and then tests for a single zero crossing in the interval
        '  or for two zero crossings in an interval or for a grazing event
        '  The flags rise and sett are set accordingly
        '
        Do While (H <= HoursPerDay) And (Not SetEventOccurred Or Not RiseEventOccurred)
            Yz = SinAltSun(H) - SinHO(EventClass)
            Yp = SinAltSun(H + 1) - SinHO(EventClass)
            Quad Ym, Yz, Yp
            
            ' case when one event is found in the interval
            If (Sun.NZ = 1) Then
                If (Ym < 0) Then
                    RiseTime = H + Sun.Z1
                    RiseEventOccurred = True
                Else
                    SetTime = H + Sun.Z1
                    SetEventOccurred = True
                End If
            End If ' end of nz = 1 case
            '
            '   case where two events are found in this interval
            '   (rare but whole reason we are not using simple iteration)
            '
            If (Sun.NZ = 2) Then
                If (Ye < 0) Then
                    RiseTime = H + Sun.Z2
                    SetTime = H + Sun.Z1
                Else
                    RiseTime = H + Sun.Z1
                    SetTime = H + Sun.Z2
                End If
                RiseEventOccurred = True
                SetEventOccurred = True
            End If
            '
            '   set up the next search interval
            '
            
            Ym = Yp
            H = H + 2

        Loop ' end of while loop
            '
            ' now search has completed, we compile the string to pass back
            ' to the user. The string depends on several combinations
            ' of the above flag (always above or always below) and the rise
            ' and sett flags
            '
            
            'default if something is detected but not queried is a NO EVENT
            outstring = NoEVENT
            
            If QueryForRiseEvent And RiseEventOccurred Then
                outstring = TimeConvert(RiseTime)
            End If
            
            If QueryForSetEvent And SetEventOccurred Then
                outstring = TimeConvert(SetTime)
            End If
            
            'when no event occured (neither a set nor a raise), then either sun is above or below the horizon for all day
            'in this case, check the above flag and return event accordingly
            NoEventOccurred = Not (RiseEventOccurred Or SetEventOccurred)
            
            If NoEventOccurred Then
                If Above Then
                    outstring = ALWAYSup
                Else
                    outstring = ALWAYSdown
                End If
            End If
            
        'return to caller
        SunEvent = outstring
End Function
Public Property Let Longitude(L As Double)
GLong = L * -1#

End Property
Public Property Let Latitude(L As Double)
GLat = L
'components of lat
SGlat = Sin(RADS * GLat)
CGlat = Cos(RADS * GLat)
End Property
Public Property Let LocalTimeZone(TZ As Long)
TimeZone = TZ
End Property
Public Property Let SetYear(y As Long)
Yeardate = y

End Property
Public Property Let SetDay(D As Long)
'Sets the date for all other calculations
'uses date format defined by stamp
'and then converts that to MJD
'day format is : DAYOFYEAR.YEAR

Dim DayYear As Single

Dim date_string As String

'validate first
If D > 0 Then
    date_string = Convert_DayNumber(D)
    JulianDateNumber = MJD(Year(date_string), month(date_string), Day(date_string)) - TimeZone / HoursPerDay
End If
End Property

Public Function EventTime(EventType As Integer) As Integer
'Returns time of event as defined by EVENTTYPE
'EVENTTYPES are enumerated by the ASTRALFUNCTION enumerated type.

Dim H As Double
Dim m As Double

Dim CodedString As String

CodedString = SunEvent(EventType)


If IsNumeric(CodedString) Then
    EventTime = ConvertToMTAPTime(CodedString)
Else
    EventTime = 0
End If

End Function
Private Function ConvertToMTAPTime(S As String) As Integer
'converts the output of this PD class function to the MTAP's 24hr form
'MTAP time format is simply, minutes after midnight (max 1440 minutes = 1 day)
'0 is midnight
'720 is midday
'1339 is 11:59pm

Dim H As Single
Dim T As Single
Dim m As Single

'get number
T = CSng(S)

'convert hour portion to minutes after midnight
H = Fix(T / 100) * 60

'get minutes portion
m = (T / 100 - Fix(T / 100)) * 100

'add
ConvertToMTAPTime = H + m

End Function
Private Sub Class_Initialize()
'set values for angles that define events
SinHO(1) = Sin(RADS * -0.833)      'sunset upper limb simple refraction
SinHO(2) = Sin(RADS * -6#)        'civil twi
SinHO(3) = Sin(RADS * -12#)       'nautical twi
SinHO(4) = Sin(RADS * -18#)       'astro twi

'time zone GMT-8
TimeZone = -8

End Sub



