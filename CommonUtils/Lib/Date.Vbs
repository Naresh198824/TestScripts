'*********************************************************************************************************************
'**  Project            : 
'**  File Name          : Date.vbs
'**  Version            : 
'**  Created on         : 
'**  Updated on         : 
'**  Description        : Common Methods related to date
'**  Copyright          : Arsin Corporation.
'**  Author             :       
'*********************************************************************************************************************

'Class clsDate

Function GetHoliDays(dDate)
		
		' @HELP
		' @class	: clsDate
		' @method	: GetHoliDays(dDate)
		' @returns	: Array of date expressions
		' @parameter: dDate: Date expression.
		' @notes	: This method returns all the holidays in the year of the date.
		' @END

		Dim lYear
		Dim dtNewYearDay
		Dim dtIPDay
		Dim dtChristmasDay
		
		lYear = Year(dDate)
		dtNewYearDay="01/01/" & lYear
		dtIPDay="04/07/" & lYear
		dtChristmasDay="12/25/" & lYear
		GetHoliDays=Array(dtNewYearDay,dtIPDay,dtChristmasDay,MartinBirthday(lYear),PresidentDay(lYear),MemorialDay(lYear), LaborDay(lYear),ColumbusDay(lYear),ThanksgivingDay(lYear))
		
End Function
'--------------------------------------------------------------------------------------------------------------------
Function IsInHolidays(dDate)		
		
		' @HELP
		' @class	: clsDate
		' @method	: IsInHolidays(dDate)
		' @returns	: True/False
		' @parameter: dDate: Date expression.
		' @notes	: This method checks wether 'dDate' in holidays or not.
		' @END

		Dim iInner
		Dim aHoliDays
		
		aHoliDays=GetHoliDays(dDate)
		IsInHolidays=False
		
		For iInner=0 To UBound(aHolidays)
			If FormatDateTime(dDate,2)=FormatDateTime(aHoliDays(iInner),2) Then
				IsInHolidays=True
				Exit Function
			End If	
		Next
	
End Function
'--------------------------------------------------------------------------------------------------------------------
Function MartinBirthday(lYear)
	
		' @HELP
		' @class	: clsDate
		' @method	: MartinBirthday(lYear)
		' @returns	: Date Expression
		' @parameter: lYear:  Year .
		' @notes	: This method returns the Martin Luther King Birthday (3rd Monday of January) of the year.
		' @END
		
		Dim sFDay
		Dim iFWeekday
		
		sFDay="January 1, " & lyear 
		iFWeekday=Weekday(sFday)
		
		If iFWeekday > 2 then
			MartinBirthday=DateAdd("WW",2,DateAdd("d",9-iFWeekday,sFDay))
		Else
			MartinBirthday=DateAdd("WW",2,DateAdd("d",2-iFWeekday,sFDay))
		End If
		
End Function
'--------------------------------------------------------------------------------------------------------------------
Function PresidentDay(lYear)
	
		' @HELP
		' @class	: clsDate
		' @method	: PresidentDay(lYear)
		' @returns	: Date Expression
		' @parameter: lYear:  Year 
		' @notes	: This method returns the PresidentDay (3rd Monday of February) of the 'lYear'.
		' @END
			
		Dim sFDay
		Dim iFWeekday
		
		sFDay="February 1, " & lyear 
		iFWeekday=Weekday(sFday)
		
		If iFWeekday > 2 then
			PresidentDay=DateAdd("WW",2,DateAdd("d",9-iFWeekday,sFDay))
		Else
			PresidentDay=DateAdd("WW",2,DateAdd("d",2-iFWeekday,sFDay))
		End If
		
End Function
'--------------------------------------------------------------------------------------------------------------------
Function MemorialDay(lYear)
	
		' @HELP
		' @class	: clsDate
		' @method	: MemorialDay(lYear)
		' @returns	: Date Expression
		' @parameter: lYear:  Year 
		' @notes	: This method returns the MemorialDay(lYear)(Last Monday of May) of the 'lYear'.
		' @END
		
		Dim sFDay
		Dim iFWeekday
		
		sFDay="May 31, " & lyear 
		iFWeekday=Weekday(sFday)
		
		If iFWeekday >= 2 Then
			MemorialDay=DateAdd("d",2-iFWeekday,sFDay)
		Else
			MemorialDay=DateAdd("d",2-iFWeekday-7,sFDay)
		End If
		
End Function
'--------------------------------------------------------------------------------------------------------------------
Function LaborDay(lYear)
	
		' @HELP
		' @class	: clsDate
		' @method	: LaborDay(lYear)
		' @returns	: Date Expression
		' @parameter: lYear:  Year 
		' @notes	: This method returns the LaborDay (First Monday of September) of the 'lYear'.
		' @END
	
		Dim sFDay
		Dim iFWeekday
		
		sFDay="September 1, " & lyear 
		iFWeekday=Weekday(sFday)
		
		If iFWeekday > 2 then
			LaborDay=DateAdd("d",9-iFWeekday,sFDay)
		Else
			LaborDay=DateAdd("d",2-iFWeekday,sFDay)
		End If
	
End Function
'--------------------------------------------------------------------------------------------------------------------
Function ColumbusDay(lYear)
	
		' @HELP
		' @class	: clsDate
		' @method	: ColumbusDay(lYear)
		' @returns	: Date Expression
		' @parameter: lYear:  Year 
		' @notes	: This method returns the Columbus Day (2nd Monday of October) of the 'lYear'.
		' @END
			
		Dim sFDay
		Dim iFWeekday
		
		sFDay="October 1, " & lyear 
		iFWeekday=Weekday(sFday)
		
		If iFWeekday > 2 then
			ColumbusDay=DateAdd("WW",1,DateAdd("d",9-iFWeekday,sFDay))
		Else
			ColumbusDay=DateAdd("WW",1,DateAdd("d",2-iFWeekday,sFDay))
		End If
	
End Function
'--------------------------------------------------------------------------------------------------------------------
Function ThanksgivingDay(lYear)
		
		' @HELP
		' @class	: clsDate
		' @method	: ThanksgivingDay(lYear)
		' @returns	: Date Expression
		' @parameter: lYear:  Year 
		' @notes	: This method returns the Thanks Giving Day (4th Thursday of Novembrr) of the 'lYear'.
		' @END
	
	
		Dim sFDay
		Dim iFWeekday
		
		sFDay="November 1, " & lyear 
		iFWeekday=Weekday(sFday)
		
		If iFWeekday > 5 Then
			ThanksgivingDay=DateAdd("WW",3,DateAdd("d",12-iFWeekday,sFDay))
		Else
			ThanksgivingDay=DateAdd("WW",3,DateAdd("d",5-iFWeekday,sFDay))
		End If
	
End Function
'--------------------------------------------------------------------------------------------------------------------
Function Today()
	
		' @HELP
		' @class	: clsDate
		' @method	: Today()
		' @returns	: Date Expression
		' @parameter: None
		' @notes	: This method returns the todays date
		' @END
	
		Today = formatDateTime(Date,2)
End Function
 '--------------------------------------------------------------------------------------------------------------------
Function NextDay()
		
		' @HELP
		' @class	: clsDate
		' @method	: NextDay()
		' @returns	: Date Expression
		' @parameter: None
		' @notes	: This method returns the next date
		' @END
	
		NextDay = formatDateTime(DateAdd("d",1,Date),2)
End Function
'--------------------------------------------------------------------------------------------------------------------
Function GetPastORFutureDay(sInterval)
		
		' @HELP
		' @class	: clsDate
		' @method	: GetPastORFutureDay()
		' @returns	: Date Expression
		' @parameter: sInterval : The Required past or future date( EX: 5,10,-4,-13)
		' @notes	: This method returns the required past or future date based on parameter
		' @END
	
		GetPastORFutureDay = formatDateTime(DateAdd("d",sInterval,Date),2)
End Function
'--------------------------------------------------------------------------------------------------------------------

Function PreviousDay()
		
		' @HELP
		' @class	: clsDate
		' @method	: PreviousDay()
		' @returns	: Date Expression
		' @parameter: None
		' @notes	: This method returns the previous date
		' @END
	
	PreviousDay=formatDateTime(DateAdd("d",-1,Date),2)
End Function
'--------------------------------------------------------------------------------------------------------------------
Function NextBusinessDay(dDate)
	
		' @HELP
		' @class	: clsDate
		' @method	: NextBusinessDay(dDate)
		' @returns	: Date Expression
		' @parameter: dDate:  The date from which we will calculate the business date .
		' @notes	: This method returns the next business date from 'dDate'.
		' @END
		
		Select Case weekday(dDate)
		Case 1,2,3,4,5
			If IsInHolidays(FormatDateTime(DateAdd("d",1,Date),2)) Then
				NextBusinessDay=NextBusinessDay(FormatDateTime(DateAdd("d",1,Date),2))
			Else
				NextBusinessDay=FormatDateTime(DateAdd("d",1,Date),2)
			End If
		Case 6
			If IsInHolidays(FormatDateTime(DateAdd("d",3,Date),2)) Then
				NextBusinessDay=NextBusinessDay(FormatDateTime(DateAdd("d",3,Date),2))
			Else
				NextBusinessDay=FormatDateTime(DateAdd("d",3,Date),2)
			End If
		Case 7	
			If IsInHolidays(FormatDateTime(DateAdd("d",2,Date),2)) Then
				NextBusinessDay=NextBusinessDay(FormatDateTime(DateAdd("d",2,Date),2))
			Else
				NextBusinessDay=FormatDateTime(DateAdd("d",2,Date),2)
			End If
		End Select
	
End Function
'--------------------------------------------------------------------------------------------------------------------
Function NextHoliday(dDate)
	
		' @HELP
		' @class	: clsDate
		' @method	: NextHoliday(dDate)
		' @returns	: Date Expression
		' @parameter: dDate:  The date from which we will calculate the holiday date .
		' @notes	: This method returns the next Holiday date from 'dDate'.
		' @END
	
		'Dim iOuter,iInner,aHolidays
		Dim dNextday
		
		'aHolidays=GetHoliDays(Date)
		dNextday=FormatDateTime(DateAdd("d",1,dDate),2)
		If IsInHolidays(dNextday) Or WeekDay(dNextday)=7 Or WeekDay(dNextday)=1 Then
			NextHoliday=dNextday
			Exit Function
		End If		
		NextHoliday=NextHoliday(dNextday)
				
		'
		'Select Case weekday(Date)
		'Case 1
		'	For iOuter= 1 To 5
		'		For iInner=0 To UBound(aHolidays)
		'			If FormatDateTime(DateAdd("d",iOuter,Date),2)=FormatDateTime(aholiDays(iInner),2) Then
		'				NextHoliday=FormatDateTime(aholiDays(iInner),2) 
		'				Exit Sub
		'			End If	
		'		Next
		'	Next
		'	NextHoliday=FormatDateTime(DateAdd("d",6,Date),2)
		'Case 2
		'	For iOuter= 1 To 4
		'		For iInner=0 To UBound(aHolidays)
		'			If FormatDateTime(DateAdd("d",iOuter,Date),2)=FormatDateTime(aholiDays(iInner),2) Then
		'				NextHoliday=FormatDateTime(aholiDays(iInner),2) 
		'				Exit Sub
		'			End If	
		'		Next
		'	Next
		'	NextHoliday=FormatDateTime(DateAdd("d",5,Date),2)
		'Case 3
		'	For iOuter= 1 To 3
		'		For iInner=0 To UBound(aHolidays)
		'			If FormatDateTime(DateAdd("d",iOuter,Date),2)=FormatDateTime(aholiDays(iInner),2) Then
		'				NextHoliday=FormatDateTime(aholiDays(iInner),2) 
		'				Exit Sub
		'			End If
		'		Next
		'	Next
		'
		'	NextHoliday=FormatDateTime(DateAdd("d",4,Date),2)
		'Case 4
		'	For iOuter= 1 To 2
		'		For iInner=0 To UBound(aHolidays)
		'			If FormatDateTime(DateAdd("d",iOuter,Date),2)=FormatDateTime(aholiDays(iInner),2) Then
		'				NextHoliday=FormatDateTime(aholiDays(iInner),2) 
		'				Exit Sub
		'			End If
		'		Next
		'	Next
		'
		'	NextHoliday=FormatDateTime(DateAdd("d",3,Date),2)
		'Case 5
		'	For iInner=0 To UBound(aHolidays)
		'		If FormatDateTime(DateAdd("d",1,Date),2)=FormatDateTime(aholiDays(iInner),2) Then
		'			NextHoliday=FormatDateTime(aholiDays(iInner),2) 
		'			Exit Sub
		'		End If
		'	Next
		'	NextHoliday=FormatDateTime(DateAdd("d",2,Date),2)
		'Case 6
		'	NextHoliday=FormatDateTime(DateAdd("d",1,Date),2)
		'Case 7	
		'	NextHoliday=FormatDateTime(DateAdd("d",1,Date),2)
		'End Select
End Function
'--------------------------------------------------------------------------------------------------------------------
Function DateAfterDays(iInterval)
	
		' @HELP
		' @class	: clsDate
		' @method	: DateAfterDays(iInterval)
		' @returns	: Date Expression
		' @parameter: iInterval:  Numeric expression that is the number of days after.
		' @notes	: This method  returns the  date after iInterval days of current Date.
		' @END
	
		DateAfterMonth = FormatDateTime(DateAdd("d",iInterval,Date),2)
	
End Function
'--------------------------------------------------------------------------------------------------------------------
Function DateAfterWeek(iInterval)
	
		' @HELP
		' @class	: clsDate
		' @method	: DateAfterWeek(iInterval)
		' @returns	: Date Expression
		' @parameter: iInterval:  Numeric expression that is the number of weeks after .
		' @notes	: This method returns the date after iInterval weeks of current Date.
		' @END
	
	
		DateAfterWeek = FormatDateTime(DateAdd("WW",iInterval,Date),2)
		
End Function
'--------------------------------------------------------------------------------------------------------------------
Function DateAfterMonth(iInterval)
	
		' @HELP
		' @class	: clsDate
		' @method	: DateAfterMonth(iInterval)
		' @returns	: Date Expression
		' @parameter: iInterval:  Numeric expression that is the number of months after .
		' @notes	: This method returns the  date after iInterval months of current Date.
		' @END
	
		DateAfterMonth = FormatDateTime(DateAdd("m",iInterval,Date),2)
		
End Function
'--------------------------------------------------------------------------------------------------------------------
Function DateAfterYear(iInterval)
	
		' @HELP
		' @class	: clsDate
		' @method	: DateAfterYear(iInterval)
		' @returns	: Date Expression
		' @parameter: iInterval:  Numeric expression that is the number of years after .
		' @notes	: This method  returns the next  date after iInterval years  of current Date.
		' @END
	
		DateAfterYear = FormatDateTime(DateAdd("YYYY",iInterval,Date),2)
		
End Function
'--------------------------------------------------------------------------------------------------------------------
Function DateBeforeDays(iInterval)
	
		' @HELP
		' @class	: clsDate
		' @method	: DateBeforeDays(iInterval)
		' @returns	: Date Expression
		' @parameter: iInterval:  Numeric expression that is the number of days before.
		' @notes	: This method  returns the  date before iInterval days of current Date.
		' @END
		
		DateAfterMonth = FormatDateTime(DateAdd("d",-1 * iInterval,Date),2)
		
End Function
'--------------------------------------------------------------------------------------------------------------------
Function DateBeforeWeek(iInterval)
		
		' @HELP
		' @class	: clsDate
		' @method	: DateBeforeWeek(iInterval)
		' @returns	: Date Expression
		' @parameter: iInterval:  Numeric expression that is the number of weeks before.
		' @notes	: This method  returns the  date before iInterval weeks  of current Date.
		' @END
	
		DateBeforeWeek=FormatDateTime(DateAdd("WW",-1 * iInterval,Date),2)
		
End Function
'--------------------------------------------------------------------------------------------------------------------
Function DateBeforeMonth(iInterval)
	
		' @HELP
		' @class	: clsDate
		' @method	: DateBeforeMonth(iInterval)
		' @returns	: Date Expression
		' @parameter: iInterval:  Numeric expression that is the number of months before .
		' @notes	: This method  returns the  date before iInterval months  of current Date.
		' @END
	
		DateBeforeMonth = FormatDateTime(DateAdd("m",-1 * iInterval,Date),2)
		
End Function
'--------------------------------------------------------------------------------------------------------------------
Function DateBeforeYear(iInterval)
	
		' @HELP
		' @class	: clsDate
		' @method	: DateBeforeYear(iInterval)
		' @returns	: Date Expression
		' @parameter: iInterval:  Numeric expression that is the number of years bfore .
		' @notes	: This method  returns the  date before iInterval years  of current Date.
		' @END
	
		DateBeforeYear = FormatDateTime(DateAdd("YYYY",-1 * iInterval,Date),2)
		
End Function
'--------------------------------------------------------------------------------------------------------------------
'End Class
