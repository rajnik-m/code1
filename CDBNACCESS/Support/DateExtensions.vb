Imports System.Runtime.CompilerServices

''' <summary>
''' A module containing Date class extension methods
''' </summary>
Module DateExtensions

  ''' <summary>
  ''' Gets the next the working day.
  ''' </summary>
  ''' <param name="value">The starting date.</param>
  ''' <param name="environment">The environment.</param>
  ''' <returns>A new <see cref="Date" /> object representing the earliest date that is a working day on or after the date provided.</returns>
  <Extension()>
  Public Function NextWorkingDay(value As Date, environment As CDBEnvironment) As Date
    Return AdjustToWorkingDay(value, environment, 1)
  End Function

  ''' <summary>
  ''' Gets the previous the working day.
  ''' </summary>
  ''' <param name="value">The starting date.</param>
  ''' <param name="environment">The environment.</param>
  ''' <returns>A new <see cref="Date" /> object representing the latest date that is a working day on or before the date provided.</returns>
  <Extension()>
  Public Function PreviousWorkingDay(value As Date, environment As CDBEnvironment) As Date
    Return AdjustToWorkingDay(value, environment, -1)
  End Function

  ''' <summary>
  ''' Adjusts the supplied date by the amount given until the resulting date is a working day.
  ''' </summary>
  ''' <param name="value">The date.</param>
  ''' <param name="environment">The environment.</param>
  ''' <param name="adjustment">The amount to adjust by.</param>
  ''' <returns>A new <see cref="Date" /> object representing the adjusted date.</returns>
  Private Function AdjustToWorkingDay(value As Date, environment As CDBEnvironment, adjustment As Integer) As Date
    Dim result As Date = value
    While Not result.IsWorkingDay(environment)
      result = result.AddDays(adjustment)
    End While
    Return result
  End Function

  ''' <summary>
  ''' Determines whether the specified date is a working day.
  ''' </summary>
  ''' <param name="value">The date.</param>
  ''' <returns>True if the specified date fals in the weekend, otherwise False</returns>
  <Extension()>
  Public Function IsWeekend(value As Date) As Boolean
    Return value.DayOfWeek = DayOfWeek.Saturday OrElse value.DayOfWeek = DayOfWeek.Sunday
  End Function

  ''' <summary>
  ''' Determines whether the specified date is a bank holiday.
  ''' </summary>
  ''' <param name="value">The date.</param>
  ''' <param name="environment">The environment.</param>
  ''' <returns>True if the specified date is a bank holiday, otherwise False</returns>
  <Extension()>
  Public Function IsBankHoliday(value As Date, environment As CDBEnvironment) As Boolean
    Return New SQLStatement(environment.Connection,
                            "bank_holiday_day",
                            "bank_holiday_days",
                            New CDBFields({New CDBField("bank_holiday_day",
                                                        CDBField.FieldTypes.cftDate,
                                                        value.ToString(CAREDateFormat))}
                                         )
                           ).GetDataTable.Rows.Count > 0
  End Function

  ''' <summary>
  ''' Determines whether the specified date is a working day.
  ''' </summary>
  ''' <param name="value">The date.</param>
  ''' <param name="environment">The environment.</param>
  ''' <returns>True if the specified date is a working day, otherwise False</returns>
  <Extension()>
  Public Function IsWorkingDay(value As Date, environment As CDBEnvironment) As Boolean
    Return (Not value.IsWeekend) AndAlso (Not value.IsBankHoliday(environment))
  End Function
End Module
