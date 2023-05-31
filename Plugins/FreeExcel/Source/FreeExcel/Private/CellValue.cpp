/**
*        Copyright(c) 2022  Weijian Tian
*
*    Permission is hereby granted, free of charge, to any person obtaining a copy
*    of this softwareand associated documentation files(the "Software"), to
*    deal in the Software without restriction, including without limitation the
*    rights to use, copy, modify, merge, publish, distribute, sublicense, and /or
*    sell copies of the Software, and to permit persons to whom the Software is
*    furnished to do so, subject to the following conditions :
*
*    The above copyright noticeand this permission notice shall be included in
*    all copies or substantial portions of the Software.
*
*   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
*    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
*    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.IN NO EVENT SHALL THE
*    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
*    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
*    FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
*    IN THE SOFTWARE.
*/


#include "CellValue.h"
 
 bool isLeapYear(int year) {

    if (year == 1900) return true;
    if (year % 400 == 0 || (year % 4 == 0 && year % 100 != 0))
        return true;
    return false;
}

 int daysInMonth(int month, int year)
{
    if (month > 12 || month < 1) return 0;
    if (month % 2)
    {
        return 31;
    }
    else if (month != 2)
    {
        return 30;
    }
    else
    {
        return (isLeapYear(year) ? 29 : 28);
    }
    /*  switch (month) {
      case 1:
          return 31;
      case 2:
          return (isLeapYear(year) ? 29 : 28);
      case 3:
          return 31;
      case 4:
          return 30;
      case 5:
          return 31;
      case 6:
          return 30;
      case 7:
          return 31;
      case 8:
          return 31;
      case 9:
          return 30;
      case 10:
          return 31;
      case 11:
          return 30;
      case 12:
          return 31;
      default:
          return 0;
      }*/
}

 double FCellValue::datetime_to_serial(FDateTime& dt)
{
    double serial = 1.0;
    int year = dt.GetYear() - 1900,
        mon = dt.GetMonth() - 1,
        day = dt.GetDay(),
        hour = dt.GetHour(),
        min = dt.GetMinute(),
        sec = dt.GetSecond(),
        millisec = dt.GetMillisecond();

    for (int i = 0; i < year; ++i) {
        serial += (isLeapYear(1900 + i) ? 366 : 365);
    }

    for (int i = 0; i < mon; ++i) {
        serial += daysInMonth(i + 1, year);
    }

    serial += day - 1;

    int32_t seconds = hour * 3600 + min * 60 + sec;
    serial += (seconds + millisec / 1000.0) / 86400.0;
    return serial;
}

 FDateTime FCellValue::serial_to_datetime(double serial)
{
    int year = 0, mon = 0, day = 0, hour = 0, min = 0, sec = 0, millisec = 0;

    while (true) {
        auto days = (isLeapYear(year + 1900) ? 366 : 365);
        if (days > serial) break;
        serial -= days;
        ++year;
    }

    // ===== Count the number of whole months in the year.
    while (true) {
        auto days = daysInMonth(mon + 1, year);
        if (days > serial) break;
        serial -= days;
        ++mon;
    }

    day = static_cast<int>(serial);
    serial -= day;

    hour = static_cast<int>(serial * 24);
    serial -= (hour / 24.0);

    min = static_cast<int>(serial * 24 * 60);
    serial -= (min / (24.0 * 60.0));

    auto tsec = std::round(serial * 24 * 60 * 60);
    sec = static_cast<int>(std::floor(tsec));
    millisec = static_cast<int>((tsec - std::floor(tsec)) * 1000);

    return FDateTime(year + 1900, mon + 1, day, hour, min, sec, millisec);
}