#pragma once

int is_run_year(int y);

int days_in_befmonth(int m);

int ca_days(int y, int m, int d);
    
void get_date(int* y, int* m, int* d);

wchar_t* charToWChar(const char* str);

int is_Valid_Date(int year1, int month1, int day1);

void excelSerialToDate(int serial, int* year, int* month, int* day);