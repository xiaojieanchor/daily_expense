#define _CRT_SECURE_NO_WARNINGS
#include<stdio.h>
#include<windows.h>
#include <io.h>

extern int days_in_month[13];
extern int year, month, day;

int is_run_year(int y) {
    if ((((y % 4) == 0) && ((y % 100) != 0)) || (((y % 100) == 0) && ((y & 400) == 0))) {
        return 1;
    }
    else return 0;
}

int days_in_inityear(void) {
    int count = 0;
    for (int n = 12; n > month; n--) {
        count += days_in_month[n];
    }
    if (is_run_year(year)) {
        if (month <= 2) {
            count = count + days_in_month[month] - day + 2;
        }
        else {
            count = count + days_in_month[month] - day + 1;
        }
    }
    else {
        count = count + days_in_month[month] - day + 1;
    }
    return count;
}

int days_in_befmonth(int m) {
    int count = 0;
    for (int n = 1; n < m; n++) {
        count += days_in_month[n];
    }
    return count;
}

int ca_days(int y, int m, int d) {
    int days = 0;
    if (y == year) {
        const int mi_month = m;
        for (int n = (month+1); n < mi_month; n++) {
            days += days_in_month[n];
        }
        if (m == month) {
            days = d - day+1;
        }
        else {
            if (is_run_year(year)) {
                if (month <= 2) {
                    days = days + d + days_in_month[month] - day + 2;
                }
                else {
                    days = days + d + days_in_month[month] - day + 1;
                }
            }
            else {
                days = days + d + days_in_month[month] - day + 1;
            }
        }

    }
    else if (y > year) {
        for (int init_year = (year + 1); init_year <= y; init_year++) {
            if (init_year == y) {
                if (is_run_year(init_year)) {
                    if (m > 2) {
                        days += days_in_inityear() + days_in_befmonth(m) + d + 1;
                    }
                    else {
                        days += days_in_inityear() + days_in_befmonth(m) + d;
                    }
                }
                else {
                    days += days_in_inityear() + days_in_befmonth(m) + d;
                }
            }
            else {
                if (is_run_year(init_year)) {
                    days += days_in_befmonth(13) + 1;
                }
                else {
                    days += days_in_befmonth(13);
                }
            }
        }
    }
    return days;
}

void get_date(int* y, int* m, int* d) {
    printf("请输入今天的年份：");
    scanf_s("%d", y);
    printf("请输入今天的月份：");
    scanf_s("%d", m);
    printf("请输入今天的日期：");
    scanf_s("%d", d);
}
wchar_t* charToWChar(const char* str) {
    if (str == NULL) {
        return NULL;
    }

    // 计算所需的宽字符串缓冲区大小
    size_t len = mbstowcs(NULL, str, 0) + 1; // +1 为了空字符

    // 分配内存
    wchar_t* wstr = (wchar_t*)malloc(len * sizeof(wchar_t));
    if (wstr == NULL) {
        // 内存分配失败
        return NULL;
    }

    // 执行转换
    mbstowcs(wstr, str, len);

    return wstr;
}

int is_Valid_Date(int year1, int month1, int day1) {
    // 检查年份是否为正整数
    if (year1 <= 0) {
        return 0;
    }

    // 检查月份是否在1到12之间
    if (month1 < 1 || month1 > 12) {
        return 0;
    }

    // 检查日期是否为正整数
    if (day1 < 1) {
        return 0;
    }

    // 确定每个月的天数，考虑闰年
    const int daysInMonth[] = { 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
    if (month1 == 2) {
        // 2月的天数，考虑闰年
        if (is_run_year(year1)) {
            if (day1 > 29) {
                return 0;
            }
        }
        else {
            if (day1 > 28) {
                return 0;
            }
        }
    }
    else {
        // 其他月份的天数
        if (day1 > daysInMonth[month1 - 1]) {
            return 0;
        }
    }

    // 如果所有检查都通过，则日期有效
    return 1;
}

void excelSerialToDate(int serial, int* year, int* month, int* day) {
    // Excel的起始日期是1900年1月1日，对应的Excel日记计数是2
    // 但是Excel错误地将1900年当作闰年，所以需要减去2天
    serial -= 2; // 调整Excel的错误

    *year = 1900;
    int days = serial;
    while (1) {
        int daysInYear = is_run_year(*year) ? 366 : 365;
        if (days <= daysInYear) {
            break;
        }
        days -= daysInYear;
        (*year)++;
    }

    *month = 1;
    for (int i = 0; i < 11; i++) {
        int daysInMonth = (*month == 2 && is_run_year(*year)) ? 29 : (*month == 4 || *month == 6 || *month == 9 || *month == 11) ? 30 : 31;
        if (days > daysInMonth) {
            days -= daysInMonth;
            (*month)++;
        }
        else {
            break;
        }
    }

    *day = days ; // Excel日记计数中的日期是从1开始的
}
