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
    printf("������������ݣ�");
    scanf_s("%d", y);
    printf("�����������·ݣ�");
    scanf_s("%d", m);
    printf("�������������ڣ�");
    scanf_s("%d", d);
}
wchar_t* charToWChar(const char* str) {
    if (str == NULL) {
        return NULL;
    }

    // ��������Ŀ��ַ�����������С
    size_t len = mbstowcs(NULL, str, 0) + 1; // +1 Ϊ�˿��ַ�

    // �����ڴ�
    wchar_t* wstr = (wchar_t*)malloc(len * sizeof(wchar_t));
    if (wstr == NULL) {
        // �ڴ����ʧ��
        return NULL;
    }

    // ִ��ת��
    mbstowcs(wstr, str, len);

    return wstr;
}

int is_Valid_Date(int year1, int month1, int day1) {
    // �������Ƿ�Ϊ������
    if (year1 <= 0) {
        return 0;
    }

    // ����·��Ƿ���1��12֮��
    if (month1 < 1 || month1 > 12) {
        return 0;
    }

    // ��������Ƿ�Ϊ������
    if (day1 < 1) {
        return 0;
    }

    // ȷ��ÿ���µ���������������
    const int daysInMonth[] = { 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 };
    if (month1 == 2) {
        // 2�µ���������������
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
        // �����·ݵ�����
        if (day1 > daysInMonth[month1 - 1]) {
            return 0;
        }
    }

    // ������м�鶼ͨ������������Ч
    return 1;
}

void excelSerialToDate(int serial, int* year, int* month, int* day) {
    // Excel����ʼ������1900��1��1�գ���Ӧ��Excel�ռǼ�����2
    // ����Excel����ؽ�1900�굱�����꣬������Ҫ��ȥ2��
    serial -= 2; // ����Excel�Ĵ���

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

    *day = days ; // Excel�ռǼ����е������Ǵ�1��ʼ��
}
