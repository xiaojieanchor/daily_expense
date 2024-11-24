#define _CRT_SECURE_NO_WARNINGS
#include<stdio.h>
#include <stdlib.h>
#include <wchar.h>
#include <iostream>
#include <windows.h>
#include "libxl.h"
#include"new_fun.h"
#include <string>
#include <locale>
#include <codecvt>
#include <iomanip>
#include <sstream>
#include <shlobj.h>
#include <io.h>

using namespace libxl;
int days_in_month[13] = { 0, 31,28,31,30,31,30,31,31,30,31,30,31 };
char volume[] = "������ˮˮ��յ���������";
double cost_in_volume[6];
int year=0, month=0, day=0;

int main() {

    char desktopPath[MAX_PATH];
    SHGetFolderPathA(NULL, CSIDL_DESKTOP, NULL, 0, desktopPath);
    strcat(desktopPath, "\\daily_expense.xlsx");
    if (_access(desktopPath, 0) != 0) {
        Book* book = xlCreateXMLBookW();
        book->setKey(L"Cabbage", L"windows-202029010bc7ec066cbe6b67a2p5r4g1");
        if (book) {
            Format* dateFormat = book->addFormat();
            dateFormat->setNumFormat(NUMFORMAT_DATE);
            Sheet* sheet = book->addSheet(L"Custom formats");
            if (sheet) {

                //��volume�ı༭

                sheet->writeStr(2, 0, L"����");
                sheet->writeStr(3, 0, L"��ˮ");
                sheet->writeStr(4, 0, L"ˮ��");
                sheet->writeStr(5, 0, L"�յ�");
                sheet->writeStr(6, 0, L"����");
                sheet->writeStr(7, 0, L"����");

                // �û������ʼ����
                back3:
                printf("��������˵���ʼ���� (��ʽ: yyyy/mm/dd): ");
                scanf_s("%d/%d/%d", &year,&month,&day); 
                double dateValue = book->datePack(year, month, day);
                sheet->writeNum(1, 1, dateValue, dateFormat); 
                sheet->setCol(1, 356, 10);

                // �����û����������
                if (is_Valid_Date(year,month,day) ){
                    //// ���û����������ת��ΪExcel���ڸ�ʽ
                    //double dateValue = book->datePack(year, month, day);
                    //// д���ʼ����
                    //sheet->writeNum(0, 1, dateValue, dateFormat); // д��A1��Ԫ��
                    //sheet->setCol(1, 356, 10);
                    //// ѭ��д�������365������
                    //for (int i = 1; i <= 365; ++i) {
                    //    // ����ʼ��������һ��
                    //    dateValue = book->datePack(year, month, day + i);
                    //    sheet->writeNum(0, i + 1, dateValue, dateFormat); // ��B1��ʼд��
                    //}
                }
                else {
                    printf("���ڸ�ʽ������ȷ����ʽΪ yyyy/mm/dd\n");
                    goto back3;
                }
            }
            wchar_t* Path = charToWChar(desktopPath);
            if (book->save(Path)) {
                printf("�˱��������\n");
            }
            book->release();
        }
        
    }
    wchar_t*Path= charToWChar(desktopPath);


    Book* book = xlCreateXMLBookW();
    book->setKey(L"Cabbage", L"windows-202029010bc7ec066cbe6b67a2p5r4g1");
    if (book->load(Path)) {
        printf("*�ļ��򿪳ɹ�*\n");
        Format* dateFormat = book->addFormat();
        dateFormat->setNumFormat(NUMFORMAT_DATE);
        Sheet* sheet = book->getSheet(0);
        if (sheet) {
            printf("*�������򿪳ɹ�*\n");
            int y = 0, m = 0, d = 0, days = 0, mode = 0, n=0,key=0;
        init:

            //��ȡ�û������ʼ����
            excelSerialToDate(sheet->readNum(1, 1), &year, &month, &day);

            //��ȡ��¼���յ�����

            get_date(&y, &m, &d);
            if (!is_Valid_Date(y, m, d) || ((y == year) && (m <= month) && (d < day))) {
                printf("����������������������\n");
                goto init;
            }

            //ѡ������ģʽ
            back1:
            printf("ѡ����������(1-�����룻2-ȫ������)��");
            scanf_s("%d", &mode);
            if ((mode != 1) && (mode != 2)) {
                printf("����������������������\n");
                goto back1;
            }

            //�����¼�������һ�ε�����֮��Ĳ��������ں���¼��

            days = ca_days(y, m, d);
            
            //����һ���Ƿ������ڣ�û�еĻ����Զ�����

            if (fabs(sheet->readNum(1, days))<0.01) {
                double dateValue = book->datePack(y, m, d);
                sheet->writeNum(1, days, dateValue, dateFormat); 
                sheet->setCol(days, days, 10);
            }

            //����λ���������ݣ�ȷ��û������Ĵ���

            for (n = 1; n <= 6; n++) {
                double value = sheet->readNum(n + 2, days);
                if (fabs(value) >= 0.01) {
                    key = 1;
                }
            }
            if (key == 1) {
                back2:
                printf("��ǰλ���������ݣ��Ƿ�Ҫ����(��-0����-1)��");
                int if_cover = 1;
                scanf_s("%d", &if_cover);
                if (if_cover == 0) {
                    goto end;
                }
                if ((if_cover != 1) && (if_cover != 0)) {
                    printf("����������������������\n");
                    goto  back2;
                }
            }

            //"������"��ȡÿ��volume�ķ���

            if (mode == 1) {
                printf("�����%c%c%c%c�����ǣ�", volume[0], volume[1], volume[2], volume[3]);
                scanf_s("%lf", &cost_in_volume[0]);
                printf("�����%c%c%c%c�����ǣ�", volume[4], volume[5], volume[6], volume[7]);
                scanf_s("%lf", &cost_in_volume[1]);
            }

            //"ȫ������"��ȡÿ��volume�ķ���

            else if(mode==2){
                for (n = 0; n <= 20; n += 4) {
                    printf("�����%c%c%c%c�����ǣ�", volume[n], volume[n + 1], volume[n + 2], volume[n + 3]);
                    scanf_s("%lf", &cost_in_volume[n / 4]);
                }
            }

            //����������ݴ���excel���ڶ�����Щ���⣬��Ҫд������xxx������ʽ

            for (n = 0; n <= 5; n++) {
                if (n == 1) {
                    if(cost_in_volume[1] >= 0){
                    RichString* richString = book->addRichString();
                    Format* format = book->addFormat();
                    if (richString) {
                        format->setAlignH(ALIGNH_RIGHT);
                        richString->addText(L"��");
                        std::ostringstream oss;
                        oss << std::fixed << std::setprecision(2) << cost_in_volume[1];
                        std::string cost_str = oss.str();
                        std::wstring cost_wstr(cost_str.begin(), cost_str.end());
                        richString->addText(cost_wstr.c_str());
                        sheet->writeRichStr(n + 2, days, richString, format);
                    }
                    }
                    else {
                        sheet->writeNum(n + 2, days, -cost_in_volume[n]);
                    }
                }
                else {
                    sheet->writeNum(n + 2, days, cost_in_volume[n]);
                }
            }

            //ԭ����Ҫ�����ܹ��Զ�ע�͵İ��ģ����������ǲ�̫����

         /*   if (mode == 2) {
                back:
                printf("������Ҫע�͵ĵ�Ԫ��(����-1,��ˮ-2,ˮ��-3,�յ�-4,����-5,����-6,��-0)��");
                scanf_s("%d", &n);
                if (n != 0) {
                    printf("����ע�͵����ݣ�");
                    char comment1[256];
                    scanf_s("%99s",comment1, (unsigned)_countof(comment1));
                    wchar_t comment[256];
                    mbstowcs(comment, comment1, _countof(comment));
                    sheet->writeComment(n + 2, days, comment, L"��");
                    goto back;
                }
            }*/
        }
        else {
            printf("�޷����ع�����.\n");
    }
        
        }
    else {
        printf("�ĵ�������");
        book->release();
        return 0;
    }
    if (book->save(Path)) {
        /*::ShellExecuteW(NULL, L"open", L"C:\\Users\\19410\\Desktop\\ÿ��֧��.xlsx", NULL, NULL, SW_SHOW);*/
    }
    end:
    book->release();
    return 0;
}
