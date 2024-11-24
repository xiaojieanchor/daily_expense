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
char volume[] = "餐饮热水水电空调出行其他";
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

                //对volume的编辑

                sheet->writeStr(2, 0, L"餐饮");
                sheet->writeStr(3, 0, L"热水");
                sheet->writeStr(4, 0, L"水电");
                sheet->writeStr(5, 0, L"空调");
                sheet->writeStr(6, 0, L"出行");
                sheet->writeStr(7, 0, L"其他");

                // 用户输入初始日期
                back3:
                printf("请输入记账的起始日期 (格式: yyyy/mm/dd): ");
                scanf_s("%d/%d/%d", &year,&month,&day); 
                double dateValue = book->datePack(year, month, day);
                sheet->writeNum(1, 1, dateValue, dateFormat); 
                sheet->setCol(1, 356, 10);

                // 解析用户输入的日期
                if (is_Valid_Date(year,month,day) ){
                    //// 将用户输入的日期转换为Excel日期格式
                    //double dateValue = book->datePack(year, month, day);
                    //// 写入初始日期
                    //sheet->writeNum(0, 1, dateValue, dateFormat); // 写入A1单元格
                    //sheet->setCol(1, 356, 10);
                    //// 循环写入后续的365天日期
                    //for (int i = 1; i <= 365; ++i) {
                    //    // 将初始日期增加一天
                    //    dateValue = book->datePack(year, month, day + i);
                    //    sheet->writeNum(0, i + 1, dateValue, dateFormat); // 从B1开始写入
                    //}
                }
                else {
                    printf("日期格式错误，请确保格式为 yyyy/mm/dd\n");
                    goto back3;
                }
            }
            wchar_t* Path = charToWChar(desktopPath);
            if (book->save(Path)) {
                printf("账本创建完毕\n");
            }
            book->release();
        }
        
    }
    wchar_t*Path= charToWChar(desktopPath);


    Book* book = xlCreateXMLBookW();
    book->setKey(L"Cabbage", L"windows-202029010bc7ec066cbe6b67a2p5r4g1");
    if (book->load(Path)) {
        printf("*文件打开成功*\n");
        Format* dateFormat = book->addFormat();
        dateFormat->setNumFormat(NUMFORMAT_DATE);
        Sheet* sheet = book->getSheet(0);
        if (sheet) {
            printf("*工作簿打开成功*\n");
            int y = 0, m = 0, d = 0, days = 0, mode = 0, n=0,key=0;
        init:

            //读取用户输入初始日期
            excelSerialToDate(sheet->readNum(1, 1), &year, &month, &day);

            //获取记录当日的日期

            get_date(&y, &m, &d);
            if (!is_Valid_Date(y, m, d) || ((y == year) && (m <= month) && (d < day))) {
                printf("输入数据有误，请重新输入\n");
                goto init;
            }

            //选择输入模式
            back1:
            printf("选择输入类型(1-简单输入；2-全局输入)：");
            scanf_s("%d", &mode);
            if ((mode != 1) && (mode != 2)) {
                printf("输入数据有误，请重新输入\n");
                goto back1;
            }

            //计算记录日期与第一次的日期之间的差数，便于后续录入

            days = ca_days(y, m, d);
            
            //检查第一行是否有日期，没有的话就自动填入

            if (fabs(sheet->readNum(1, days))<0.01) {
                double dateValue = book->datePack(y, m, d);
                sheet->writeNum(1, days, dateValue, dateFormat); 
                sheet->setCol(days, days, 10);
            }

            //检查该位置有无数据，确保没有输入的错误

            for (n = 1; n <= 6; n++) {
                double value = sheet->readNum(n + 2, days);
                if (fabs(value) >= 0.01) {
                    key = 1;
                }
            }
            if (key == 1) {
                back2:
                printf("当前位置已有数据，是否要覆盖(否-0，是-1)：");
                int if_cover = 1;
                scanf_s("%d", &if_cover);
                if (if_cover == 0) {
                    goto end;
                }
                if ((if_cover != 1) && (if_cover != 0)) {
                    printf("输入数据有误，请重新输入\n");
                    goto  back2;
                }
            }

            //"简单输入"获取每个volume的费用

            if (mode == 1) {
                printf("今天的%c%c%c%c消费是：", volume[0], volume[1], volume[2], volume[3]);
                scanf_s("%lf", &cost_in_volume[0]);
                printf("今天的%c%c%c%c消费是：", volume[4], volume[5], volume[6], volume[7]);
                scanf_s("%lf", &cost_in_volume[1]);
            }

            //"全局输入"获取每个volume的费用

            else if(mode==2){
                for (n = 0; n <= 20; n += 4) {
                    printf("今天的%c%c%c%c消费是：", volume[n], volume[n + 1], volume[n + 2], volume[n + 3]);
                    scanf_s("%lf", &cost_in_volume[n / 4]);
                }
            }

            //将输入的数据存入excel，第二项有些特殊，需要写出“花xxx”的形式

            for (n = 0; n <= 5; n++) {
                if (n == 1) {
                    if(cost_in_volume[1] >= 0){
                    RichString* richString = book->addRichString();
                    Format* format = book->addFormat();
                    if (richString) {
                        format->setAlignH(ALIGNH_RIGHT);
                        richString->addText(L"花");
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

            //原本想要做出能够自动注释的板块的，看起来还是不太行呢

         /*   if (mode == 2) {
                back:
                printf("输入想要注释的单元格(餐饮-1,热水-2,水电-3,空调-4,出行-5,其他-6,无-0)：");
                scanf_s("%d", &n);
                if (n != 0) {
                    printf("输入注释的内容：");
                    char comment1[256];
                    scanf_s("%99s",comment1, (unsigned)_countof(comment1));
                    wchar_t comment[256];
                    mbstowcs(comment, comment1, _countof(comment));
                    sheet->writeComment(n + 2, days, comment, L"杰");
                    goto back;
                }
            }*/
        }
        else {
            printf("无法加载工作簿.\n");
    }
        
        }
    else {
        printf("文档不存在");
        book->release();
        return 0;
    }
    if (book->save(Path)) {
        /*::ShellExecuteW(NULL, L"open", L"C:\\Users\\19410\\Desktop\\每日支出.xlsx", NULL, NULL, SW_SHOW);*/
    }
    end:
    book->release();
    return 0;
}
