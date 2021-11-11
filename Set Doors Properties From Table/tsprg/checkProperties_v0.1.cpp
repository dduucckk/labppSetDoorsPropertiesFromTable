/*******************************************************************************
***
***     Author: Ivan Matveev
***     email:  matveyev.ivan@gmail.com
***
***     Version: 1.0
***     Проверить соотвествие привязки двери к зоне в свойствах объекта и в таблице привязок в LabPP Automat
***
***     2021
***
*******************************************************************************/
// Комбинации зон и соответствущие типы дверей — в таблице «Зоны и двери_2.xlsx»




string sExcelGUIDs = "ЗОНЫ И ДВЕРИ.xlsx";
int iTableGUIDs; // номер таблицы с GUIDами и значениями. Заголовки в таблице Excel находятся в первой строке.
int tableRowsNumber;

int main()
{
    int ires;
    int icount;

    //--------------------------------------------------
    // СОЗДАЕМ ТАБЛИЦУ
    //--------------------------------------------------

    createTable();
    ts_table(iTableGUIDs, "get_rows_count", tableRowsNumber);
    cout << "\n";

    //--------------------------------------------------
    // ЗАГРУЖАЕМ СПИСОК ДВЕРЕЙ ДЛЯ ПРОВЕРКИ
    //--------------------------------------------------

    ac_request_special("load_elements_list_from_selection", 1, "DoorType", 2);

    ac_request("get_loaded_elements_list_count", 1); // считать количество элементов, находящихся в списке №1
    icount = ac_getnumvalue(); // получить в переменную icount результат предыдущей операции как число
    cout << "Количество выбранных дверей: " << icount << "\n"; // выдать сообщение о количестве элементов

    if (icount == 0)
    {
        // значит в списке нет элементов, поэтому сообщаем в окно сообщений и завершаем программу
        cout << "Нет элементов для обработки\n";
        return -1; // число здесь не важно, но обычно если возвращается отрицательное значение, значит программа не сделала то, что хотелось.
    }

    cout << "Список дверей загружен.\n";
    cout << "Проходим по выбранным дверям. \n";
    cout << "\n";




    // ..######..##....##..######..##.......########
    // .##....##..##..##..##....##.##.......##......
    // .##.........####...##.......##.......##......
    // .##..........##....##.......##.......######..
    // .##..........##....##.......##.......##......
    // .##....##....##....##....##.##.......##......
    // ..######.....##.....######..########.########

    //--------------------------------------------------
    // ПРОХОДИМ ПО ДВЕРЯМ
    //--------------------------------------------------


    for (int i = 0; i < icount; i++)
    {
        ac_request("set_current_element_from_list", 1, i); // установить текущим элемент из списка № 1 с индексом i

        //--------------------------------------------------
        // ВЫЯСНЯЕМ, К КАКИМ ЗОНАМ ОТНОСИТСЯ ЭЛЕМЕНТ
        //--------------------------------------------------

        cout << "Получаем зоны двери.\n";
        string sGUID,
               curObjFromGUID,
               curObjToGUID,
               curDBObjFromNumber,
               curDBObjToNumber,
               curDBObjFromCat,
               curDBObjToCat,
               curDBObjFromName,
               curDBObjToName;

        ac_request("get_element_value", "GuidAsText");   // считываем guid текущего элемента как текст
        sGUID = ac_getstrvalue();

        // ТУТ НАДО ПОЛУЧИТЬ ЗОНЫ ТЕКУЩЕЙ ДВЕРИ

        curObjFromGUID = getLinkedZoneGUID (sGUID, true); // получить айди зоны: true откуда ведет дверь, false — куда ведёт дверь
        curObjToGUID = getLinkedZoneGUID (sGUID, false);  // получить айди зоны: true откуда ведет дверь, false — куда ведёт дверь

        // ВЫЯСНЯЕМ ИМЯ И ПРОЧИЕ ПАРАМЕТРЫ ПОЛУЧЕНЫХ ЗОН

        if (curObjToGUID == "") { cout << "   Зона выхода не привязана.\n"; }
        if (curObjFromGUID == "") { cout << "   Зона входа не привязана.\n"; }


        //--------------------------------------------------
        // ПОЛУЧАЕМ ПАРАМЕТРЫ ТЕКУЩЕГО ОБЪЕКТА ИЗ БАЗЫ ДАННЫХ ПЛАГИНА LABPP
        //--------------------------------------------------

        // ЭТИ ДАННЫЕ ИСПОЛЬЗУЮТСЯ ДАЛЬШЕ,
        // СРАВНЕНИЕ ПРОИСХОДИТ ПО ПОЛЯМ ОБЪЕКТА И БАЗЕ ДАННЫХ ПЛАГИНА LABPP.
        // ПОСЛЕ ЭТОГО ПОДКЛЮЧАЕТСЯ ТАБЛИЦА И В НЕЙ ИЩЕТСЯ СООТВЕТСТВУЮЩИЙ ТИП ДВЕРИ


        curDBObjFromNumber = "";
        curDBObjFromName = "";
        curDBObjFromCat = "";

        curDBObjToNumber = "";
        curDBObjToName = "";
        curDBObjToCat = "";


        if (curObjToGUID != "")
        {
            ac_request("set_element_by_guidstr_as_current", curObjToGUID); // here is the set current element, beware
            ac_request("get_element_value", "ZoneNumber"); // номер зоны. фактически, не используется.
            curDBObjToNumber = ac_getstrvalue();
            ac_request("get_element_value", "ZoneName");
            curDBObjToName = ac_getstrvalue();
            ac_request("get_element_value", "ZoneCatCode");
            curDBObjToCat = ac_getstrvalue();
            cout << "    ZONE TO GUID = " << curObjToGUID << ".\n";
            cout << "    ZONE TO NAME = " << curDBObjToName << ".\n";
        }

        if (curObjFromGUID != "")
        {
            ac_request("set_element_by_guidstr_as_current", curObjFromGUID); // here is the set current element, beware
            // ac_request("get_element_value", "ZoneNumber");
            // curDBObjFromNumber = ac_getstrvalue();

            ac_request("get_element_value", "ZoneName");
            curDBObjFromName = ac_getstrvalue();

            ac_request("get_element_value", "ZoneCatCode");
            curDBObjFromCat = ac_getstrvalue();

            cout << "    ZONE FROM GUID = " << curObjFromGUID << ".\n";
            cout << "    ZONE FROM NAME = " << curDBObjFromName << ".\n";
        }

        // [ ZONE_FROM_CATEGORY | ZONE_FROM_NAME | ZONE_TO_CATEGORY | ZONE_TO_NAME | DOOR_CATEGORY ]

        //--------------------------------------------------
        // ПОЛУЧАЕМ СВОЙСТВА (ПАРАМЕТРЫ) ИЗ САМОЙ ДВЕРИ (ОБЪЕКТА)
        //--------------------------------------------------

        //
        // СВОЙСТВА ДВЕРИ ПО КЛАССИФИКАТОРУ:
        //
        // DOOR_CATEGORY — категория двери
        //
        // ZONE_FROM_NAME — имя зоны, откуда ведет дверь
        // ZONE_FROM_CATEGORY — категория зоны, откуда ведет дверь
        // ZONE_FROM_GUID — GUID зоны, откуда ведет дверь
        //
        // ZONE_TO_NAME — имя зоны, куда ведет дверь
        // ZONE_TO_CATEGORY — категория зоны, куда ведет дверь
        // ZONE_TO_GUID — GUID зоны, куда ведет дверь
        //
        //---------------------------------------

        string  curObjPropFromName,
                curObjPropFromCat,
                curObjPropFromGUID,
                curObjPropToName,
                curObjPropToCat,
                curObjPropToGUID,
                curDoorCategory;

        // ОБНУЛЕНИЕ — МАТЬ УЧЕНИЯ ©ГОРОДЕЦКИЙ
        curObjPropFromName = "";
        curObjPropFromCat = "";
        curObjPropFromGUID = "";
        curObjPropToName = "";
        curObjPropToCat = "";
        curObjPropToGUID = "";
        curDoorCategory = "";


        // ПОНЯСЛАСЬ, ПОЛУЧАЕМ СВОЙСТВА ОБЪЕКТА ИЗ ПОЛЕЙ ОБЪЕКТА
        ac_request("set_current_element_from_list", 1, i);
        ires = ac_request("elem_user_property", "get", "DOOR_CATEGORY");
        curDoorCategory = ac_getstrvalue();

        ires = ac_request("elem_user_property", "get", "ZONE_FROM_NAME");
        curObjPropFromName = ac_getstrvalue();
        ires = ac_request("elem_user_property", "get", "ZONE_FROM_CATEGORY");
        curObjPropFromCat = ac_getstrvalue();
        ires = ac_request("elem_user_property", "get", "ZONE_FROM_GUID");
        curObjPropFromGUID = ac_getstrvalue();

        ires = ac_request("elem_user_property", "get", "ZONE_TO_NAME");
        curObjPropToName = ac_getstrvalue();
        ires = ac_request("elem_user_property", "get", "ZONE_TO_CATEGORY");
        curObjPropToCat = ac_getstrvalue();
        ires = ac_request("elem_user_property", "get", "ZONE_TO_GUID");
        curObjPropToGUID = ac_getstrvalue();


        // ПРОВЕРКА НА НАЛИЧИЕ ПРОБЕЛОВ В СВОЙСТВАХ (ОН ЗАЧЕМ-ТО ПУСТЫЕ СВОЙСТВА КАК ПРОБЕЛ ЗАПИСЫВАЕТ, А НАМ ЭТО МЕШАЕТ В ПРОВЕРКЕ ДАЛЬШЕ)

        if (curObjPropFromName == " ") { // переменные из таблицы: тут обхожу жопу с пробелом
            curObjPropFromName = "";
        }
        if (curObjPropFromCat == " ") {
            curObjPropFromCat = "";
        }
        if (curObjPropToName == " ") {
            curObjPropToName = "";
        }
        if (curObjPropToCat == " ") {
            curObjPropToCat = "";
        }
        if (curDoorCategory == " ") {
            curDoorCategory = "";
        }


        //--------------------------------------------------
        // ПРОВЕРЯЕМ, ЕСЛИ ЗАПИСАННЫЕ В ОБЪЕКТ СВОЙСТВА
        // ОТЛИЧАЮТСЯ ОТ ТЕХ, ЧТО СОХРАНЕНЫ
        // В БАЗЕ ДАННЫХ ПЛАГИНА LABPP
        //
        // ЕСЛИ ОНИ ОТЛИЧАЮТСЯ, СТАВИМ ФЛАЖОК whaaaaa = TRUE,
        // ЭТО ОБОБЩЕННЫЙ ФЛАЖОК ОШИБКИ
        //--------------------------------------------------



        // ZONE_FROM_NAME — имя зоны, откуда ведет дверь
        // ZONE_FROM_CATEGORY — категория зоны, откуда ведет дверь
        // ZONE_FROM_GUID — GUID зоны, откуда ведет дверь
        //
        // ZONE_TO_NAME — имя зоны, куда ведет дверь
        // ZONE_TO_CATEGORY — категория зоны, куда ведет дверь
        // ZONE_TO_GUID — GUID зоны, куда ведет дверь
        //

        int e1, e2, e3, e4, e5; // ИНДЕКСЫ ОШИБКИ, ЧТОБЫ НА СЛЕДУЮЩИХ ЭТАПАХ МОЖНО БЫЛО СООБЩИТЬ ПОЛЬЗОВАТЕЛЮ, ЧТО ИМЕННО ЗАПИСАНО НЕВЕРНО.
        e1 = e2 = e3 = e4 = e5 = 0;

        bool whaaaaa = false;


        if (curObjPropFromName != curDBObjFromName) {
            cout << "       ! Cвойство ZONE_FROM_NAME:    obj" << curObjPropFromName << "» != db«" << curDBObjFromName << "»";
            cout << "\n";
            whaaaaa = true;
            e1 = 1;
        } else {
            cout << "   OK, Cвойство ZONE_FROM_NAME:    obj«" << curObjPropFromName << "» == db«" << curDBObjFromName << "»";
            cout << "\n";
        }

        if (curObjPropFromCat != curDBObjFromCat) {
            cout << "       ! Cвойство ZONE_FROM_CATEGORY:    obj«" << curObjPropFromCat << "» != db«" << curDBObjFromCat << "»";
            cout << "\n";
            whaaaaa = true;
            e2 = 1;
        } else {
            cout << "   OK, Cвойство ZONE_FROM_CATEGORY:    obj«" << curObjPropFromCat << "» == db«" << curDBObjFromCat << "»";
            cout << "\n";
        }

        if (curObjPropToName != curDBObjToName) {
            cout << "       ! Cвойство ZONE_TO_NAME:    obj«" << curObjPropToName << "» != db«" << curDBObjToName << "»";
            cout << "\n";
            whaaaaa = true;
            e3 = 1;
        } else {
            cout << "   OK, Cвойство ZONE_TO_NAME:    obj«" << curObjPropToName << "» == db«" << curDBObjToName << "»";
            cout << "\n";
        }

        if (curObjPropToCat != curDBObjToCat) {
            cout << "       ! Cвойство ZONE_TO_CATEGORY:    obj«" << curObjPropToCat << "» != db«" << curDBObjToCat << "»";
            cout << "\n";
            whaaaaa = true;
            e4 = 1;
        } else {
            cout << "   OK, Cвойство ZONE_TO_CATEGORY:    obj«" << curObjPropToCat << "» == db«"  << curDBObjToCat << "»";
            cout << "\n";
        }


        if (whaaaaa != true) { //  ЕСЛИ ВСЕ ПРОЧИЕ ПОЛЯ ЗАДАНЫ ВЕРНО, ИЩЕМ ТИП ДВЕРИ В ТАБЛИЦЕ

            // ..######..########....###....########...######..##.....##
            // .##....##.##.........##.##...##.....##.##....##.##.....##
            // .##.......##........##...##..##.....##.##.......##.....##
            // ..######..######...##.....##.########..##.......#########
            // .......##.##.......#########.##...##...##.......##.....##
            // .##....##.##.......##.....##.##....##..##....##.##.....##
            // ..######..########.##.....##.##.....##..######..##.....##

            //--------------------------------------------------
            // ПОИСК!
            //--------------------------------------------------
            // ЗАПРОС К ТАБЛИЦЕ, и сопоставление значений имен зон с категорией (типом) двери
            //--------------------------------------------------

            // int irow = ts_table(iTableGUIDs, "search", 0, currentObjectZoneCombination);

            // ZONE_FROM_NAME — имя зоны, откуда ведет дверь
            // ZONE_FROM_CATEGORY — категория зоны, откуда ведет дверь
            // ZONE_FROM_GUID — GUID зоны, откуда ведет дверь

            // ZONE_TO_NAME — имя зоны, куда ведет дверь
            // ZONE_TO_CATEGORY — категория зоны, куда ведет дверь
            // ZONE_TO_GUID — GUID зоны, куда ведет дверь


            string curObjFromCat, curObjToCat, curObjFromName, curObjToName;

            string curTableFromCat, curTableToCat,
                   curTableFromName, curTableToName;

            string curDoorCategoryFromTable, tempCatFromTable;

            curDoorCategoryFromTable = "";

            // ранее уже получили свойства текущего объекта :
            /*
            curObjPropFromName
            curObjPropFromCat
            curObjPropFromGUID

            curObjPropToName
            curObjPropToCat
            curObjPropToGUID

            curDoorCategory
            */

            curObjFromCat = tolower(curObjPropFromCat);
            curObjFromName = tolower(curObjPropFromName);
            curObjToCat = tolower(curObjPropToCat);
            curObjToName = tolower(curObjPropToName);

            curDoorCategory = tolower(curDoorCategory);

            //--------------------------------------------------
            // ПРОХОДИМ ПО ВСЕМ СТРОКАМ ТАБЛИЦЫ В ПОИСКАХ СООТВЕТСТВУЮЩЕЙ СТРОКИ
            //--------------------------------------------------

            for (int row = 0; row < tableRowsNumber; row++) // cycle through all table rows
            {
                cout << "       Row: " << row + 1 << " / " << tableRowsNumber << "\n";
                ts_table(iTableGUIDs, "select_row", row); // set the row

                curTableFromCat = ""; // сбрасываем переменные, чтобы не думать о возможных проблемах с повторами
                curTableFromName = "";
                curTableToCat = "";
                curTableToName = "";

                ts_table(iTableGUIDs, "get_value_of", 0, curTableFromCat);  // get current zone_from cat. in this row
                ts_table(iTableGUIDs, "get_value_of", 1, curTableFromName); // get current zone_from name in this row
                ts_table(iTableGUIDs, "get_value_of", 2, curTableToCat);    // get current zone_from name in this row
                ts_table(iTableGUIDs, "get_value_of", 3, curTableToName);   // get current zone_from name in this row
                ts_table(iTableGUIDs, "get_value_of", 4, tempCatFromTable);  // get current zone_from name in this row

                curTableFromCat = tolower(curTableFromCat);
                curTableFromName = tolower(curTableFromName);
                curTableToCat = tolower(curTableToCat);
                curTableToName = tolower(curTableToName);

                tempCatFromTable = tolower(tempCatFromTable);

                if (curTableFromCat == "*") { // переменные из таблицы: тут обхожу жопу с пробелом
                    curTableFromCat = "";
                }
                if (curTableFromName == "*") {
                    curTableFromName = "";
                }
                if (curTableToCat == "*") {
                    curTableToCat = "";
                }
                if (curTableToName == "*") {
                    curTableToName = "";
                }

                string currentObjectFromToCatName, currentTableFromToCatName;

                currentTableFromToCatName = "_|_" + curTableFromCat + "_|_" + curTableFromName + "_|_" + curTableToCat + "_|_" + curTableToName + "_|_";

                currentObjectFromToCatName  = "_|_" + curObjFromCat + "_|_" + curObjFromName + "_|_" + curObjToCat + "_|_" + curObjToName + "_|_";

                // cout << currentObjectFromToCatName << "\\\\\\" << currentTableFromToCatName << "\n";

                if (currentObjectFromToCatName == currentTableFromToCatName) { // ЕСЛИ
                    curDoorCategoryFromTable = tempCatFromTable;
                    curDoorCategoryFromTable = tolower(curDoorCategoryFromTable);
                }

            } // end of table for loop


            //--------------------------------------------------
            // ПРОВЕРКА НА СООТВЕТСТВИЕ ТИПУ ДВЕРИ
            //--------------------------------------------------

            if (curDoorCategoryFromTable != curDoorCategory) {
                cout << "       ! Cвойство DOOR_CATEGORY:    obj«" << curDoorCategory << "» != table«" << curDoorCategoryFromTable << "»";
                cout << "\n";
                whaaaaa = true;
                e5 = 1;
            } else {
                cout << "   OK, Cвойство DOOR_CATEGORY:    obj«" << curDoorCategory << "» == table«"  << curDoorCategoryFromTable << "»";
                cout << "\n";
            }



        } else { // ЕСЛИ КАКОЕ-ТО ПОЛЕ НЕ ЗАДАНО, ГОВОРИМ ПОЛЬЗОВАТЕЛЮ, ЧТО НАЙТИ ТИП ДВЕРИ НЕВОЗМОЖНО
            cout << "       Тип двери найти нельзя, пока прочие поля неверны.";
            cout << "\n";
        }





        //--------------------------------------------------
        // СВОЙСТВО ДЛЯ ГРАФЗАМЕНЫ
        //--------------------------------------------------

        if (whaaaaa == true) { // ЕСЛИ ОШИБКА В ОБЪЕКТЕ ВСЕ ЖЕ ЕСТЬ, ТО ЕЕ НАДО ОБРАБОТАТЬ

            // тут надо создать и задать свойство двери «верно или нет»,
            // чтобы показывать неверно заданные двери графзаменой
            //
            // ...

            int r0, ires;
            bool the_res = true;

            r0 = ac_request_special("get_element_value", "UP", "Wrong Category/Type");
            // cout << r0 << "\n";

            if (r0 == -1222) {  //-1222 — ошибка какая-то
                cout << "Одно из свойств не найдено, надо создать.\n";
                ires = ac_request("elem_user_property", "create", "Wrong Category/Type", " ", "Boolean", "OPENINGS");
                if (ires == 0) { cout << "Создали свойство Wrong Category/Type.\n"; } else { cout << "Не вышло создать свойство Wrong Category/Type.\n"; break();}
            }


            cout << "   Объект с ошибкой: " << sGUID << "\n";

            if (e1 == 1) {
                cout << "       Св. не совп.: " << sGUID << "\n";
            }


            ires = ac_request("elem_user_property", "set", "Wrong Category/Type", 1); // ПОСТАВИЛИ ФЛАЖОК ДЛЯ ЭЛЕМЕНТА С ОШИБКОЙ, ДАЛЬШЕ ЕГО ГРАФЗАМЕНОЙ ВЫДЕЛИМ В AC



            // ТУТ НАДО ПОКАЗАТЬ, КАКИЕ ИМЕННО СВОЙСТВА НЕВЕРНЫЕ В ОБЪЕКТЕ

            // <!-- НЕ СТАНУ ЭТО ДЕЛАТЬ, ДОСТАТОЧНО ГРАФЗАМЕНЫ
            // ДОБАВИТЬ ОБЪЕКТ С ОШИБКОЙ В СПИСОК 2

            /*
                        [16:33, 09/11/2021] Юрий Цепов: Добрый день!)
            Сейчас
            [16:36, 09/11/2021] Юрий Цепов: ac_request("store_current_element_to_list",2,-1);
                ac_request("select_elements_from_list",2,1);
                ac_request("Automate","ZoomToElements",2);
            [16:37, 09/11/2021] Юрий Цепов: ac_request("clear_list",2);
            [16:37, 09/11/2021] Юрий Цепов: сначала очистим список номер какой-то
            [16:37, 09/11/2021] Юрий Цепов: здесь - №2
            [16:38, 09/11/2021] Юрий Цепов: Потом когда перебираем элементы в цикле, например, ставим "store_current_element_to_list"
            [16:38, 09/11/2021] Юрий Цепов: т.е. текущий элемент если понравился - записываем его в список 2
            [16:39, 09/11/2021] Юрий Цепов: там -1 значит в конец списка
            [16:39, 09/11/2021] Юрий Цепов: Потом просто select_elements_from_list
            [16:39, 09/11/2021] Юрий Цепов: там номер списка 2 и 1/0 - снимать или нет текущее выделение
            [16:39, 09/11/2021] Юрий Цепов: при желании можем ZoomToElements сделать
            [16:40, 09/11/2021] Юрий Цепов: Если в 3d окне - работает сразу
            [16:40, 09/11/2021] Юрий Цепов: А в 2d окне надо сначала этаж подсветить
            [16:40, 09/11/2021] Юрий Цепов: ac_request("Environment","Story_GoTo",storeindex);
            [16:41, 09/11/2021] Юрий Цепов: storeindex - индекс этажа
            [16:41, 09/11/2021] Юрий Цепов: Это в самом начале ставим, конечно.
            Опять же если есть необходимость.
            [16:42, 09/11/2021] Юрий Цепов: А так можно список заполнять, дополнять и т.п.
            [16:43, 09/11/2021] Юрий Цепов: Можно сформировать таблицу ts_table с guid-ами и закинуть это в список 2

            */
// НЕ СТАНУ ЭТО ДЕЛАТЬ, ДОСТАТОЧНО ГРАФЗАМЕНЫ -->

        } else {
            ires = ac_request("elem_user_property", "set", "Wrong Category/Type", 0); // сбросить флаг
        }
        cout << "\n";
    }

    // ВЫДЕЛИТЬ ОБЪЕКТЫ ИЗ СПИСКА 2
    //
    //
    //
    //  НЕ СТАНУ ЭТОГО ДЕЛАТЬ, ДОСТАТОЧНО ГРАФЗАМЕНЫ
    //
    //

    deleteTable();

    cout << "\n";
    cout << "Программа отработала успешно.\n";
}




// .########.##.....##.##....##..######..########.####..#######..##....##..######.
// .##.......##.....##.###...##.##....##....##.....##..##.....##.###...##.##....##
// .##.......##.....##.####..##.##..........##.....##..##.....##.####..##.##......
// .######...##.....##.##.##.##.##..........##.....##..##.....##.##.##.##..######.
// .##.......##.....##.##..####.##..........##.....##..##.....##.##..####.......##
// .##.......##.....##.##...###.##....##....##.....##..##.....##.##...###.##....##
// .##........#######..##....##..######.....##....####..#######..##....##..######.


int LoadExcel()
{
    // ОШИБКА IDispatch GetIDsOfNames....
    // [17:56, 27/10/2021] Юрий Цепов: В двух случаях
    // [17:56, 27/10/2021] Юрий Цепов: 1 - если Excel находится в режиме редактирования ячейки
    // [17:57, 27/10/2021] Юрий Цепов: 2 - открыт еще какой-то Excel в виде процесса
    // [17:58, 27/10/2021] Юрий Цепов: Лечение - запустить диспетчер задач и выкинуть Excel с процессом
    // [17:58, 27/10/2021] Юрий Цепов: Еще вариант - Excel запустить от имени администратора
    // [17:58, 27/10/2021] Юрий Цепов: Виндовские заморочки

    int i, tsalert, res;
    string sguid, str;

    cout << "Загружаем данные из Excel\n";

    //  runtimecontrol("workline", "setpos", 0);

    // Очень странный способ загрузки таблицы: эксель Подключить Excel для работы.
    // В момент запуска программа Excel должна быть открыта.
    // По умолчанию активной становится текущая страница Excel.


    res = excel_attach();

    if (res != 0)
    {
        tsalert(-1, "Ошибка во время выполнения", "Нет подключения к excel", "");
        return -1;
    }


    res = excel_request("workbook_select", sExcelGUIDs);

    if (res != 0)
    {
        tsalert(-1, "Ошибка во время выполнения", "Не получается переключиться в файл excel", sExcelGUIDs);
        excel_detach();
        return -1;
    }

    ts_table(iTableGUIDs, "import_columns_from_excel", "A", 1, -1);  // с первой строки колонки А, до первой пустой ячейки
    ts_table(iTableGUIDs, "import_from_excel", "A", 2, -1, 0, 1);  // со второй строки колонки А, -1 — с текущей строки, ... , очистить таблицу перед добавлением

    excel_detach();

    cout << "Загрузка закончена.\n";

    //ts_table(iTableGUIDs, "print_to_str", str); // выгрузить всю таблицу в текстовую переменную для проверки
    //cout << "содержимое таблицы guids ->" << str << "\n"; // вывести в окно сообщений

    return 0;
}


string getLinkedZoneGUID(string sDoorGuid, bool bRoomFrom) // получить айди зоны: если true — откуда ведет дверь, если false — куда ведёт дверь
{
    string sGUID_linked_zone_guid = "";

    int flag1, flag2;
    if (bRoomFrom) { flag1 = 1024; flag2 = 2048; } else { flag1 = 4096; flag2 = 8192; }

    int iGuid;
    object("create", "ts_guid", iGuid);
    ts_guid(iGuid, "ConvertFromString", sDoorGuid);
    int iTableRoomsTmp;
    object("create", "ts_table", iTableRoomsTmp);

    ac_request_special("linkingElems", "getLinkedElemsByFlags", iGuid, flag2, iTableRoomsTmp);
    int icount;
    ts_table(iTableRoomsTmp, "get_rows_count", icount);
    if (icount > 0)
    {
        ts_table(iTableRoomsTmp, "select_row", 0);
        ts_table(iTableRoomsTmp, "get_value_of", "sGUID", sGUID_linked_zone_guid);
    }
    object("delete", iGuid);
    object("delete", iTableRoomsTmp);
    return sGUID_linked_zone_guid;
}

int createTable() {
    string sguid;   // переменная для guid
    string svalue;  // для строковых
    int ivalue;     // для числовых
    int jcount;
    int icount;
    int ires;
    object("create", "ts_table", iTableGUIDs); // создаем объект ts_table и получаем его номер для работы с ним
    ires = LoadExcel(); // загружаем данные из эксель (см. определение функции ниже)
    if (ires != 0) {
        cout << "Ошибка в ходе загрузки из Excel, программа завершена";
        return -1;
    }
    cout << "Создали таблицу. \n";
    return 0;
}

int deleteTable() {
    object("delete", iTableGUIDs);  // очищаем память
    cout << "Очистили память от таблицы. \n";
    return 0;
}