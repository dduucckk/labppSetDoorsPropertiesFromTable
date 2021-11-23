/* encoding windows-1251 */
// Задать свойства дверям в соответствии с прилежащими зонами.
// LABPP AUTOMAT скрипт
// Иван Матвеев
// 2021
//
//
//
// Комбинации зон — в таблице «Зоны и двери.xlsx»
//
// СТРУКТУРА ТАБЛИЦЫ:
// [ ZONE_FROM_CATEGORY | ZONE_FROM_NAME | ZONE_TO_CATEGORY | ZONE_TO_NAME | DOOR_CATEGORY | ... ]
// [ 7 | zone 1 | 3 | zone 5 | 015 | ... ]


//--------------------------------------------------
// ЗАДАТЬ СВОЙСТВА ДВЕРЯМ ИЗ ТАБЛИЦЫ
//--------------------------------------------------

string sExcelGUIDs = "ЗОНЫ И ДВЕРИ.xlsx";
string sCSVFilepath = "D:\\table.csv";
int iTableGUIDs; // номер таблицы с GUIDами и значениями. Заголовки в таблице Excel находятся в первой строке.
int tableRowsNumber, tableColsNumber;

int main()
{
	int icount;
  int iFileDescr;
  object("create", "ts_file", iFileDescr);
  int ires = ts_file(iFileDescr, "open", sCSVFilepath, "ignore", "r");
  if (ires != 0)
  {
      // Это копирование файла от Юрия. Обещал прислать коды ошибок
      // надо всю конструкцию в отдельную функцию перенести бы
      // https://www.labpp.ru/TSHelp/LabPP_Automat/RUS/copy.html
      // int err = shell_func("file","copy","..\\table.csv",sCSVFilepath);
      // coutvar << err;
      ac_request("dialog_get_filename", "Файл таблицы помещений", "csv", "", sCSVFilepath);
  }
  ts_file(iFileDescr, "close");
  object("delete", iFileDescr);

	//--------------------------------------------------
	// ЗАГРУЖАЕМ ДВЕРИ (ВСЕ)
	//--------------------------------------------------




	// ..######..########.##.......########..######..########
	// .##....##.##.......##.......##.......##....##....##...
	// .##.......##.......##.......##.......##..........##...
	// ..######..######...##.......######...##..........##...
	// .......##.##.......##.......##.......##..........##...
	// .##....##.##.......##.......##.......##....##....##...
	// ..######..########.########.########..######.....##...


	//--------------------------------------------------
	// ПОМЕНЯТЬ НА ЗАГРУЖАЕМ ДВЕРИ ИЗ ВЫБРАННЫХ !!!!
	//--------------------------------------------------

	// ЗАГРУЗИТЬ ДВЕРИ ИЗ СПИСКА ВЫБРАННЫХ ЭЛЕМЕНТОВ
	ac_request_special("load_elements_list_from_selection", 1, "DoorType", 2);
        icount = ac_getnumvalue();
        if (icount == 0)
        {
            // ЗАГРУЗИТЬ ВСЕ ДВЕРИ ИЗ ПРОЕКТА
	    ac_request("load_elements_list", 1, "DoorType", "MainFilter", 2);
            cout << "Поскольку ничего не выделено, обрабатываю все двери в проекте.";
        }


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

	checkAndCreateUserParameters(); // проверка параметров и создание недостающих, если вдруг


	//--------------------------------------------------
	// СОЗДАЕМ ТАБЛИЦУ
	//--------------------------------------------------

	createTable();
	ts_table(iTableGUIDs, "get_rows_count", tableRowsNumber);
	ts_table(iTableGUIDs, "get_columns_count", tableColsNumber);
	cout << tableRowsNumber << "  - количество строк файла!!! Это смешно!\n";


	// КОЛИЧЕСТВО КОЛОНОК ВЫЯСНЯЕМ ТУТ, ЧТОБЫ НЕ ДЕЛАТЬ ЭТО В ЦИКЛЕ
	int tableExtraColumnsNumber = tableColsNumber - 5;

	// ДАЛЬШЕ – ТАНЦЫ С БУБНОМ, ПОСКОЛЬКУ НЕТ МАССИВОВ
	// ПРОХОДИМ ПО ДОПОЛНИТЕЛЬНЫМ ПОЛЯМ ТАБЛИЦЫ
	// ЭТО НЕ НАДО, СДЕЛАЮ В ЦИКЛЕ, ВСЕ РАВНО
	/*
	string temp;

	for (int i = tableExtraColumnsNumber + 1; i < tableColsNumber; i++) { // ИДЁМ ПО ДОПОЛНИТЕЛЬНЫМ КОЛОНКАМ
		cout << i << " ";
		// ts_table(iTableGUIDs, "select_row", 0); // set the row
		ts_table(iTableGUIDs, "get_heading_of", i, temp); //  ТУТ НАДО ПОЛУЧИТЬ ИМЯ КОЛОНКИ!
		var_by_txt("init", "tableExtraColumnName" + itoa(i), "string", "global", temp);
		cout << var_by_txt("get", "tableExtraColumnName" + itoa(i)) << "\n";
	}
	*/

	// ..######..##....##..######..##.......########
	// .##....##..##..##..##....##.##.......##......
	// .##.........####...##.......##.......##......
	// .##..........##....##.......##.......######..
	// .##..........##....##.......##.......##......
	// .##....##....##....##....##.##.......##......
	// ..######.....##.....######..########.########

	//--------------------------------------------------
	// ЦИКЛ ПО ДВЕРЯМ
	//--------------------------------------------------

	cout << "Проходим по выбранным дверям...\n";
	cout << "\n";

	for (int i = 0; i < icount; i++)
	{
		ac_request("set_current_element_from_list", 1, i); // установить текущим элемент из списка № 1 с индексом i

		//--------------------------------------------------
		// ВЫЯСНЯЕМ, К КАКИМ ЗОНАМ ОТНОСИТСЯ ЭЛЕМЕНТ
		//--------------------------------------------------

		cout << "Получаем зоны двери...\n";
		string sText, sguid, curObjFromGUID, curObjToGUID, sIDdoor, curObjFromNumber, curObjToNumber, curObjFromCat, curObjToCat, curObjFromName, curObjToName;
		string curObjToCatName, curObjFromCatName;



		ac_request("get_element_value", "GuidAsText");   // считываем guid текущего элемента как текст
		sguid = ac_getstrvalue();

		// ТУТ НАДО ПОЛУЧИТЬ ЗОНЫ ТЕКУЩЕЙ ДВЕРИ

		curObjFromGUID = getLinkedZoneGUID (sguid, true); // получить айди зоны: true откуда ведет дверь, false — куда ведёт дверь
		curObjToGUID = getLinkedZoneGUID (sguid, false);  // получить айди зоны: true откуда ведет дверь, false — куда ведёт дверь

		// ВЫЯСНЯЕМ ИМЯ И ПРОЧИЕ ПАРАМЕТРЫ ПОЛУЧЕНЫХ ЗОН

		if (curObjToGUID == "") {
			cout << "   Зона _откуда не привязана.\n";
		}
		if (curObjFromGUID == "") {
			cout << "   Зона _куда не привязана.\n";
		}
		curObjToNumber = "";
		curObjToName = "";
		curObjToCat = "";
		curObjToCatName = "";
		curObjFromNumber = "";
		curObjFromName = "";
		curObjFromCat = "";
		curObjFromCatName = "";

		if (curObjToGUID != "")
		{
			ac_request("set_element_by_guidstr_as_current", curObjToGUID);
			ac_request("get_element_value", "ZoneNumber");
			curObjToNumber = ac_getstrvalue();
			ac_request("get_element_value", "ZoneName");
			curObjToName = ac_getstrvalue();
			ac_request("get_element_value", "ZoneCatCode");
			curObjToCat = ac_getstrvalue();
			ac_request("get_element_value", "ZoneCatName");
			curObjToCatName = ac_getstrvalue();

			cout << "    ZONE TO GUID =" << curObjToGUID << ".\n";
			cout << "    ZONE TO NAME=" << curObjToName << ".\n";
			cout << "    ZONE TO CATEGORY=" << curObjToCatName << ".\n";
		}

		if (curObjFromGUID != "")
		{
			ac_request("set_element_by_guidstr_as_current", curObjFromGUID);
			ac_request("get_element_value", "ZoneNumber");
			curObjFromNumber = ac_getstrvalue();
			ac_request("get_element_value", "ZoneName");
			curObjFromName = ac_getstrvalue();
			ac_request("get_element_value", "ZoneCatCode");
			curObjFromCat = ac_getstrvalue();
			ac_request("get_element_value", "ZoneCatName");
			curObjFromCatName = ac_getstrvalue();

			cout << "    ZONE FROM GUID=" << curObjFromGUID << ".\n";
			cout << "    ZONE FROM NAME=" << curObjFromName << ".\n";
			cout << "    ZONE FROM CATEGORY=" << curObjFromCatName << ".\n";
			cout << "    ZONE FROM CATEGORY N=" << curObjFromCat << ".\n";
		}




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

		string 	curTableFromCatName, curTableToCatName,
		        curTableFromName, curTableToName,
		        curDoorCategory;

		bool doorWithoutCategory = true;


		ac_request("set_current_element_from_list", 1, i);
		cout << "	Цикл записи параметров из таблицы...\n";

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

		if (curObjFromGUID != "")
		{
			ires = ac_request("elem_user_property", "set", "ZONE_FROM_NAME", curObjFromName);
			// cout << "       имя зоны > " << ires << "\n";
			ires = ac_request("elem_user_property", "set", "ZONE_FROM_CATEGORY", curObjFromCatName);
			// cout << "       категорию зоны >" << ires << "\n";
			ires = ac_request("elem_user_property", "set", "ZONE_FROM_GUID", curObjFromGUID);
			// cout << "       GUID зоны откуда >" << ires << "\n";
		} else {
			ires = ac_request("elem_user_property", "set", "ZONE_FROM_NAME", "");
			ires = ac_request("elem_user_property", "set", "ZONE_FROM_CATEGORY", "");
			ires = ac_request("elem_user_property", "set", "ZONE_FROM_GUID", "");
		}

		if (curObjToGUID != "")
		{
			ires = ac_request("elem_user_property", "set", "ZONE_TO_NAME", curObjToName);
			// cout << "       зоны < " << ires << "\n";
			ires = ac_request("elem_user_property", "set", "ZONE_TO_CATEGORY", curObjToCatName);
			// cout << "       категорию зоны < " << ires << "\n";
			ires = ac_request("elem_user_property", "set", "ZONE_TO_GUID", curObjToGUID);
			// cout << "       GUID зоны < " << ires << "\n";
		} else {
			ires = ac_request("elem_user_property", "set", "ZONE_TO_NAME", "");
			ires = ac_request("elem_user_property", "set", "ZONE_TO_CATEGORY", "");
			ires = ac_request("elem_user_property", "set", "ZONE_TO_GUID", "");
		}


		curObjFromCatName = tolower(curObjFromCatName);
		curObjFromName = tolower(curObjFromName);
		curObjToCatName = tolower(curObjToCatName);
		curObjToName = tolower(curObjToName);

		string currentObjectFromToCatName, currentTableFromToCatName;
		string prop1, prop2, prop3, prop4;


		prop1 = prop2 = prop3 = prop4 = "";

		currentObjectFromToCatName = currentTableFromToCatName ="";


		// .########..########...#######..########..########.########..########.####.########..######.
		// .##.....##.##.....##.##.....##.##.....##.##.......##.....##....##.....##..##.......##....##
		// .##.....##.##.....##.##.....##.##.....##.##.......##.....##....##.....##..##.......##......
		// .########..########..##.....##.########..######...########.....##.....##..######....######.
		// .##........##...##...##.....##.##........##.......##...##......##.....##..##.............##
		// .##........##....##..##.....##.##........##.......##....##.....##.....##..##.......##....##
		// .##........##.....##..#######..##........########.##.....##....##....####.########..######.

		/*

		ires = ac_request("elem_user_property", "get", "Помещение"); // значение в объекте
		if (ires!=0){cout<<"Свойство не найдено!\n"; return;};
		prop1 = ac_getstrvalue(); // получили значение свойства объекта

		*/

		ires = ac_request("elem_user_property", "get", "Производитель"); // значение в объекте
		if (ires!=0){cout<<"Свойство не найдено!\n"; return;};
		prop2 = ac_getstrvalue(); // получили значение свойства объекта

		ires = ac_request("elem_user_property", "get", "Заполнение"); // значение в объекте
		if (ires!=0){cout<<"Свойство не найдено!\n"; return;};
		prop3 = ac_getstrvalue(); // получили значение свойства объекта


		ires = ac_request("elem_user_property", "get", "Замок"); // значение в объекте
		if (ires!=0){cout<<"Свойство не найдено!\n"; return;};
		prop4 = ac_getstrvalue(); // получили значение свойства объекта

		// if (prop1 == " ") {prop1 = "";}
		if (prop2 == " ") {prop2 = "";}
		if (prop3 == " ") {prop3 = "";}
		if (prop4 == " ") {prop4 = "";}

		// currentObjectFromToCatName = currentObjectFromToCatName + "_|_" + prop1 + "_|_" + prop2 + "_|_" + prop3 + "_|_" + prop4;
		// currentObjectFromToCatName = currentObjectFromToCatName + "_|_" + prop2 + "_|_" + prop3 + "_|_" + prop4;

		cout << "			Строк в таблице: " << tableRowsNumber << "\n";
		cout << "			Колонок: " << tableColsNumber << "\n";

		//  поехал таймер тут, унц-унц

		codemeter(0);

		if ((curObjFromGUID != "") || (curObjToGUID != "")) {

			for (int row = 0; row < tableRowsNumber; row++) // cycle through all table rows
			{
				// cout << "		Row: " << row + 1 << " / " << tableRowsNumber << "\n";
				ts_table(iTableGUIDs, "select_row", row); // set the row

				curTableFromCatName = ""; // сбрасываем переменные, чтобы не думать о возможных проблемах с повторами
				curTableFromName = "";
				curTableToCatName = "";
				curTableToName = "";
				curDoorCategory = "";

				string col1, col2, col3, col4;
				col1 = col2 = col3 = col4 = "";

				// последовательные вызовы очень медленные. надо формировать 1 запрос. но это дико неудобно.

				ts_table(iTableGUIDs, "get_value_of",
				         0, curTableFromCatName,
				         1, curTableFromName,
				         2, curTableToCatName,
				         3, curTableToName,
				         4, curDoorCategory); //, /* ТУТ СТОИТ ПРОВЕРИТЬ КОЛИЧЕСТВО ДОП КОЛОНОК, ПОСКОЛЬКУ ВАРИАНТ С АВТОПОДСЧЕТОМ НЕ ИСПОЛЬЗУЕТСЯ, см. ниже*/
				      /* 5, col1,
				         6, col2,
				         7, col3,
				         8, col4); */


				curTableFromCatName = tolower(curTableFromCatName);
				curTableFromName = tolower(curTableFromName);
				curTableToCatName = tolower(curTableToCatName);
				curTableToName = tolower(curTableToName);
				curDoorCategory = tolower(curDoorCategory);

				col1 = tolower(col1);
				col2 = tolower(col2);
				col3 = tolower(col3);
				col4 = tolower(col4);


				// ..######..########.########
				// .##....##.##..........##...
				// .##.......##..........##...
				// ..######..######......##...
				// .......##.##..........##...
				// .##....##.##..........##...
				// ..######..########....##...

				// таблица
				// curTableFromName, curTableFromCatName, curTableToName, curTableToCatName — это параметры из текущей строки таблицы
				// объект
				// curObjFromName, curObjFromCatName, curObjToName, curObjToCatName — это параметры текущего объекта

				if (curTableFromCatName == "*") {curTableFromCatName = "";}
				if (curTableFromName == "*") {curTableFromName = "";}
				if (curTableToCatName == "*") {curTableToCatName = "";}
				if (curTableToName == "*") {curTableToName = "";}


				//  TABLE - здесь поперепутаны входы и выходы по причине того, что ассоциировали не верно, недоразумение вышло. Или вошло.
				currentTableFromToCatName = "_|_" + curTableFromCatName + "_|_" + curTableFromName + "_|_" + curTableToCatName + "_|_" + curTableToName + "_|_";

				// 	OBJECT
				currentObjectFromToCatName = "_|_" + curObjToCatName + "_|_" + curObjToName + "_|_" + curObjFromCatName + "_|_" + curObjFromName + "_|_";

				//---------------------------------------
				//
				// ПРОХОД ПО ДОПОЛНИТЕЛЬНЫМ ПОЛЯМ
				//
				//---------------------------------------


				// ЭТО ОЧЕНЬ МЕДЛЕННО РАБОТАЕТ!!!!!

				// ЭТО ОЧЕНЬ МЕДЛЕННООО!!!!!
				// ЗАПРОСЫ ПРИДЕТСЯ БЛОКИРОВАТЬ В ОДИН ВЫЗОВ >  СТАТИЧЕСКОЕ РЕШЕНИЕ БУДЕТ ПОЭТОМУ. НО ЕСЛИ ЮРИЙ ОПТИМИЗИРУЕТ

				// придется ПРАВИТЬ НАСТРОЙКИ ДЛЯ КАЖДОГО НОВОГО ПРОЕКТА!
				// количество дополнительных колонок для прохода


				// .########....###....########..##.......########
				// ....##......##.##...##.....##.##.......##......
				// ....##.....##...##..##.....##.##.......##......
				// ....##....##.....##.########..##.......######..
				// ....##....#########.##.....##.##.......##......
				// ....##....##.....##.##.....##.##.......##......
				// ....##....##.....##.########..########.########

				// .########.##.....##.########.########.....###...
				// .##........##...##.....##....##.....##...##.##..
				// .##.........##.##......##....##.....##..##...##.
				// .######......###.......##....########..##.....##
				// .##.........##.##......##....##...##...#########
				// .##........##...##.....##....##....##..##.....##
				// .########.##.....##....##....##.....##.##.....##

				// ..######...#######..##.......##.....##.##.....##.##....##..######.
				// .##....##.##.....##.##.......##.....##.###...###.###...##.##....##
				// .##.......##.....##.##.......##.....##.####.####.####..##.##......
				// .##.......##.....##.##.......##.....##.##.###.##.##.##.##..######.
				// .##.......##.....##.##.......##.....##.##.....##.##..####.......##
				// .##....##.##.....##.##.......##.....##.##.....##.##...###.##....##
				// ..######...#######..########..#######..##.....##.##....##..######.


				// 4 ПАРАМЕТРА ИЗ 4 ДОПОЛНИТЕЛЬНЫХ КОЛОНОК:

				if (col1 == "*") {col1 = "";}
				if (col2 == "*") {col2 = "";}
				if (col3 == "*") {col3 = "";}
				if (col4 == "*") {col4 = "";}

				// currentTableFromToCatName = currentTableFromToCatName + "_|_" + col1 + "_|_" + col2 + "_|_" + col3 + "_|_" + col4;
				// currentTableFromToCatName = currentTableFromToCatName + "_|_" + col2 + "_|_" + col3 + "_|_" + col4;

				// ВАРИАНТ ДЛЯ ТОГО СЛУЧАЯ, КОГДА ЮРИЙ ОПТИМИЗИРУЕТ РАБОТУ ТАБЛИЦ
				/*
				string currentExtraColumnName;
				string currentExtraColumnValue;
				string curObjExtraPropertyValue;

				for (int i = tableExtraColumnsNumber + 1; i < tableColsNumber; i++) { // ИДЁМ ПО ДОПОЛНИТЕЛЬНЫМ КОЛОНКАМ
					// cout << i << " ";
					ts_table(iTableGUIDs, "get_heading_of", i, currentExtraColumnName); // имя свойства
					ts_table(iTableGUIDs, "get_value_of", i, currentExtraColumnValue); // значение в колонке

					ires = ac_request("elem_user_property", "get", currentExtraColumnName); // значение в объекте
					curObjExtraPropertyValue = ac_getstrvalue(); // получили значение свойства объекта

					if (currentExtraColumnValue == "*") {
						currentExtraColumnValue = "";
					}

					if (curObjExtraPropertyValue == " ") {
						curObjExtraPropertyValue = "";
					}
					currentTableFromToCatName = currentTableFromToCatName + "_|_" + currentExtraColumnValue + "_|_";
					currentObjectFromToCatName = currentObjectFromToCatName + "_|_" + curObjExtraPropertyValue + "_|_";

				}
				*/

				//---------------------------------------
				//
				// ЗАДАТЬ СВОЙСТВО ДВЕРИ
				//
				//---------------------------------------

				// ЕСЛИ ВСЕ ПАРАМЕТРЫ СОВПАДАЮТ, ТО ЗАПИСАТЬ В ДВЕРЬ КАТЕГОРИЮ ДВЕРИ




				if (currentObjectFromToCatName == currentTableFromToCatName) {
					ires = ac_request("elem_user_property", "set", "DOOR_CATEGORY", curDoorCategory);
					cout << "	Записываем категорию двери: " << curDoorCategory << " статус: " << ires << "\n";
					doorWithoutCategory = false;
                                        // cout << "    Далее объект / табица значения для сравнения \n";
                                        // cout << currentObjectFromToCatName << "\n";
                                        // cout << currentTableFromToCatName << "\n";
                                        row = tableRowsNumber;
				} else {
					// cout << " Совпадений не найдено. \n";
				}

			} // end of table rows for loop
			cout << "Цикл в " << tableRowsNumber << " проходов выполнен за " << codemeter(2)/1000 << " сек. \n";
		} else { cout << "	Пропускаем эту дверь.\n"; }

		if (doorWithoutCategory == true) {
			ires = ac_request("elem_user_property", "set", "DOOR_CATEGORY", "—");
			cout << "	Записываем ПУСТУЮ категорию двери: " << ires << "\n";
		}

	} // end of doors for loop


	cout << "\n";
	deleteTable(); // очищаем память от таблицы
	cout << "Программа отработала успешно\n";
}



// .########.##.....##.##....##..######..########.####..#######..##....##..######.
// .##.......##.....##.###...##.##....##....##.....##..##.....##.###...##.##....##
// .##.......##.....##.####..##.##..........##.....##..##.....##.####..##.##......
// .######...##.....##.##.##.##.##..........##.....##..##.....##.##.##.##..######.
// .##.......##.....##.##..####.##..........##.....##..##.....##.##..####.......##
// .##.......##.....##.##...###.##....##....##.....##..##.....##.##...###.##....##
// .##........#######..##....##..######.....##....####..#######..##....##..######.


int LoadCSV(int iTable, string filepath, string column_separator) {
	int iFileDescr;

	object("create", "ts_file", iFileDescr); // создать объект типа файл в памяти

// открыть для записи чистый файл

	int ires = ts_file(iFileDescr, "open", filepath, "create", "r");
	if (ires != 0)
	{
		cout << "Файл не удалось открыть:" << filepath << "\n"; // выдать в окно сообщений
		return;
	} else {
		cout << "Файл открыли.\n";
	}

	string csv_source;
	ires = ts_file(iFileDescr, "read", csv_source); // записать в файл две строки

	if (ires != 0)
	{
		cout << "Не удалось считать csv из файла.\n";
		return;
	} else {
		cout << "Считали CSV из файла.\n";
	}
	ires = ts_file(iFileDescr, "close"); // закрыть файл
	if (ires != 0)
	{
		cout << "Не удалось закрыть файл.\n";
		return;
	} else {
		cout << "Файл закрыли.\n";
	}

	object("delete", iFileDescr); // удалить объект файла из памяти
	cout << "разделитель: " << column_separator << "\n";

	ts_table(iTable, "import_from_csv", csv_source, column_separator);

	return 0;
}


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

	//	runtimecontrol("workline", "setpos", 0);

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
	string sguid_linked_zone_guid = "";

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
		ts_table(iTableRoomsTmp, "get_value_of", "sGUID", sguid_linked_zone_guid);
	}
	object("delete", iGuid);
	object("delete", iTableRoomsTmp);
	return sguid_linked_zone_guid;
}

int createTable() {
	string sguid; 	// переменная для guid
	string svalue; 	// для строковых
	int ivalue; 	// для числовых
	int jcount;
	int icount;
	int ires;
	object("create", "ts_table", iTableGUIDs); // создаем объект ts_table и получаем его номер для работы с ним
	// ires = LoadExcel(); // загружаем данные из эксель (см. определение функции ниже)
	//
	ires = LoadCSV(iTableGUIDs, sCSVFilepath, ";");
	// int LoadCSV(int iTableGUID, string column_separator, string filepath){}
	//
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


int checkAndCreateUserParameters() {
	// СЮДА НАДО СРАЗУ ВОТКНУТЬ СОЗДАНИЕ ПАРАМЕТРОВ В КЛАССИФИКАТОРЕ, ЕСЛИ ИХ НЕТ

	// ПО ИДЕЕ, ЭТО НАДО ПЕРЕПИСАТЬ ЦИКЛАМИ ПО МАССИВУ, НО ПОТОМ

	// DOOR_CATEGORY — категория двери

	// ZONE_FROM_NAME — имя зоны, откуда ведет дверь
	// ZONE_FROM_CATEGORY — категория зоны, откуда ведет дверь
	// ZONE_FROM_GUID — GUID зоны, откуда ведет дверь

	// ZONE_TO_NAME — имя зоны, куда ведет дверь
	// ZONE_TO_CATEGORY — категория зоны, куда ведет дверь
	// ZONE_TO_GUID — GUID зоны, куда ведет дверь

	cout << "\nПроверяем набор параметров в группе OPENINGS.\n";

	ac_request("set_current_element_from_list", 1, 0);
	int r0, r1, r2, r3, r4, r5, r6, ires;
	bool the_res = true;

	r0 = ac_request_special("get_element_value", "UP", "DOOR_CATEGORY");
	r1 = ac_request_special("get_element_value", "UP", "ZONE_FROM_NAME");
	r2 = ac_request_special("get_element_value", "UP", "ZONE_FROM_CATEGORY");
	r3 = ac_request_special("get_element_value", "UP", "ZONE_FROM_GUID");
	r4 = ac_request_special("get_element_value", "UP", "ZONE_TO_NAME");
	r5 = ac_request_special("get_element_value", "UP", "ZONE_TO_CATEGORY");
	r6 = ac_request_special("get_element_value", "UP", "ZONE_TO_GUID");

	if (r0 == -1222) { 	//-1222 — ошибка какая-то
		cout << "Одно из свойств не найдено, надо создать.\n";
		ires = ac_request("elem_user_property", "create", "DOOR_CATEGORY", " ", "String", "OPENINGS");
		if (ires == 0) { cout << "Создали свойство DOOR_CATEGORY.\n"; } else { cout << "Не вышло создать свойство DOOR_CATEGORY, создайте вручную.\n"; break();}
	}

	if (r1 == -1222) { 	//-1222 — ошибка какая-то
		cout << "Одно из свойств не найдено, надо создать.\n";
		ires = ac_request("elem_user_property", "create", "ZONE_FROM_NAME", " ", "String", "OPENINGS");
		if (ires == 0) { cout << "Создали свойство ZONE_FROM_CATEGORY.\n"; } else { cout << "Не вышло создать свойство ZONE_FROM_CATEGORY, создайте вручную.\n"; break();}
	}

	if (r2 == -1222) { 	//-1222 — ошибка какая-то
		cout << "Одно из свойств не найдено, надо создать.\n";
		ires = ac_request("elem_user_property", "create", "ZONE_FROM_CATEGORY", " ", "String", "OPENINGS");
		if (ires == 0) { cout << "Создали свойство ZONE_FROM_CATEGORY.\n"; } else { cout << "Не вышло создать свойство ZONE_FROM_CATEGORY, создайте вручную.\n"; break();}
	}

	if (r3 == -1222) { 	//-1222 — ошибка какая-то
		cout << "Одно из свойств не найдено, надо создать.\n";
		ires = ac_request("elem_user_property", "create", "ZONE_FROM_GUID", " ", "String", "OPENINGS");
		if (ires == 0) { cout << "Создали свойство ZONE_FROM_GUID.\n"; } else { cout << "Не вышло создать свойство ZONE_FROM_GUID, создайте вручную.\n"; break();}
	}

	if (r4 == -1222) { 	//-1222 — ошибка какая-то
		cout << "Одно из свойств не найдено, надо создать.\n";
		ires = ac_request("elem_user_property", "create", "ZONE_TO_NAME", " ", "String", "OPENINGS");
		if (ires == 0) { cout << "Создали свойство ZONE_TO_NAME.\n"; } else { cout << "Не вышло создать свойство ZONE_TO_NAME, создайте вручную.\n"; break();}
	}

	if (r5 == -1222) { 	//-1222 — ошибка какая-то
		cout << "Одно из свойств не найдено, надо создать.\n";
		ires = ac_request("elem_user_property", "create", "ZONE_TO_CATEGORY", " ", "String", "OPENINGS");
		if (ires == 0) { cout << "Создали свойство ZONE_TO_CATEGORY.\n"; } else { cout << "Не вышло создать свойство ZONE_TO_CATEGORY, создайте вручную.\n"; break();}
	}

	if (r6 == -1222) { 	//-1222 — ошибка какая-то
		cout << "Одно из свойств не найдено, надо создать.\n";
		ires = ac_request("elem_user_property", "create", "ZONE_TO_GUID", " ", "String", "OPENINGS");
		if (ires == 0) { cout << "Создали свойство ZONE_TO_GUID.\n"; } else { cout << "Не вышло создать свойство ZONE_TO_GUID, создайте вручную.\n";  break();}
	}

	cout << "\n";
	return 0;
}
