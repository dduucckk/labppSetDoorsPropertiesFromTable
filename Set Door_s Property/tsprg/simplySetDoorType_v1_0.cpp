// Задать тип выбранным дверям, просто задать и всё.
// LABPP AUTOMAT скрипт
// Иван Матвеев
// 2021
//
//
//


int main()
{


	string theProperty = "DOOR_TYPE";

	string curDoorCategory;
	int iArg1;
	double dArg2;
	string sArg3;

	run_cpp("get_args", iArg1, dArg2, sArg3);


	curDoorCategory = sArg3;

	int ires;
	int icount;

	//--------------------------------------------------
	// ЗАГРУЖАЕМ ДВЕРИ (ВСЕ)
	//--------------------------------------------------

	// ЗАГРУЗИТЬ ДВЕРИ ИЗ СПИСКА ВЫБРАННЫХ ЭЛЕМЕНТОВ
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

	//--------------------------------------------------
	// ЦИКЛ ПО ДВЕРЯМ
	//--------------------------------------------------

	cout << "Проходим по выбранным дверям. \n";
	cout << "\n";

	for (int i = 0; i < icount; i++)
	{
		ac_request("set_current_element_from_list", 1, i); // установить текущим элемент из списка № 1 с индексом i

		ires = ac_request("elem_user_property", "set", theProperty, curDoorCategory);
		cout << "	Записываем тип двери: " << curDoorCategory;
		if (ires == 0) {cout << " + \n";} else {cout << " « тут не получилось записать. \n"; }


	} // end of doors for loop
	cout << "\n";
	cout << "Программа отработала успешно\n";
}