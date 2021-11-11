int main()
{
	int iret;

	int w = 40;
	int sx = 1;
	int offsetx = 1;
	int offsety = 1;
	int iPos=0;
	int sy = 1;
	int h = 40;

    ac_request("create_iconbutton","CALC_ZONE256.png",sx,sy,sx+w,sy+h,"Set Door and Zones Connections","attachDoorsToZones.cpp");

	sx = sx + w + offsetx;
	ac_request("create_iconbutton","SELECT.png",sx,sy,sx+w,sy+h,"Set Selected Doors Parameters","setDoorsPropertiesFromTable_v1_0.cpp");

	sx = sx + w + offsetx;
	ac_request("create_iconbutton","Excel_to.png",sx,sy,sx+w,sy+h,"Save Doors Parameters List to CSV table","saveDoorsPropertiesToCSV_v1_0.cpp");

	sx = sx + w + offsetx;
	ac_request("create_iconbutton","CALC_LAMP256.png",sx,sy,sx+w,sy+h,"Check properties","checkProperties_v0.1.cpp");

	sy = sy + h + offsety;
	sx = 1;
	ac_request("set_palette_size_and_message_place",80,100,405,300,sx,sy,326-sx*2,200-sy);


	cout << "Программа назначения типов дверей в соответствии с типами прилежащих зон.";
	
	cout << "\n";
	
	cout << "Инструкция и последняя версия: https://github.com/dduucckk/labppSetDoorsPropertiesFromTable";
	
	cout << "\n";
	cout << "\n";
	
	cout << "1: Связать двери с зонами;";
	cout << "\n";

	cout << "2: Задать выбранным дверям тип из таблицы;";
	cout << "\n";

	cout << "3: Если таблицы нет, сгенерировать таблицу уникальных комбинаций зон из свойств выбранных дверей;";
	cout << "\n";

	cout << "4: Поиск не заданных/ошибочных параметров в выбранных дверях;";
	cout << "\n";

	cout << "\n";
	cout << "\n";
	cout << "Для работы с 2 и 4 у вас должен быть открыт Excel с таблицей ЗОНЫ И ДВЕРИ.xlsx.";
	cout << "\n";
	cout << "Для подсветки ошибок используйте графическую замену для свойства Wrong Category Type = true,\ncм. параметры двери."; cout << "\n";
	cout << "Если какие-то параметры не созданы в проекте, программа постарается их создать."




}
