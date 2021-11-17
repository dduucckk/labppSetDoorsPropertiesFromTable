int main()
{
	int iret;
	int w = 40;
	int sx = 1;
	int offsetx = 1;
	int offsety = 1;
	int iPos = 0;
	int sy = 1;
	int h = 40;

	int numberOfTypes = 10;

	// ac_request("create_iconbutton","CALC_ZONE256.png",sx,sy,sx+w,sy+h,"Set Door and Zones Connections","attachDoorsToZones.cpp");

	for (int i = 1; i <= numberOfTypes; i++) {
		ac_request("create_iconbutton", "icon_0" + i + ".png", sx, sy, sx + w, sy + h, "Set Selected Doors Parameters", "simplySetDoorType_v1_0.cpp", "текстовый аргумент");
		sx = sx + w + offsetx;
	}


	sy = sy + h + offsety;
	sx = 1;
	ac_request("set_palette_size_and_message_place", 80, 100, 405, 300, sx, sy, 326 - sx * 2, 200 - sy);

	cout << "Программа задает тип двери, который вы выбрали.";
	cout << "\n";
	cout << "Инструкция и последняя версия: https://github.com/dduucckk/labppSetDoorsPropertiesFromTable";



}