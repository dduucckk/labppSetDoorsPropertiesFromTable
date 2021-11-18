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

	int numberOfTypes = 30;

	// ac_request("create_iconbutton","CALC_ZONE256.png",sx,sy,sx+w,sy+h,"Set Door and Zones Connections","attachDoorsToZones.cpp");

	for (int i = 1; i <= numberOfTypes; i++) {
		ac_request("create_iconbutton", "icon_0" + itoa(i) + ".png", sx, sy, sx + w, sy + h, "Set Selected Doors Parameters", "simplySetDoorType_v1_0.cpp", 0, 0.0, itoa(i));
		sx = sx + w + offsetx;
		if ((i==10) || (i==20)) {
			sy = sy + h + offsety;
			sx=1;
		}
	}


	sy = sy + h + offsety;
	ac_request("set_palette_size_and_message_place", 0, 0, sx+1, sy*2, 2, sy+4, sx-4, sy);

	// тут очень странно,  вообще не понимаю, как это работает.
	// x — по-горизонтали, y — по-вертикали. но работает очень странно.

	cout << "Программа задает тип двери, который вы выбрали.";
	cout << "\n";
	cout << "Инструкция и последняя версия: https://github.com/dduucckk/labppSetDoorsPropertiesFromTable";

}