// ������ ��� ��������� ������
// 26.2021
// LabPP
//string TSModuleVersion = "26.10.2021 - ��������� ������";
string TSModuleVersion = "28.10.2021 - ������ ������� �������";
//-----------------------

int iDialogDescr; // ���������� �������
int iListBoxDoors, iTableDoors;    // �������� ��������� ������ � ��� ������
int iEditSearchListBoxDoors, iBarControlSearchListBoxDoors; // �������� ������ � ���������

int iNormalTab, iTabPage1, iTabPage2, iTabPage3;
int iProgressBar; 

int iButtonAttachDoorsToRoomAsExits, iButtonAttachDoorsToRoomAsEntrances, iButtonShowRoomExits, iButtonShowRoomEntrances, iButtonDetachDoorsFromRoom;


int iButtonLoadDoors;

int main()
{
	coutvar << TSModuleVersion;

	// pragma region - ������� ���������� ��� ��������� C++ ��� ��������� �������� ������� ����� ������ � ��������� ������������ ������
#pragma region ������� ������
	int x, y, w, h;
	object("create", "ts_dialog", iDialogDescr);
	ts_dialog(iDialogDescr, "init_dialog", "palette", 0, 0, 450, 400); // ������� ���� ������� ��� �������, �.�. �����������
	ts_dialog(iDialogDescr, "set_as_main_panel"); // ���� ��� �������, �� ��� ����������� ���� ����� ������ ����� ����������� ������ � ���� �����
	ts_dialog(iDialogDescr, "SetTitle", "����� �������� ��������� ����");

	// ������� ������ � ����������
	object("create", "ts_dialogcontrol", iNormalTab, "NT1");
	ts_dialogcontrol(iNormalTab, "init_control", "normaltab", iDialogDescr, 0, 0, 450, 340); 
	// ��������� ������� ����� �������� �� ������ ��������� ������� ��� �������� �����, ������, ������, �����
	// 0,0,1,1 - �������� ��� ��� ��������� ������ ���� ������� ����� ������� �� �����, ���� �� �����, ����� - �������� ����� �� ���������� � ��� - ����.
	// ����� ���� -1, ��� ����� �������, ��� ��� ���������� ������� �����.������� �������� � �������� �������.
	ts_dialogcontrol(iNormalTab, "SetAnchorToPanelResize", 0, 0, 1, 1); 

	// ��������� ��������
	ts_dialogcontrol(iNormalTab, "AppendItem");
	ts_dialogcontrol(iNormalTab, "AppendItem");
	ts_dialogcontrol(iNormalTab, "AppendItem");

	// ������� ��������
	ts_dialogcontrol(iNormalTab, "SetItemText", 1, "�����");
	ts_dialogcontrol(iNormalTab, "SetItemText", 2, "-");
	ts_dialogcontrol(iNormalTab, "SetItemText", 3, "-");

	// ������� �� ��������� ������ ��� ���������� ���������
	object("create", "ts_dialogcontrol", iTabPage1, "TP1");
	object("create", "ts_dialogcontrol", iTabPage2, "TP2");
	object("create", "ts_dialogcontrol", iTabPage3, "TP3");

	// ���������� ��� ������
	ts_dialogcontrol(iTabPage1, "init_control", "tabpage", iNormalTab, 0, 0, 450, 340, 1);
	ts_dialogcontrol(iTabPage1, "SetAnchorToPanelResize", 0, 0, 1, 1);
	ts_dialogcontrol(iTabPage2, "init_control", "tabpage", iNormalTab, 0, 0, 450, 340, 2);
	ts_dialogcontrol(iTabPage2, "SetAnchorToPanelResize", 0, 0, 1, 1);
	ts_dialogcontrol(iTabPage3, "init_control", "tabpage", iNormalTab, 0, 0, 450, 340, 3);
	ts_dialogcontrol(iTabPage3, "SetAnchorToPanelResize", 0, 0, 1, 1);
	// �������� ������� ��������
	ts_dialogcontrol(iNormalTab, "selectitem", 1);

	// ������ ������ -------------------------------------------------------------------------------

	// �������� ��� ��������
	object("create", "ts_dialogcontrol", iListBoxDoors, "iListBoxDoors");
	ts_dialogcontrol(iListBoxDoors, "init_control", "singlesellistbox", iTabPage1, 0, 0, 450, 290, 48, 20);
	ts_dialogcontrol(iListBoxDoors, "SetAnchorToPanelResize", 0, 0, 1, 1);
	ts_dialogcontrol(iListBoxDoors, "eventreaction", "Event_ListBoxDoubleClicked"); // ���������� � ������� ���������� �� ������� ������ 

	// ������� ��� ��������� ������ ������
	object("create", "ts_table", iTableDoors);
	ts_table(iTableDoors, "add_column", 0,  "string", "GUID �����");
	ts_table(iTableDoors, "add_column", -1,  "string", "ID �����");
	ts_table(iTableDoors, "add_column", -1,  "int", "������");
	ts_table(iTableDoors, "add_column", -1,  "int", "������");
	ts_table(iTableDoors, "add_column", -1,  "int", "������ �����");
	ts_table(iTableDoors, "add_column", -1,  "string", "�� ���� �");
	ts_table(iTableDoors, "add_column", -1,  "string", "��������� ���� \"��\"");
	ts_table(iTableDoors, "add_column", -1,  "string", "��� ����\"��\"");
	ts_table(iTableDoors, "add_column", -1,  "string", "GUID ���� \"��\"");
	ts_table(iTableDoors, "add_column", -1,  "string", "� ���� �");
	ts_table(iTableDoors, "add_column", -1,  "string", "GUID ���� \"�\"");
	ts_table(iTableDoors, "add_column", -1, "string", "��������� ���� \"�\"");
	ts_table(iTableDoors, "add_column", -1, "string", "��� ���� \"�\"");
	ts_table(iTableDoors, "add_column", -1, "string", "��� �����");         // ����� ������� ����� ������� -1 - ����� ������� ���� ����������� �����
	//ts_table(iTableDoors, "resetofffromexport");
	ts_table(iTableDoors, "export_to_dialogcontrol", iListBoxDoors, -1, -1);
	//ts_table(iTableDoors, "set_first_key", 0);

	int delta = 3;
	int yy = 295;
	x = 25; y = yy; w = 120; h = 20;

	object("create", "ts_dialogcontrol", iEditSearchListBoxDoors, "iEditSearchListBoxDoors");
	ts_dialogcontrol(iEditSearchListBoxDoors, "init_control", "textedit", iTabPage1, x, y, w, h);
	object("create", "ts_dialogcontrol", iBarControlSearchListBoxDoors, "iBarControlSearchListBoxDoors");
	ts_dialogcontrol(iBarControlSearchListBoxDoors, "init_control", "singlespin", iTabPage1, 120, 0, 18, 22);
	ts_dialogcontrol(iEditSearchListBoxDoors, "create_simple_searcher", iListBoxDoors, iBarControlSearchListBoxDoors);
	ts_dialogcontrol(iEditSearchListBoxDoors, "SetAnchorToPanelResize", 0, 1, 0, 0);

	x = x + w + delta+20; y = yy; w = 120; h = 20;
	object("create", "ts_dialogcontrol", iButtonAttachDoorsToRoomAsExits, "iButtonAttachDoorsToRoomAsExits");
	ts_dialogcontrol(iButtonAttachDoorsToRoomAsExits, "init_control", "button", iTabPage1, x, y, w, h);
	ts_dialogcontrol(iButtonAttachDoorsToRoomAsExits, "eventreaction", "Event_ButtonClicked");
	ts_dialogcontrol(iButtonAttachDoorsToRoomAsExits, "settext", "��������� ������");
	ts_dialogcontrol(iButtonAttachDoorsToRoomAsExits, "SetAnchorToPanelResize", 0, 1, 0, 0);
	ts_dialogcontrol(iButtonAttachDoorsToRoomAsExits, "SetToolTip", "��������� ��������� ����� ��� ������ �� ���������� ���������");

	x = x + w + delta; y = yy; w = 120; h = 20;
	object("create", "ts_dialogcontrol", iButtonAttachDoorsToRoomAsEntrances, "iButtonAttachDoorsToRoomAsEntrances");
	ts_dialogcontrol(iButtonAttachDoorsToRoomAsEntrances, "init_control", "button", iTabPage1, x, y, w, h);
	ts_dialogcontrol(iButtonAttachDoorsToRoomAsEntrances, "eventreaction", "Event_ButtonClicked");
	ts_dialogcontrol(iButtonAttachDoorsToRoomAsEntrances, "settext", "��������� �����");
	ts_dialogcontrol(iButtonAttachDoorsToRoomAsEntrances, "SetAnchorToPanelResize", 0, 1, 0, 0);
	ts_dialogcontrol(iButtonAttachDoorsToRoomAsEntrances, "SetToolTip", "��������� ��������� ����� ��� ����� � ��������� ���������");

	x = x + w + delta; y = yy; w = 120; h = 20;
	object("create", "ts_dialogcontrol", iButtonShowRoomExits, "iButtonShowRoomExits");
	ts_dialogcontrol(iButtonShowRoomExits, "init_control", "button", iTabPage1, x, y, w, h);
	ts_dialogcontrol(iButtonShowRoomExits, "eventreaction", "Event_ButtonClicked");
	ts_dialogcontrol(iButtonShowRoomExits, "settext", "�������� ������");
	ts_dialogcontrol(iButtonShowRoomExits, "SetAnchorToPanelResize", 0, 1, 0, 0);
	ts_dialogcontrol(iButtonShowRoomExits, "SetToolTip", "�������� � ������� ��� �����-������ � ��������� ����");

	x = x + w + delta; y = yy; w = 120; h = 20;
	object("create", "ts_dialogcontrol", iButtonShowRoomEntrances, "iButtonShowRoomEntrances");
	ts_dialogcontrol(iButtonShowRoomEntrances, "init_control", "button", iTabPage1, x, y, w, h);
	ts_dialogcontrol(iButtonShowRoomEntrances, "eventreaction", "Event_ButtonClicked");
	ts_dialogcontrol(iButtonShowRoomEntrances, "settext", "�������� �����");
	ts_dialogcontrol(iButtonShowRoomEntrances, "SetAnchorToPanelResize", 0, 1, 0, 0);
	ts_dialogcontrol(iButtonShowRoomEntrances, "SetToolTip", "�������� � ������� ��� �����-����� � ��������� ����");

	x = x + w + delta; y = yy; w = 120; h = 20;
	object("create", "ts_dialogcontrol", iButtonDetachDoorsFromRoom, "iButtonDetachDoorsFromRoom");
	ts_dialogcontrol(iButtonDetachDoorsFromRoom, "init_control", "button", iTabPage1, x, y, w, h);
	ts_dialogcontrol(iButtonDetachDoorsFromRoom, "eventreaction", "Event_ButtonClicked");
	ts_dialogcontrol(iButtonDetachDoorsFromRoom, "settext", "�������� �����");
	ts_dialogcontrol(iButtonDetachDoorsFromRoom, "SetAnchorToPanelResize", 0, 1, 0, 0);
	ts_dialogcontrol(iButtonDetachDoorsFromRoom, "SetToolTip", "�������� �������� ����� � ��������� ����");

	//-----------------------------------------------------------------------------------------------------------------
	delta = 3;
	yy = 374;
	x = delta / 2; y = yy; w = 80; h = 20;
	object("create", "ts_dialogcontrol", iButtonLoadDoors, "iButtonLoadDoors");
	ts_dialogcontrol(iButtonLoadDoors, "init_control", "button", iDialogDescr, x, y, w, h);
	ts_dialogcontrol(iButtonLoadDoors, "eventreaction", "Event_ButtonClicked");
	ts_dialogcontrol(iButtonLoadDoors, "settext", "��������");
	ts_dialogcontrol(iButtonLoadDoors, "SetAnchorToPanelResize", 0, 1, 0, 0);
	ts_dialogcontrol(iButtonLoadDoors, "SetToolTip", "�������� ������ �� �������");

	y = yy - 16; w = 415; h = 10;
	object("create", "ts_dialogcontrol", iProgressBar, "ProgressBar");
	ts_dialogcontrol(iProgressBar, "init_control", "progressbar", iDialogDescr, x, y, w, h);
	ts_dialogcontrol(iProgressBar, "SetMin", 0);
	ts_dialogcontrol(iProgressBar, "SetMax", 100);
	ts_dialogcontrol(iProgressBar, "SetValue", 0);
	ts_dialogcontrol(iProgressBar, "SetAnchorToPanelResize", 0, 1, 1, 0);

#pragma endregion 

	bool bres;
    ts_dialog(iDialogDescr, "invoke", bres);
    return 0;
}
// ���������� ������� ������ �� ������
int Event_ButtonClicked(int iDescr, string sDescr)
{
	if (sDescr == "iButtonAttachDoorsToRoomAsExits") {
		cout << sDescr << "\n";
		do_iButtonAttachDoorsToRoom(true, true);
	}
	else if (sDescr == "iButtonAttachDoorsToRoomAsEntrances") {
		cout << sDescr << "\n";
		do_iButtonAttachDoorsToRoom(true, false);
	}
	else if (sDescr == "iButtonShowRoomExits") {
		cout << sDescr << "\n";
		do_iButtonShowDoors(true);
	}
	else if (sDescr == "iButtonShowRoomEntrances") {
		cout << sDescr << "\n";
		do_iButtonShowDoors(false);
	}
	else if (sDescr == "iButtonDetachDoorsFromRoom") {
		do_iButtonDetachDoorsFromRoom();
	}
	else if (sDescr == "iButtonLoadDoors") {
		cout << sDescr << "\n";
		Load();
	}
}
// ���������� ������� ���������� �� ������� ������
int Event_ListBoxDoubleClicked(int iDescr, string sDescr)
{
	if (sDescr == "iListBoxDoors") { 
		cout << sDescr << "\n";
		Zoom(); 
	}
}
// ������ ���� �� ������

//--------------------------------------------------
// �������� ������
//--------------------------------------------------
int Load()
{
	int ires;

	cout << "�������� ������\n";
	ts_dialogcontrol(iListBoxDoors, "DeleteItem", 0); // ������� ��� �������� � ���������
	ts_table(iTableDoors, "clear_rows");              // �������� ������� ������ - ������ ������� ������, �������� ���������
	ts_dialogcontrol(iNormalTab,"SelectItem",1);      // ������� ������ �������� �� �������
	ac_request("clear_list", 1);                      // �������� ������ �1

	// ��������� �������� ������ � ������ �1
	ac_request_special("load_elements_list", 1, "DoorType", 2 + 1024);

	//ac_request_special("load_elements_list", 1, "DoorType", 2 + 1024,
	//	"", "Cls", "������������� �� ��������", "=", sUPClassifValue, "", "OR",
	//	"", "Cls", "������������� �� ��������", "=", sUPClassifValue2, "", "OR",
	//	"", "Cls", "������������� �� ��������", "=", sUPClassifValue3, "");

	// ��������� ���������� ��������� ��������� ������
	ac_request("get_loaded_elements_list_count", 1);
	int icount = ac_getnumvalue();
	coutvar << icount;

	if (icount == 0)
	{
		cout << "� ������� �� ������� ����� (�������� ������� ����)";
		return -1;
	}

	int i;

	string floorname;
	int floorindex;
    string sText, sguid, szoneguidFrom, szoneguidTo, sIDdoor, sZoneNumberFrom, sZoneNumberTo, sZoneCatFrom, sZoneCatTo, sZoneNameFrom, sZoneNameTo;

    ts_dialogcontrol(iProgressBar, "SetMax", icount);

	// ������� ������ ���� ts_guid
	int iGuid;
	object("create", "ts_guid", iGuid);
	int width, height;

	for (i = 0; i < icount; i++)
	{
		ts_dialogcontrol(iProgressBar, "SetValue", i);

		// ���������� ������� i-����� ������� ������ �1 ��� ��������� �� ������� (guid �����)
		ac_request("set_current_element_from_list", 1, i);

		ac_request("get_element_value", "ID");
		sIDdoor = ac_getstrvalue();

		ac_request("get_element_value", "StoreIndex"); // ������� ������ ����� � �����
		floorindex = ac_getnumvalue();
		ac_request("get_floor_name_by_floor_index", floorindex, floorname); // �������� ��� ����� �� ��� �������

		ac_request("get_element_value", "GuidAsText"); // ��������� guid �������� �������� ��� �����
		sguid = ac_getstrvalue();

		ac_request_special("get_element_value", "GDL", "A");
		width = ac_getnumvalue() * 1000;

		ac_request_special("get_element_value", "GDL", "B");
		height = ac_getnumvalue() * 1000;

		szoneguidFrom = get_linked_zone_guid(sguid, true);
		if (szoneguidFrom != "")
		{
			ac_request("set_element_by_guidstr_as_current", szoneguidFrom);
			ac_request("get_element_value", "ZoneNumber");
			sZoneNumberFrom = ac_getstrvalue();
			ac_request("get_element_value", "ZoneName");
			sZoneNameFrom = ac_getstrvalue();
			ac_request("get_element_value", "ZoneCatCode");
			sZoneCatFrom = ac_getstrvalue();
		}
		szoneguidTo = get_linked_zone_guid(sguid, false);
		if (szoneguidTo != "")
		{
			ac_request("set_element_by_guidstr_as_current", szoneguidTo);
			ac_request("get_element_value", "ZoneNumber");
			sZoneNumberTo = ac_getstrvalue();
			ac_request("get_element_value", "ZoneName");
			sZoneNameTo = ac_getstrvalue();
			ac_request("get_element_value", "ZoneCatCode");
			sZoneCatTo = ac_getstrvalue();
		}

		ts_table(iTableDoors, "add_row",
			"ID �����", sIDdoor,
			"GUID �����", sguid,
			"������ �����", floorindex,
			"������", width,
			"������", height,
			"��� �����", floorname,
			"�� ���� �", sZoneNumberFrom,
			"��������� ���� \"��\"", sZoneCatFrom,
			"��� ����\"��\"", sZoneNameFrom,
			"GUID ���� \"��\"",szoneguidFrom,
			"� ���� �", sZoneNumberTo,
			"��������� ���� \"�\"", sZoneCatTo,
			"��� ���� \"�\"", sZoneNameTo,
			"GUID ���� \"�\"", szoneguidTo);
	}
	ts_dialogcontrol(iProgressBar, "SetValue", 0);
	ts_table(iTableDoors, "sort", "������ �����", "������", "������");
	ts_table(iTableDoors, "export_to_dialogcontrol", iListBoxDoors, -1, -1);
	ts_dialogcontrol(iListBoxDoors, "RepaintBackgroundItemsByColumnValue", 3, 255, 255, 255, 247, 247, 247);
}
//--------------------------------------------------
// Zoom 
//--------------------------------------------------
int Zoom()
{
	int item, storeindex;
	string sguid;

	ts_dialogcontrol(iListBoxDoors, "GetSelectedItem", item);
	if(item == 0) {
		return -1;
	}
	
	ts_table(iTableDoors,"select_row",item-1);
	ts_table(iTableDoors,"get_value_of","GUID �����",sguid);
	ts_table(iTableDoors,"get_value_of","������ �����",storeindex);

	ac_request("set_element_by_guidstr_as_current", sguid);
	ac_request("Environment","Story_GoTo",storeindex);
	ac_request("clear_list",7);
    ac_request("store_current_element_to_list",7,-1);
    ac_request("select_elements_from_list",7,1);
	ac_request("Automate","ZoomToElements",7);
	return 0;
}

int do_iButtonAttachDoorsToRoom(bool bOn, bool bRoomFrom)
{
	cout << "attach floors to room";

	int flag1, flag2;
	if (bRoomFrom)
	{
		flag1 = 1024; flag2 = 2048;
	}
	else
	{
		flag1 = 4096; flag2 = 8192;
	}

	// ��������� ��������� ���� � ������ 1
	ac_request_special("load_elements_list_from_selection", 1, "ZoneType", 2);
	ac_request("get_loaded_elements_list_count", 1);
	int icount = ac_getnumvalue();
	if (icount == 0)
	{
		tsalert(-2, "����������", "�� ������� ����", "����� ��������� ��������� ������ ���� ���� �� ���� ����", "Ok");
		return -1;
	}

	coutvar << icount;

	// ��������� ��������� ����� � ������ 2
	ac_request_special("load_elements_list_from_selection", 2, "DoorType", 2);

	ac_request("get_loaded_elements_list_count", 2);
	int icount_doors = ac_getnumvalue();
	if (icount_doors == 0)
	{
		tsalert(-2, "����������", "�� ������� �����", "����� ��������� ��������� ������ ���� ���� �� ���� �����", "Ok");
		return -1;
	}
	coutvar << icount_doors;

	ac_request("set_current_element_from_list", 1, 0);
	ac_request("get_element_value", "GuidAsText");
	string sGUIDzone = ac_getstrvalue();
	coutvar << sGUIDzone;
	int iGuid;
	object("create", "ts_guid", iGuid);
	ts_guid(iGuid, "ConvertFromString", sGUIDzone);

	int iTableDoorsTmp;
	object("create", "ts_table", iTableDoorsTmp);
	ts_table(iTableDoorsTmp, "load_sguids_from_list", 2);
	ac_request_special("linkingElems", "uplinkBiWardByFlags", iGuid, flag1, bOn, flag2, bOn, iTableDoorsTmp);

	object("delete", iGuid);
	object("delete", iTableDoorsTmp);
}

int do_iButtonShowDoors(bool bRoomFrom)
{
	cout << "show doors";
	int flag1, flag2;
	if (bRoomFrom)
	{
		flag1 = 1024; flag2 = 2048;
	}
	else
	{
		flag1 = 4096; flag2 = 8192;
	}

	int ires;
	ac_request_special("load_elements_list_from_selection", 1, "ZoneType", 2);
	ac_request("get_loaded_elements_list_count", 1);
	int icount = ac_getnumvalue();
	if (icount == 0)
	{
		tsalert(-2, "����������", "�� ������� ����", "����� ��������� ��������� ������ ���� ����", "Ok");
		return -1;
	}
	coutvar << icount;

	int iGuid;
	object("create", "ts_guid", iGuid);
	ires = ac_request("set_current_element_from_list", 1, 0);
	ires = ac_request("get_element_value", "GuidAsText");
	string sGUID = ac_getstrvalue();
	ts_guid(iGuid, "ConvertFromString", sGUID);

	int iTableDoorsTmp;
	object("create", "ts_table", iTableDoorsTmp);
	ac_request_special("linkingElems", "getLinkedElemsByFlags", iGuid, flag1, iTableDoorsTmp);

	ac_request("clear_list", 2);
	ts_table(iTableDoorsTmp, "add_sguids_to_list", 2);

	ac_request("select_elements_from_list", 2, 1);
	object("delete", iGuid);
	object("delete", iTableDoorsTmp);
}

int do_iButtonDetachDoorsFromRoom()
{
	cout << "detach doors from rooms";

	// ��������� ��������� ���� � ������ 1
	ac_request_special("load_elements_list_from_selection", 1, "ZoneType", 2);
	ac_request("get_loaded_elements_list_count", 1);
	int icount = ac_getnumvalue();
	if (icount == 0)
	{
		tsalert(-2, "����������", "�� ������� ����", "����� ��������� ��������� ������ ���� ���� �� ���� ����", "Ok");
		return -1;
	}

	coutvar << icount;

	// ��������� ��������� ����� � ������ 2
	ac_request_special("load_elements_list_from_selection", 2, "DoorType", 2);

	ac_request("get_loaded_elements_list_count", 2);
	int icount_doors = ac_getnumvalue();
	if (icount_doors == 0)
	{
		tsalert(-2, "����������", "�� ������� �����", "����� ��������� ��������� ������ ���� ���� �� ���� �����", "Ok");
		return -1;
	}
	coutvar << icount_doors;

	ac_request("set_current_element_from_list", 1, 0);
	ac_request("get_element_value", "GuidAsText");
	string sGUIDzone = ac_getstrvalue();
	coutvar << sGUIDzone;
	int iGuid;
	object("create", "ts_guid", iGuid);
	ts_guid(iGuid, "ConvertFromString", sGUIDzone);

	int iTableDoorsTmp;
	object("create", "ts_table", iTableDoorsTmp);
	ts_table(iTableDoorsTmp, "load_sguids_from_list", 2);

	int flag1, flag2;

	bool bOff = false;
	flag1 = 1024; flag2 = 2048;
	ac_request_special("linkingElems", "uplinkBiWardByFlags", iGuid, flag1, bOff, flag2, bOff, iTableDoorsTmp);
	flag1 = 4096; flag2 = 8192;
	ac_request_special("linkingElems", "uplinkBiWardByFlags", iGuid, flag1, bOff, flag2, bOff, iTableDoorsTmp);

	object("delete", iGuid);
	object("delete", iTableDoorsTmp);
}

string get_linked_zone_guid(string sDoorGuid, bool bRoomFrom) // �������� ���� ����: ���� true � ������ ����� �����, ���� false � ���� ���� �����
{
	string sguid_linked_zone_guid = "";

	int flag1, flag2;
	if (bRoomFrom)
	{
		flag1 = 1024; flag2 = 2048;
	}
	else
	{
		flag1 = 4096; flag2 = 8192;
	}

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
